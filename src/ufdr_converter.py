import json
import re
import shutil
import subprocess
import tarfile
import zipfile
from collections import defaultdict
from pathlib import Path
import xml.etree.ElementTree as ET

from openpyxl import Workbook


NS = {"c": "http://pa.cellebrite.com/report/2.0"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".heic", ".heif", ".webp"}
VIDEO_EXTENSIONS = {".mp4", ".mov", ".avi", ".m4v", ".3gp"}
AUDIO_EXTENSIONS = {".ogg", ".mp3", ".m4a", ".wav", ".aac", ".amr"}
CAMERA_ROLL_HINTS = ("dcim", "pictures", "photos", "camera")
VOICE_MEMO_HINTS = ("voice", "memo", "recordings", "audio")

TOP_LEVEL_EXPORTS = {
    "SNS_Activities.xlsx": {"SocialMediaActivity", "AppActivity"},
    "Location_History.xlsx": {"Location", "Coordinate", "Journey", "Route", "GpsLocation"},
    "Browser_History.xlsx": {"BrowserHistory", "VisitedPage", "WebHistory", "BrowserBookmark"},
    "Notes.xlsx": {"Note", "Memo", "NoteEntry", "StickyNote"},
    "Contacts.xlsx": {"Contact"},
    "Call_History.xlsx": {"Call", "CallLog", "CallRecord"},
    "Emails.xlsx": {"Email"},
}


def normalize_name(value):
    """ファイル名やキー名に使える安全な文字列へ正規化する。"""
    text = (value or "").strip()
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
    return text.strip("._") or "unknown"


def text_or_none(node):
    """XMLノードの文字列を取り出し、空文字ならNoneを返す。"""
    if node is None:
        return None
    text = (node.text or "").strip()
    return text or None


def get_scalar_value(element):
    """UFDRのfield要素から単一値として扱える内容を抽出する。"""
    values = [text_or_none(node) for node in element.findall("c:value", NS)]
    values = [value for value in values if value is not None]
    if values:
        return " | ".join(values)
    if element.find("c:empty", NS) is not None:
        return None
    return text_or_none(element)


def shallow_model_summary(model):
    """子モデルを1行で表せる程度の浅い辞書に変換する。"""
    summary = {"ModelType": model.get("type"), "ModelId": model.get("id")}
    for field in model.findall("c:field", NS):
        name = field.get("name")
        if name:
            summary[name] = get_scalar_value(field)
    return {key: value for key, value in summary.items() if value not in (None, "")}


def model_to_record(model):
    """UFDRのmodel要素をExcel出力向けのフラットな辞書へ変換する。"""
    record = {
        "ModelType": model.get("type"),
        "ModelId": model.get("id"),
        "DeletedState": model.get("deleted_state"),
        "DecodingConfidence": model.get("decoding_confidence"),
        "ExtractionId": model.get("extractionId"),
    }

    for field in model.findall("c:field", NS):
        name = field.get("name")
        if name:
            record[name] = get_scalar_value(field)

    for multi_field in model.findall("c:multiField", NS):
        name = multi_field.get("name")
        if not name:
            continue
        values = [text_or_none(value) for value in multi_field.findall("c:value", NS)]
        values = [value for value in values if value is not None]
        record[name] = " | ".join(values) if values else None

    for model_field in model.findall("c:modelField", NS):
        name = model_field.get("name")
        child_model = model_field.find("c:model", NS)
        if not name:
            continue
        if child_model is None:
            record[name] = None
            continue
        for key, value in shallow_model_summary(child_model).items():
            record[f"{name}.{key}"] = value

    for multi_model_field in model.findall("c:multiModelField", NS):
        name = multi_model_field.get("name")
        child_models = multi_model_field.findall("c:model", NS)
        if not name:
            continue
        record[f"{name}.Count"] = len(child_models)
        summaries = [shallow_model_summary(child_model) for child_model in child_models]
        record[name] = json.dumps(summaries, ensure_ascii=False) if summaries else None

    return record


def ensure_directory(path):
    """指定ディレクトリが存在しなければ再帰的に作成する。"""
    path.mkdir(parents=True, exist_ok=True)


def reset_directory(path):
    """指定ディレクトリを作り直して前回生成物を消す。"""
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def write_records_to_xlsx(records, output_path, sheet_name="Data"):
    """辞書の配列を1シートのxlsxとして保存する。"""
    if not records:
        return False

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name[:31] or "Data"

    headers = sorted({key for record in records for key in record.keys()})
    worksheet.append(headers)
    for record in records:
        worksheet.append([record.get(header) for header in headers])

    ensure_directory(output_path.parent)
    workbook.save(output_path)
    return True


def extract_ufdr(ufdr_path, extract_dir):
    """UFDRアーカイブを指定ディレクトリへ展開する。"""
    print(f"Extracting {ufdr_path}...")
    try:
        with zipfile.ZipFile(ufdr_path, "r") as archive:
            archive.extractall(extract_dir)
    except zipfile.BadZipFile:
        try:
            with tarfile.open(ufdr_path, "r:*") as archive:
                archive.extractall(extract_dir)
        except tarfile.TarError:
            result = subprocess.run(
                ["tar", "-xf", str(ufdr_path), "-C", str(extract_dir)],
                capture_output=True,
                text=True,
                check=False,
            )
            if result.returncode != 0:
                raise ValueError(
                    f"Unsupported UFDR archive format: {ufdr_path}\n{result.stderr.strip()}"
                ) from None
    print("Extraction complete.")


def extract_phone_info(report_root):
    """レポート直下のメタデータから端末情報一覧を抽出する。"""
    rows = []
    for metadata in report_root.findall("c:metadata", NS):
        section = metadata.get("section")
        for item in metadata.findall("c:item", NS):
            rows.append(
                {
                    "Section": section,
                    "Name": item.get("name"),
                    "Group": item.get("group"),
                    "SourceExtraction": item.get("sourceExtraction"),
                    "Value": text_or_none(item),
                }
            )
    return rows


def build_tagged_file_index(report_root):
    """taggedFilesをID検索しやすい辞書へ変換する。"""
    tagged_files = {}
    for file_node in report_root.findall(".//c:taggedFiles/c:file", NS):
        file_id = file_node.get("id")
        metadata = {
            "FileId": file_id,
            "OriginalPath": file_node.get("path"),
            "Size": file_node.get("size"),
            "Deleted": file_node.get("deleted"),
            "ExtractionId": file_node.get("extractionId"),
        }

        for timestamp in file_node.findall("c:accessInfo/c:timestamp", NS):
            metadata[f"Timestamp.{timestamp.get('name')}"] = text_or_none(timestamp)

        for metadata_node in file_node.findall("c:metadata", NS):
            section = normalize_name(metadata_node.get("section") or "metadata")
            for item in metadata_node.findall("c:item", NS):
                item_name = normalize_name(item.get("name") or "item")
                key = f"{section}.{item_name}"
                value = text_or_none(item)
                if key in metadata and metadata[key]:
                    metadata[key] = f"{metadata[key]} | {value}"
                else:
                    metadata[key] = value

        local_path = metadata.get("File.Local_Path")
        metadata["ArchivePath"] = local_path.replace("\\", "/") if local_path else None
        tagged_files[file_id] = metadata

    return tagged_files


def copy_unique_file(source_path, destination_dir):
    """同名衝突を避けながらファイルをコピーする。"""
    ensure_directory(destination_dir)
    destination_path = destination_dir / source_path.name
    counter = 1
    while destination_path.exists():
        destination_path = destination_dir / f"{source_path.stem}_{counter}{source_path.suffix}"
        counter += 1
    shutil.copy2(source_path, destination_path)
    return destination_path


def resolve_extracted_file(extract_dir, tagged_file):
    """展開済みディレクトリからtagged fileの実ファイルを解決する。"""
    archive_path = tagged_file.get("ArchivePath")
    if not archive_path:
        return None
    candidate = extract_dir / archive_path
    return candidate if candidate.exists() else None


def infer_app_name(*values):
    """会話情報やドメイン情報からアプリ名らしい値を推定する。"""
    for value in values:
        if not value:
            continue
        text = str(value)
        if ":" in text:
            return normalize_name(text.split(":", 1)[0])
        if "AppDomain-" in text:
            return normalize_name(text.split("AppDomain-", 1)[1].split("|", 1)[0])
        return normalize_name(text)
    return "unknown_app"


def extract_chat_exports(chat_models, tagged_files, extract_dir, output_dir):
    """チャット一覧、メッセージ履歴、添付メディアを出力する。"""
    thread_rows = []
    message_rows = []
    attachment_ids = set()
    media_by_app = defaultdict(list)

    for chat_model in chat_models:
        thread_record = model_to_record(chat_model)
        thread_rows.append(thread_record)
        chat_source = thread_record.get("Source")

        for messages_field in chat_model.findall("c:multiModelField", NS):
            if messages_field.get("name") != "Messages":
                continue

            for message_model in messages_field.findall("c:model", NS):
                message_record = model_to_record(message_model)
                message_record["ChatSource"] = chat_source
                message_record["ChatModelId"] = chat_model.get("id")
                message_rows.append(message_record)

                for attachments_field in message_model.findall("c:multiModelField", NS):
                    if attachments_field.get("name") != "Attachments":
                        continue

                    for attachment_model in attachments_field.findall("c:model", NS):
                        attachment_record = model_to_record(attachment_model)
                        file_id = attachment_model.get("file_id")
                        if not file_id:
                            continue
                        attachment_ids.add(file_id)

                        tagged_file = tagged_files.get(file_id, {})
                        app_name = infer_app_name(
                            message_record.get("Source"),
                            chat_source,
                            tagged_file.get("MetaData.iPhone-Domain"),
                        )
                        media_dir = output_dir / "SNS_Chat_Media" / app_name
                        copied_path = None
                        source_path = resolve_extracted_file(extract_dir, tagged_file)
                        if source_path is not None:
                            copied_path = copy_unique_file(source_path, media_dir)

                        media_metadata = {f"Message.{k}": v for k, v in message_record.items()}
                        media_metadata.update({f"Attachment.{k}": v for k, v in attachment_record.items()})
                        media_metadata.update({f"TaggedFile.{k}": v for k, v in tagged_file.items()})
                        media_metadata["CopiedPath"] = str(copied_path) if copied_path else None
                        media_by_app[app_name].append(media_metadata)

    write_records_to_xlsx(thread_rows, output_dir / "Chat_Threads.xlsx", sheet_name="Chats")
    write_records_to_xlsx(message_rows, output_dir / "Chat_Messages.xlsx", sheet_name="Messages")

    for app_name, rows in media_by_app.items():
        write_records_to_xlsx(rows, output_dir / "SNS_Chat_Media" / app_name / "Metadata.xlsx", sheet_name="Media")

    return attachment_ids


def is_camera_roll_candidate(tagged_file, attachment_ids):
    """添付済み以外の画像や動画をカメラロール候補として判定する。"""
    if tagged_file.get("FileId") in attachment_ids:
        return False

    archive_path = (tagged_file.get("ArchivePath") or "").lower()
    original_path = (tagged_file.get("OriginalPath") or "").lower()
    extension = Path(archive_path or original_path).suffix.lower()
    if extension not in IMAGE_EXTENSIONS | VIDEO_EXTENSIONS:
        return False

    tags = (tagged_file.get("File.Tags") or "").lower()
    if any(hint in archive_path or hint in original_path for hint in CAMERA_ROLL_HINTS):
        return True
    return "image" in tags or "video" in tags or not archive_path.startswith("files/audio/")


def is_voice_memo_candidate(tagged_file, attachment_ids):
    """添付済み以外の音声ファイルをボイスメモ候補として判定する。"""
    if tagged_file.get("FileId") in attachment_ids:
        return False

    archive_path = (tagged_file.get("ArchivePath") or "").lower()
    original_path = (tagged_file.get("OriginalPath") or "").lower()
    extension = Path(archive_path or original_path).suffix.lower()
    tags = (tagged_file.get("File.Tags") or "").lower()

    if extension in AUDIO_EXTENSIONS:
        return True
    if "audio" in tags:
        return True
    return any(hint in archive_path or hint in original_path for hint in VOICE_MEMO_HINTS)


def export_tagged_file_group(tagged_files, extract_dir, output_dir, folder_name, selector, write_metadata):
    """条件に一致したtagged file群をフォルダへコピーし必要なら台帳も作る。"""
    group_dir = output_dir / folder_name
    ensure_directory(group_dir)
    metadata_rows = []
    copied_count = 0

    for tagged_file in tagged_files.values():
        if not selector(tagged_file):
            continue

        source_path = resolve_extracted_file(extract_dir, tagged_file)
        if source_path is None:
            continue

        destination_path = copy_unique_file(source_path, group_dir)
        copied_count += 1
        row = dict(tagged_file)
        row["CopiedPath"] = str(destination_path)
        metadata_rows.append(row)

    if write_metadata:
        write_records_to_xlsx(metadata_rows, group_dir / "Metadata.xlsx", sheet_name="Metadata")

    print(f"Copied {copied_count} files to {group_dir}")


def export_top_level_categories(top_level_models, output_dir):
    """トップレベルのモデルをカテゴリ別Excelへ書き出す。"""
    for file_name, model_types in TOP_LEVEL_EXPORTS.items():
        rows = []
        for model_type in model_types:
            for model in top_level_models.get(model_type, []):
                rows.append(model_to_record(model))
        write_records_to_xlsx(rows, output_dir / file_name, sheet_name=Path(file_name).stem[:31])


def parse_report(report_xml_path):
    """report.xmlを読み込んでXMLルート要素を返す。"""
    tree = ET.parse(report_xml_path)
    return tree.getroot()


def collect_top_level_models(report_root):
    """decodedData配下のトップレベルmodelを型ごとに集約する。"""
    top_level_models = defaultdict(list)
    for model_type_node in report_root.findall(".//c:decodedData/c:modelType", NS):
        type_name = model_type_node.get("type")
        for model in model_type_node.findall("c:model", NS):
            top_level_models[type_name].append(model)
    return top_level_models


def format_case_name(case_index):
    """ケース番号を4桁ゼロ埋め文字列へ変換する。"""
    return f"{case_index:04d}"


def process_ufdr(ufdr_path, output_root, temp_root, case_name):
    """1件のUFDRを展開し、各種成果物と展開済みテンポラリを作成する。"""
    case_output_dir = output_root / case_name
    extract_dir = temp_root / case_name
    reset_directory(case_output_dir)
    reset_directory(extract_dir)

    extract_ufdr(ufdr_path, extract_dir)

    report_xml_path = extract_dir / "report.xml"
    if not report_xml_path.exists():
        raise FileNotFoundError(f"report.xml not found in {ufdr_path}")

    report_root = parse_report(report_xml_path)
    top_level_models = collect_top_level_models(report_root)
    tagged_files = build_tagged_file_index(report_root)

    phone_info = extract_phone_info(report_root)
    write_records_to_xlsx(phone_info, case_output_dir / "Phone_Info.xlsx", sheet_name="PhoneInfo")

    attachment_ids = extract_chat_exports(top_level_models.get("Chat", []), tagged_files, extract_dir, case_output_dir)
    export_top_level_categories(top_level_models, case_output_dir)

    export_tagged_file_group(
        tagged_files,
        extract_dir,
        case_output_dir,
        "CameraRoll",
        lambda tagged_file: is_camera_roll_candidate(tagged_file, attachment_ids),
        write_metadata=True,
    )
    export_tagged_file_group(
        tagged_files,
        extract_dir,
        case_output_dir,
        "VoiceMemos",
        lambda tagged_file: is_voice_memo_candidate(tagged_file, attachment_ids),
        write_metadata=False,
    )


def find_input_files(input_dir):
    """入力ディレクトリ配下のUFDRファイルを列挙する。"""
    return sorted(path for path in input_dir.glob("*.ufdr") if path.is_file())


def main():
    """入力フォルダ内のUFDRを順番に処理して成果物を生成する。"""
    workspace_root = Path(__file__).resolve().parents[1]
    input_dir = workspace_root / "data" / "in"
    output_dir = workspace_root / "data" / "out"
    temp_dir = workspace_root / "data" / "temp"

    ensure_directory(output_dir)
    ensure_directory(temp_dir)
    ufdr_files = find_input_files(input_dir)
    if not ufdr_files:
        raise FileNotFoundError(f"No .ufdr files found in {input_dir}")

    for case_index, ufdr_path in enumerate(ufdr_files, start=1):
        case_name = format_case_name(case_index)
        print(f"Processing {ufdr_path.name} -> {case_name}")
        process_ufdr(ufdr_path, output_dir, temp_dir, case_name)

    print(f"Finished. Output written to {output_dir}")


if __name__ == "__main__":
    main()