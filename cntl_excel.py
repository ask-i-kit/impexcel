from pathlib import Path
from datetime import datetime

def list_excel_files_with_timestamps(directory_path):
    """
    指定ディレクトリ内のExcelファイルとその最終更新日時を
    タプルのリストで返却する関数

    Args:
        directory_path (str): 調査対象のディレクトリパス

    Returns:
        List[Tuple[str, str]]: (ファイルパス, 最終更新日時) のタプルのリスト
    """
    excel_extensions = ['.xls', '.xlsx', '.xlsm', '.xlsb']
    result = []
    directory = Path(directory_path)

    if not directory.is_dir():
        raise ValueError(f"指定されたパスはディレクトリではありません: {directory_path}")

    for file in directory.iterdir():
        if file.is_file() and file.suffix.lower() in excel_extensions:
            timestamp = datetime.fromtimestamp(file.stat().st_mtime)
            formatted_time = timestamp.strftime("%Y-%m-%d %H:%M:%S")
            result.append((str(file), formatted_time))

    return result