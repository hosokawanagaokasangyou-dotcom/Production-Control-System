from pathlib import Path

from openpyxl import Workbook


def make_konan_shisan(path: Path, ym_label: str, sheet_name: str = '東レT.V.C'):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws['A1'] = ym_label
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
