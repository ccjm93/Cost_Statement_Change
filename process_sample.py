from pathlib import Path

from excel_change_processor import list_sheet_names, preferred_sheet_name, process_workbook


workspace = Path(__file__).resolve().parent
samples = sorted(
    path
    for path in workspace.glob("sample_*.xlsx")
    if not path.name.endswith("_설계변경자동화.xlsx") and not path.name.startswith("~$")
)

if not samples:
    raise SystemExit("처리할 sample_*.xlsx 파일을 찾지 못했습니다.")

sample = samples[0]
sheet_name = preferred_sheet_name(list_sheet_names(sample))
if not sheet_name:
    raise SystemExit("처리할 시트를 찾지 못했습니다.")

result = process_workbook(sample, sheet_name, insert_before_col=1)
print(f"output={result.output_path}")
print(f"sheet={result.sheet_name}")
print(f"processed_rows={result.processed_rows}")
print(f"validation={','.join(result.validation)}")
