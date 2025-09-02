import os
import pandas as pd
import re
import requests
import zipfile
import subprocess

download_dir = r"C:/Users/hwfrz/Downloads"

# 병용금기 고시 파일 및 추가 zip 파일 압축 해제
seven_zip = r"C:/Program Files/7-Zip/7z.exe" # 7zip으로 해제해야 인코딩 문제 없음
# download.zip 해제
zip_path = os.path.join(download_dir, "download.zip")
if os.path.exists(zip_path):
    cmd = [seven_zip, 'x', '-y', f'-o{download_dir}', zip_path]
    subprocess.run(cmd, check=True)
    print(f"압축 해제 완료(7z): {zip_path} → {download_dir}")
    if os.path.exists(zip_path):
        os.remove(zip_path)
# 게시_병용금기 급여/비급여 품목리스트_*.zip 해제
import glob
for zip_file in glob.glob(os.path.join(download_dir, "게시_병용금기*품목리스트_*.zip")):
    cmd = [seven_zip, 'x', '-y', f'-o{download_dir}', zip_file]
    subprocess.run(cmd, check=True)
    print(f"압축 해제 완료(7z): {zip_file} → {download_dir}")
    if os.path.exists(zip_file):
        os.remove(zip_file)

drug_contra_pay_pattern = re.compile(r"게시_병용금기 급여 품목리스트_\d{4}\.xlsb$")
drug_contra_nonpay_pattern = re.compile(r"게시_병용금기 비급여 품목리스트_\d{4}\.xlsx$")

files = [
    os.path.join(download_dir, f)
    for f in os.listdir(download_dir)
    if drug_contra_pay_pattern.match(f) or drug_contra_nonpay_pattern.match(f)
]

if not files:
    raise Exception("해당 패턴의 파일이 없습니다.")

latest_file = max(files, key=os.path.getctime)

# 예제파일 포맷 컬럼명, pay_a pay_b는 추후 drop시켜야 함
example_columns = [
    "ingName_a", "ppMainCode_a", "ppCode_a", "proKorName_a", "pdKorName_a", "pay_a",
    "ingName_b", "ppMainCode_b", "ppCode_b", "proKorName_b", "pdKorName_b", "pay_b",
    "reference", "appliedDate", "remark", "etc"
]

# 두 파일의 모든 시트 데이터 합치기
all_dfs = []
for file in files:
    if file.endswith('.xlsb'):
        sheets = pd.read_excel(file, sheet_name=None, engine='pyxlsb')
    else:
        sheets = pd.read_excel(file, sheet_name=None)
    for sheet, df in sheets.items():
        df = df.iloc[:, :len(example_columns)]
        df.columns = example_columns[:df.shape[1]]
        # ppCode_a, ppCode_b는 항상 문자열로 변환 (0으로 시작하는 보험코드 보존)
        for col in ["ppCode_a", "ppCode_b"]:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x) if not pd.isna(x) else "")
        all_dfs.append(df)

if not all_dfs:
    raise Exception("병합할 데이터가 없습니다. 파일/시트 구조를 확인하세요.")

# 모든 데이터 병합
merged_df = pd.concat(all_dfs, ignore_index=True)

# 컬럼명 공백 제거 후 pay_a, pay_b 컬럼만 최종적으로 drop
merged_df.columns = merged_df.columns.str.strip()
merged_df = merged_df.drop(columns=[col for col in ["pay_a", "pay_b"] if col in merged_df.columns], errors='ignore')

output_path = os.path.join(download_dir, "(예제파일)작업용_병용금기_급여_품목리스트.xlsx")

if os.path.exists(output_path):
    os.remove(output_path)
    
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    merged_df.to_excel(writer, sheet_name="drug_contra_product_temp", index=False)
print(f"저장 완료: {output_path}")

