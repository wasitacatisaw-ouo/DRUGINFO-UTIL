import os
import re
import pandas as pd
import glob

download_dir = r"C:/Users/hwfrz/Downloads"

pattern = "*_DUR점검대상 비급여의약품 품목리스트(*).xlsx"
files = glob.glob(os.path.join(download_dir, pattern))

if not files:
    raise Exception("해당 패턴의 파일이 없습니다.")

latest_file = max(files, key=os.path.getctime)

sheets = pd.read_excel(latest_file, sheet_name=None)

# 첫 번째 시트만 처리
first_sheet = list(sheets.keys())[0]
df = sheets[first_sheet]

# 컬럼명 공백 제거 및 strip
df.columns = [c.replace(" ", "").strip() for c in df.columns]

# 컬럼 매핑: 원본 → 새 컬럼명
col_map = {
    "점검코드": "ppCode",
    "제품명": "ppName",
    "주성분코드": "ppMainCode",
    "규격": "ppNorm",
    "단위": "ppUnit",
    "업체명": "ppCompany"
}

# 필요한 컬럼만 추출 및 이름 변경
df = df[list(col_map.keys())].rename(columns=col_map)

# ppCode는 문자열로 강제 변환 (앞자리 0 보존) 및 공백 제거
df["ppCode"] = df["ppCode"].astype(str).str.replace(" ", "").str.strip()

# 새 엑셀로 저장
output_path = os.path.join(download_dir, "(예제파일)작업용_비급여_품목리스트.xlsx")
if os.path.exists(output_path):
    os.remove(output_path)
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="3_DRUG_비급여업데이트", index=False)
print(f"저장 완료: {output_path} (시트명: 3_DRUG_비급여업데이트, 행수: {len(df)})")
