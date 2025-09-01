import os
import re
import pandas as pd

download_dir = r"C:/Users/hwfrz/Downloads"

pattern = re.compile(r"게시_임부금기_품목리스트_\d{4}\.xlsx$")
files = [
    os.path.join(download_dir, f)
    for f in os.listdir(download_dir)
    if pattern.match(f)
]
if not files:
    raise Exception("해당 패턴의 파일이 없습니다.")

latest_file = max(files, key=os.path.getctime)

sheets = pd.read_excel(latest_file, sheet_name=None)

all_dfs = []
for sheet, df in sheets.items():
    df.columns = [
        "ingName", "ppMainCode", "ppCode", "proKorname", "pdKorName",
        "appliedDate", "notificationInfo", "contraRank", "etc"
    ]
    # ppCode는 문자열로 강제 변환 (앞자리 0 보존)
    df["ppCode"] = df["ppCode"].astype(str)
    # ppCode가 null, NaN, 빈 문자열인 행 제거
    df = df[df["ppCode"].notna() & (df["ppCode"].str.strip() != "")].reset_index(drop=True)
    all_dfs.append(df)

if not all_dfs:
    raise Exception("병합할 데이터가 없습니다. 파일/시트 구조를 확인하세요.")

merged_df = pd.concat(all_dfs, ignore_index=True)
# id 컬럼 1부터 부여
merged_df.insert(0, "id", range(1, len(merged_df) + 1))
final_cols = [
    "id", "ingName", "ppMainCode", "ppCode", "proKorname", "pdKorName",
    "appliedDate", "notificationInfo", "contraRank", "etc"
]
merged_df = merged_df[final_cols]
print(f"전체 병합 행 개수: {len(merged_df)}")
# 병합파일 저장
output_path = os.path.join(download_dir, "(예제파일)작업용_임부금기_급여_품목리스트.xlsx")
if os.path.exists(output_path):
    os.remove(output_path)
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    merged_df.to_excel(writer, sheet_name="preg_contra_product_temp", index=False)
print(f"저장 완료: {output_path}")
