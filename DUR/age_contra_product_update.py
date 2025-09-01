
import os
import pandas as pd
import re
import requests

# 다운로드 폴더에서 '게시_연령금기_품목리스트_yymm.xlsx' 형식의 파일명만 필터링
download_dir = r"C:/Users/hwfrz/Downloads"
pattern = re.compile(r"게시_연령금기_품목리스트_\d{4}\.xlsx$")
files = [
    os.path.join(download_dir, f)
    for f in os.listdir(download_dir)
    if pattern.match(f)
]
if not files:
    raise Exception("해당 패턴의 파일이 없습니다.")

latest_file = max(files, key=os.path.getctime)

def process_df(df):
    # 컬럼명을 무시하고 순서대로 강제 할당
    df.columns = [
        "ingName", "ppMainCode", "ppCode", "proKorname", "pdKorName",
        "contraAge", "contraAgeUnit1", "contraAgeUnit2", "reference", "appliedDate", "remark"
    ]

    # ppCode는 문자열로 강제 변환 (앞자리 0 보존)
    df["ppCode"] = df["ppCode"].astype(str)
    # ppCode가 null, NaN, 빈 문자열인 행 제거
    df = df[df["ppCode"].notna() & (df["ppCode"].str.strip() != "")].reset_index(drop=True)

    def process_unit1(val):
        if pd.isna(val):
            return ""
        if "세 (3)" in str(val):
            return "세"
        elif "개월 (2)" in str(val):
            return "개월"
        elif "주 (1)" in str(val):
            return "주"
        else:
            return str(val).strip()

    def process_unit2(val):
        if pd.isna(val):
            return ""
        if "이상 (4)" in str(val):
            return "이상"
        elif "이하 (2)" in str(val):
            return "이하"
        elif "미만 (1)" in str(val):
            return "미만"
        else:
            return str(val).strip()

    df["contraAgeUnit"] = df["contraAgeUnit1"].apply(process_unit1) + " " + df["contraAgeUnit2"].apply(process_unit2)
    df["contraAgeUnit"] = df["contraAgeUnit"].str.strip()

    final_cols = [
        "ingName", "ppMainCode", "ppCode", "proKorname", "pdKorName",
        "contraAge", "contraAgeUnit1", "contraAgeUnit2", "reference",
        "appliedDate", "remark", "contraAgeUnit"
    ]
    return df[final_cols]


# 두 시트 모두 읽어서 전처리 (ppCode 컬럼을 문자열로 강제 지정)
def read_excel_force_str_col(file, sheet_name=None):
    # 일단 전체를 object로 읽고, 3번째 컬럼(ppCode)만 str로 변환
    sheets = pd.read_excel(file, sheet_name=sheet_name, dtype=object)
    for name, df in sheets.items() if isinstance(sheets, dict) else [(None, sheets)]:
        # 3번째 컬럼(ppCode)만 str로 변환 (0으로 시작하는 보험코드)
        if df.shape[1] > 2:
            df.iloc[:,2] = df.iloc[:,2].apply(lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x) if not pd.isna(x) else "")
        sheets[name] = df
    return sheets

sheets = read_excel_force_str_col(latest_file, sheet_name=None)
dfs = []
for sheet_name, df in sheets.items():
    dfs.append(process_df(df))
merged_df = pd.concat(dfs, ignore_index=True)

# (예제파일)작업용_연령금기_급여_품목리스트.xlsx 파일의 age_contra_product_backup 시트에 저장
output_path = os.path.join(download_dir, "(예제파일)작업용_연령금기_급여_품목리스트.xlsx")

# 기존 파일이 있으면 삭제
if os.path.exists(output_path):
    os.remove(output_path)
with pd.ExcelWriter(output_path) as writer:
    merged_df.to_excel(writer, sheet_name="age_contra_product_backup", index=False)

# 임시테이블 Truncate
# truncate_url = "https://manage.druginfo.co.kr/Druginfo/DURUpdate/AgeContraProductTableTruncate.aspx"
# try:
#     truncate_resp = requests.post(truncate_url, timeout=10)
#     print("Truncate 결과:", truncate_resp.status_code, truncate_resp.text)
# except Exception as e:
#     print("Truncate 요청 실패:", e)

# # 생성된 파일을 AgeContraProductExcelUpload.aspx로 임시 테이블에 업로드
# upload_url = "https://manage.druginfo.co.kr/Druginfo/DURUpdate/AgeContraProductExcelUpload.aspx"
# try:
#     with open(output_path, "rb") as f:
#         files = {'file': (os.path.basename(output_path), f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
#         response = requests.post(upload_url, files=files, timeout=30)
#         print("업로드 결과:", response.status_code, response.text)
# except Exception as e:
#     print("업로드 요청 실패:", e)

# # 전체 적용
# apply_url = "https://manage.druginfo.co.kr/Druginfo/DURUpdate/AgeContraProductAllApply.aspx"
# try:
#     apply_resp = requests.post(apply_url, timeout=10)
#     print("전체 적용 결과:", apply_resp.status_code, apply_resp.text)
# except Exception as e:
#     print("전체 적용 요청 실패:", e)
