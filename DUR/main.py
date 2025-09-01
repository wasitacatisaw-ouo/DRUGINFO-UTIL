import datetime
import subprocess
import sys

today = datetime.date.today()

# 매월 1일 또는 첫 번째 월요일 병용금기, 연령금기, 임부금기, 비급여 고시 다운로드
is_first_day = today.day == 1
is_first_monday = today.weekday() == 0 and 1 <= today.day <= 7

if is_first_day or is_first_monday or 1==1:
    subprocess.run([sys.executable, "dur_crawler.py"])


# 매월 15일 또는 15일 이후 첫 번째 워킹데이 비용효과적 함량 고시 다운로드
is_fifteenth = today.day == 15
is_first_working_day_after_15th = today.day > 15 and today.weekday() < 5 and all(
    (today - datetime.timedelta(days=offset)).day < 15 or (today - datetime.timedelta(days=offset)).weekday() >= 5
    for offset in range(1, today.day - 14)
)

if is_fifteenth or is_first_working_day_after_15th:
    subprocess.run([sys.executable, "ldlist_crawler.py"])