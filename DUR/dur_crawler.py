from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time
import os
import re
import glob

def wait_for_download(download_dir, pattern, timeout=120):
    import time
    start = time.time()
    while time.time() - start < timeout:
        files = glob.glob(os.path.join(download_dir, pattern))
        if files:
            return files[0]
        time.sleep(3)
    raise Exception("다운로드 파일이 지정 시간 내에 생성되지 않았습니다.")

# chromedriver.exe 경로
chrome_driver_path = r"d:/python/WORK/DUR/chromedriver-win64/chromedriver.exe"
# 다운로드 경로
download_dir = r"C:/Users/hwfrz/Downloads"


# glob 패턴 문자열로 정의
nonpay_pattern = "*_DUR점검대상 비급여의약품 품목리스트*.xlsx"
preg_contra_pattern = "게시_임부금기_품목리스트_*.xlsx"
age_contra_pattern = "게시_연령금기_품목리스트_*.xlsx"
drug_contra_pattern = "download.zip"


chrome_options = webdriver.ChromeOptions()
# chrome_options.add_experimental_option("prefs", {
#     "download.default_directory": download_dir,
#     "download.prompt_for_download": False,
#     "directory_upgrade": True,
#     "safebrowsing.enabled": True
# })
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# 사이트 접속
driver.get("https://biz.hira.or.kr")
time.sleep(5)
    
try:        
    # 메뉴 이동: 모니터링 > DUR 대상 의약품 (팝업이 있으면 닫고 monitoring_div를 다시 찾음)
    from selenium.common.exceptions import NoSuchElementException
    while True:
        try:
            monitoring_div = driver.find_element(By.XPATH, "//div[contains(text(), '모니터링')]")
            break  # 찾으면 반복 종료
        except NoSuchElementException:
            # '모니터링' 없는 팝업 새 창 닫기
            main_window = driver.current_window_handle
            for handle in driver.window_handles:
                if handle != main_window:
                    driver.switch_to.window(handle)
                    driver.close()
            driver.switch_to.window(main_window)
            time.sleep(1)  # 잠시 대기 후 재시도

    monitoring_div.click()
    time.sleep(2)
    dur_list_div = driver.find_element(By.XPATH, "//div[contains(text(), 'DUR 대상 의약품')]")
    dur_list_div.click()
    time.sleep(10)

    # 연령금기
    latest_age_dur = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000314_form_divWork_grdList_body_gridrow_0_cell_0_1")
    latest_age_dur.click()
    time.sleep(2)
    # 첨부파일 다운로드 버튼 클릭
    attach_btn = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000314_form_divWork_Div00_grd_list_head_gridrow_-1_cell_-1_0_controlcheckbox")
    attach_btn.click()
    time.sleep(2)
    download_btn = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000314_form_divWork_Div00_btnDownFileTextBoxElement")
    download_btn.click()   
    downloaded_file = wait_for_download(download_dir, age_contra_pattern)
    
    # 목록 버튼 클릭
    drugage_list_button_div = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000314_form_divWork_btnListTextBoxElement")
    drugage_list_button_div.click()
    time.sleep(4)
    # alert = driver.switch_to.alert
    # alert.accept() # 확인
        
    # 병용금기
    latest_drug_dur = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000314_form_divWork_grdList_body_gridrow_1_cell_1_0")
    latest_drug_dur.click()
    time.sleep(3)
    # 첨부파일 다운로드 버튼 클릭
    attach_btn = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000314_form_divWork_Div00_grd_list_head_gridrow_-1_cell_-1_0_controlcheckbox")
    attach_btn.click()
    time.sleep(2)
    download_btn = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000314_form_divWork_Div00_btnDownFileTextBoxElement")
    download_btn.click()   
    downloaded_file = wait_for_download(download_dir, drug_contra_pattern)

    # 임부금기
    # 임부금기 리스트 조회
    preg_dur_list_menu = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_LeftFrame_form_div_LF_Menu_grdLeftMenu_body_gridrow_7_cell_7_0_controltreeTextBoxElement")
    preg_dur_list_menu.click()
    time.sleep(3)
    
    # 최신 임부금기 고시
    latest_preg_dur = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000315_form_divWork_grdList_body_gridrow_0_cell_0_1")
    latest_preg_dur.click()
    time.sleep(3)
    # 첨부파일 다운로드 버튼 클릭
    attach_btn = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000315_form_divWork_Div00_grd_list_head_gridrow_-1_cell_-1_0_controlcheckbox")
    attach_btn.click()
    time.sleep(2)
    download_btn = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000315_form_divWork_Div00_btnDownFileTextBoxElement")
    download_btn.click()
    downloaded_file = wait_for_download(download_dir, preg_contra_pattern)
    
    # 비급여
    # 비급여 리스트 조회
    nonpay_dur_list_menu = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_LeftFrame_form_div_LF_Menu_grdLeftMenu_body_gridrow_10_cell_10_0_controltreeTextBoxElement")
    nonpay_dur_list_menu.click()
    time.sleep(3)

    # 최신 비급여 고시
    latest_nonpay_dur = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000319_form_divWork_grdList_body_gridrow_0_cell_0_1")
    latest_nonpay_dur.click()
    time.sleep(3)
    # 첨부파일 다운로드 버튼 클릭
    attach_btn = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000319_form_divWork_Div00_grd_list_head_gridrow_-1_cell_-1_0_controlcheckbox")
    attach_btn.click()
    time.sleep(2)
    download_btn = driver.find_element(By.ID, "MainFrame_VFrameSet0_HFrameSet0_VFrameSet1_WorkFrame_MP00000319_form_divWork_Div00_btnDownFileTextBoxElement")
    download_btn.click()
    downloaded_file = wait_for_download(download_dir, nonpay_pattern)
    time.sleep(10)
    
except Exception as e:
    print("Error occurred:", e)

print("DUR 크롤러 작업이 완료되었습니다.")