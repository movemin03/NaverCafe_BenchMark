from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import openpyxl
import os
from datetime import datetime
from openpyxl.styles import PatternFill

# 사용자 지정
exception_keyword = ["부동산", "투자", "실거래가"]
member_min = 500 # 무제한은 0 입력
start_pg = 1

keyword = input("검색할 키워드를 입력해주세요:")
print("검색할 키워드는 " + keyword + "입니다")
print("키워드 정확도순으로 검색하려면 1을, 랭킹 순으로 검색하시려면 2를 눌러주세요. 기본값은 정확도 순입니다")
qu = input("")
if qu =="2":
    url = "https://section.cafe.naver.com/ca-fe/home/search/cafes?q=" + keyword + "&t=1700821664452&od=1"
    print("랭킹순으로 검색합니다")
else:
    url = "https://section.cafe.naver.com/ca-fe/home/search/cafes?q=" + keyword + "&t=1701084772896"
    print("정확도순으로 검색합니다")


user = os.getlogin()

# 데이터프레임 생성
data = {
    'Link': [],
    'Cafe Title': [],
    'Member Num': [],
    'Ranking': [],
    'NewPostNum': [],
    'TotalPostNum': [],
    'TotalScore': []
}

start_time = datetime.now()
# 크롬드라이버 실행
print("크롬 로딩 중입니다... 컴퓨터 사양에 따라 실행에 시간이 걸릴 수 있습니다")
options = webdriver.ChromeOptions()
options.add_argument("headless")
# driver = webdriver.Chrome() 크롬창 headless 모드를 해제하려면 이것 사용
driver = webdriver.Chrome(options=options)
driver.get(url)

wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="mainContainer"]/div[2]/div[1]/div[2]/div[1]/div[1]')))
end_time = datetime.now()
execution_time = end_time - start_time
print(f"크롬 로딩 시간: {execution_time}")

cafe_total = int(driver.find_element(By.XPATH, '//*[@id="mainContainer"]/div[2]/div[1]/div[2]/div[1]/div[1]').text.replace(',', ''))
max_pg = round(cafe_total / 12)
print("최대 탐색가능한 페이지는 " + str(max_pg) + "입니다")

last_pg = int(input("어디까지 검색할까요? 숫자만 입력: "))
print("시작페이지는 " + str(start_pg) + ", 마지막페이지는 " + str(last_pg) + "로 지정되었습니다")

def search(search_pg):
    print("탐색 시작")
    for item_num in range(1, 13):
        item_num_link = "//div[@class='item_list']//div[@class='CafeItem'][" + str(item_num) + "]//a"
        print(str(search_pg) + "페이지 중 " + str(item_num) + "/12항목 = 전체 중 " + str(12*(search_pg-1)+item_num) + "번 검색 중")
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, item_num_link)))
            link = driver.find_element(By.XPATH, item_num_link).get_attribute("href")
            cafe_title = driver.find_element(By.XPATH,item_num_link + "/div[@class='cafe_info_area']/strong[@class='cafe_name']").text
            member_num = int(driver.find_element(By.XPATH,item_num_link + "/div[@class='cafe_info_area']/div[@class='cafe_info_detail']/*[2]/*[2]").text.replace(",", ""))
            ranking = driver.find_element(By.XPATH,item_num_link + "/div[@class='cafe_info_area']/div[@class='cafe_info_detail']/*[3]/*[2]").text
            ranking_main = ranking[:2]
            ranking_dic = {'숲': 6.83, '나무': 6, '열매': 5, '가지': 4, '잎새': 3, '새싹': 2, '씨앗': 1}
            if ranking_main in ranking_dic:
                ranking_main_num = ranking_dic[ranking_main]
            ranking_sub = int(ranking[2:3])
            if ranking_sub == "":
                ranking_sub = 1
            post_num = driver.find_element(By.XPATH,item_num_link + "/div[@class='cafe_info_area']/div[@class='cafe_info_detail']/*[4]/*[2]").text
            new_post_num, total_post_num = post_num.split("/")
            new_post_num = int(new_post_num.replace(",", "").replace(" ", ""))
            total_post_num = int(total_post_num.replace(",", "").replace(" ", ""))

            # 점수계산 = (멤버 수 * 0.3) + ((랭킹대분류+1/6*(랭킹소분류-1)) * 0.6 * 100) + (새로운 글 수 * 0.015) + (전체 글 수 * 0.085)
            score = round((member_num * 0.3) + ((ranking_main_num + 1/6 * (ranking_sub - 1))*0.6*100) + (new_post_num *0.015) + (total_post_num *0.085))

            if any(word in cafe_title for word in exception_keyword):
                print("예외 처리: Cafe Title 에 예외처리한 단어가 포함되어 있으므로 엑셀 파일에는 포함되지 않습니다")
            elif member_min != 0 and member_num < member_min:
                print("예외 처리: 회원 수가 사용자가 지정한 최저 사용자 수인 " + str(member_min) + "보다 적은 " + str(member_num) + "명이라 엑셀 파일에 포함하지 않습니다")

            else:
                data['Link'].append(link)
                data['Cafe Title'].append(cafe_title)
                data['Member Num'].append(str(member_num))
                data['Ranking'].append(ranking)
                data['NewPostNum'].append(str(new_post_num))
                data['TotalPostNum'].append(str(total_post_num))
                data['TotalScore'].append(str(score))
                print("검색 완료")
        except:
            print("조회된 항목이 없습니다")
            pass

start_time = datetime.now()
for search_pg in range(start_pg, last_pg+1):
    wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='item_list']//div[@class='CafeItem'][1]//a")))
    if search_pg % 10 == 1 and search_pg != 1:
        print("페이지 넘어갑니다")
        driver.find_element(By.XPATH,'//button[@type="button" and contains(@class, "btn type_next")]').click()
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='item_list']//div[@class='CafeItem'][1]//a")))
    pg_button = driver.find_element(By.XPATH, f"//div[@class='ArticlePaginate']//button[contains(text(), '{search_pg}')]")
    pg_button.click()
    wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='item_list']//div[@class='CafeItem'][1]//a")))
    print("\n페이지" + str(search_pg) + "을(를) 탐색합니다\n")
    search(search_pg)

# 데이터프레임 생성
df = pd.DataFrame(data)
df['TotalScore'] = pd.to_numeric(df['TotalScore'])
median = df['TotalScore'].drop_duplicates().median()

file_time = datetime.today().strftime("%Y%m%d_%H%M")
file_path = 'C:\\Users\\' + user + '\\Desktop\\cafe_data_' + keyword +"_"+ file_time
sorted_df = df.sort_values(by='TotalScore', ascending=False)

# 엑셀 파일로 저장
with pd.ExcelWriter(file_path +".xlsx") as writer:
    sorted_df.to_excel(writer, index=False, sheet_name='Sheet1')

    # 중간값 셀을 빨간색으로 표시
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    median_cell = worksheet.cell(row=1, column=sorted_df.columns.get_loc('TotalScore') + 2)
    median_cell.fill = red_fill
    median_cell.value = median

print("\n" + file_path + ".xlsx 에 저장 완료되었습니다")
driver.quit()
end_time = datetime.now()
execution_time = end_time - start_time
print(f"검색 소요 시간: {execution_time}")
print("total_Score 의 중복 제외 중간값은 " + str(median)+ " 입니다")

