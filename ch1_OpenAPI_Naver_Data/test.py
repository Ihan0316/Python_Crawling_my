# step1.프로젝트에 필요한 패키지 불러오기
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import pandas as pd

# step2.로그인 정보 및 검색할 회사1 미리 정의, 해당 회사의 리뷰 끝 페이지도 정의
print("-" * 60)
usr = input('아이디를 입력해주세요 : ')
pwd = input('비밀번호를 입력해주세요 : ')
query = input('검색할 회사를 입력해주세요 : ')
page = int(input('검색할 페이지 수를 입력해주세요 : '))
print("-" * 60)

# step3.크롬드라이버 실행 및 잡플래닛 로그인
service = Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

driver.get("https://www.jobplanet.co.kr/users/sign_in?_nav=gb")
driver.implicitly_wait(10)

login_id = driver.find_element(By.CSS_SELECTOR, "input#user_email")
login_id.send_keys(usr)

login_pwd = driver.find_element(By.CSS_SELECTOR, "input#user_password")
login_pwd.send_keys(pwd)

login_id.send_keys(Keys.RETURN)
driver.implicitly_wait(10)

# step4.알림창 뜨면 닫기
try:
    driver.find_element(By.CLASS_NAME, 'jp-x-circle_fill').click()
    print('알림창을 닫았습니다.')
    print("-" * 60)
except:
    print('알림창이 없습니다.')
    print("-" * 60)

# step5.원하는 회사의 리뷰 페이지까지 이동
search_query = driver.find_element(By.CSS_SELECTOR, "input#search_bar_search_query")
search_query.send_keys(query)
search_query.send_keys(Keys.RETURN)
driver.implicitly_wait(10)

driver.find_element(By.CSS_SELECTOR, "a.tit").click()
driver.implicitly_wait(10)

# step6.알림창 뜨면 닫기
try:
    driver.find_element(By.CSS_SELECTOR, "button.btn_close_x_ty1").click()
    print('알림창을 닫았습니다.')
    print("-" * 60)
except:
    print('알림창이 없습니다.')
    print("-" * 60)

# step7.크롤링한 정보를 담을 리스트명 정의
list_company = []
list_div = []
list_cur = []
list_date = []
list_stars = []
list_summery = []
list_merit = []
list_disadvantages = []
list_managers = []

# step8.원하는 회사의 직무/근속여부/일시/요약/평점/장점/단점/경영진에게 바라는 점 크롤링 (for문으로 반복)
counter = 1
for i in range(page):

    # 직무, 근속여부, 일시
    user_info = driver.find_elements(By.CSS_SELECTOR, "span.txt1")

    count = int(len(user_info) / 4)

    for j in range(count):
        list_company.append(query)
        list_user_info = []

    for j in user_info:
        list_user_info.append(j.text)

    # 한 페이지에 정보 5set씩 나옴. 마지막 페이지는 5개 미만일 수 있으므로 count 변수를 반복횟수로 넣어줌.
    for j in range(count):
        a = list_user_info[4 * j]
        list_div.append(a)

        b = list_user_info[4 * j + 1]
        list_cur.append(b)

        c = list_user_info[4 * j + 3]
        list_date.append(c)

        print(f"{counter}번째 리뷰를 크롤링중입니다\n")
        counter += 1

    # 별점
    stars = driver.find_elements(By.CSS_SELECTOR, "div.star_score")
    for j in stars:
        a = j.get_attribute('style')
        if a[7:9] == '20':
            list_stars.append("1점")
        elif a[7:9] == '40':
            list_stars.append("2점")
        elif a[7:9] == '60':
            list_stars.append("3점")
        elif a[7:9] == '80':
            list_stars.append("4점")
        else:
            list_stars.append("5점")

    # 요약 정보
    summery = driver.find_elements(By.CSS_SELECTOR, "h2.us_label")

    for j in summery:
        list_summery.append(j.text)

    # 장점, 단점, 경영진에게 바라는 점
    list_review = []

    review = driver.find_elements(By.CSS_SELECTOR, "dd.df1")

    for j in review:
        list_review.append(j.text)

    # 한 페이지에 정보 5set씩 나옴. 마지막 페이지는 5개 미만일 수 있으므로 count 변수를 반복횟수로 넣어줌.
    for j in range(count):
        a = list_review[3 * j]
        list_merit.append(a)

        b = list_review[3 * j + 1]
        list_disadvantages.append(b)

        c = list_review[3 * j + 2]
        list_managers.append(c)

    # 다음 페이지 클릭 후 for문 진행, 끝 페이지에서 다음 페이지 클릭 안되는 것 대비해서 예외처리 구문 추가
    try:
        driver.find_element(By.CSS_SELECTOR, "a.btn_pgnext").click()
        time.sleep(5)
    except:
        pass

# step9.pandas 라이브러리로 표 만들기
total_data = pd.DataFrame()
total_data['회사'] = pd.Series(list_company)
total_data['날짜'] = pd.Series(list_date)
total_data['직무'] = pd.Series(list_div)
total_data['재직여부'] = pd.Series(list_cur)
total_data['별점'] = pd.Series(list_stars)
total_data['요약'] = pd.Series(list_summery)
total_data['장점'] = pd.Series(list_merit)
total_data['단점'] = pd.Series(list_disadvantages)
total_data['경영진에게 바라는 점'] = pd.Series(list_managers)

# step10.엑셀 형태로 저장하기
total_data.index = total_data.index + 1
total_data.to_excel("잡플래닛 리뷰_" + query + ".xlsx", index=True)

# step11.크롬 드라이버 종료
print("-" * 60)
print("크롤링이 완료되었습니다!")
print("-" * 60)
driver.close()