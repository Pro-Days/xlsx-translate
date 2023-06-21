import googletrans
import kakaotrans
import pandas as pd
import urllib.request
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

options = Options()
# options.add_argument("--headless=new")
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()), options=options
)


def cefr_checker(words: list):
    driver.get("https://www.oxfordlearnersdictionaries.com/text-checker/")

    textbox = driver.find_element(By.XPATH, '//*[@id="start_text"]')
    textbox.send_keys(" ".join(words))
    wait = WebDriverWait(driver, 10)
    element = wait.until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="check_text_btn"]'))
    )

    driver.execute_script("arguments[0].click();", element)
    driver.implicitly_wait(10)

    cefr = []
    for i in range(len(words)):
        difficulty = (
            driver.find_element(By.XPATH, f'//*[@id="span_{i*2}"]')
            .get_attribute("class")
            .lower()
        )

        if "a1" in difficulty:
            difficulty = "A1"
        elif "a2" in difficulty:
            difficulty = "A2"
        elif "b1" in difficulty:
            difficulty = "B1"
        elif "b2" in difficulty:
            difficulty = "B2"
        elif "c1" in difficulty:
            difficulty = "C1"
        else:
            difficulty = "정보없음"
        cefr.append(difficulty)

    return cefr


def papagotrans(text):
    encText = urllib.parse.quote(j)
    data = "source=en&target=ko&text=" + encText
    url = "https://openapi.naver.com/v1/papago/n2mt"
    request = urllib.request.Request(url)
    request.add_header("X-Naver-Client-Id", client_id)
    request.add_header("X-Naver-Client-Secret", client_secret)
    response = urllib.request.urlopen(request, data=data.encode("utf-8"))
    response_body = response.read()

    res = json.loads(response_body.decode("utf-8"))
    return res["message"]["result"]["translatedText"]


path = "C:\\Users\\joohj\\OneDrive\\바탕 화면\\코드\\프로젝트\\엑셀번역\\"
client_id = "WEEOmXyUd_UIhMOH4avO"  # 개발자센터에서 발급받은 Client ID 값
client_secret = "JyUEl1yeot"  # 개발자센터에서 발급받은 Client Secret 값

df = pd.read_excel(path + "main_words.xlsx", index_col="no")
words = df["단어"]

google_translator = googletrans.Translator()
kakao_translator = kakaotrans.Translator()
# papago_translator = pypapago.Translator()
groupcount = 50

for i, j in enumerate(words[:100]):
    # 구글번역
    if pd.isna(df["구글번역"][i + 1]):
        try:
            google_result = google_translator.translate(j, dest="ko").text
            df["구글번역"][i + 1] = google_result
        except:
            google_result = "'구글오류'"

    else:
        google_result = "'구글x'"

    # 파파고번역
    if pd.isna(df["파파고"][i + 1]):
        try:
            papago_result = papagotrans(j)
            df["파파고"][i + 1] = papago_result
        except:
            papago_result = "'파파고오류'"
    else:
        papago_result = "'파파고x'"

    # 카카오번역
    if pd.isna(df["카카오번역"][i + 1]):
        try:
            kakao_result = kakao_translator.translate(j)
            df["카카오번역"][i + 1] = kakao_result
        except:
            kakao_result = "'카카오오류'"
    else:
        kakao_result = "'카카오x'"

    # print(i, j, google_result, papago_result, kakao_result)

    if i % groupcount == (groupcount - 1):
        wordlist = words[i - (groupcount - 1) : i + 1].values
        for w in range(len(wordlist)):
            wordlist[w] = wordlist[w].replace(".", "")
        print(wordlist)

        difficulties = cefr_checker(wordlist)
        for x, y in enumerate(difficulties):
            if (
                pd.isna(df["난이도"][i - (groupcount - 2) + x])
                or df["난이도"][i - (groupcount - 2) + x] == "정보없음"
            ):
                df["난이도"][i - (groupcount - 2) + x] = y

        for a in range(len(wordlist)):
            print(wordlist[a], difficulties[a])

        df.to_excel(path + "result.xlsx")


df.to_excel(path + "result.xlsx")


# https://www.oxfordlearnersdictionaries.com/text-checker/
