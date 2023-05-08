import googletrans
import pandas as pd
import time

translator = googletrans.Translator()

path = ""

# 시트1
df = pd.read_excel(path + "words.xlsx", engine="openpyxl", header=3, sheet_name="1")

# 엑셀파일에 맞게 셀 설정 필요
for i in range(4):
    for j in range(len(df.values)):
        en_word = str(df.values[j][i * 3 + 1])
        result = translator.translate(en_word, dest="ko")
        df[f"Unnamed: {i*3+2}"].loc[j] = result.text
        print(
            str(df.values[j][i * 3])
            + ". "
            + str(df.values[j][i * 3 + 1])
            + ": "
            + str(df.values[j][i * 3 + 2])
        )
        time.sleep(0.1)

df.to_excel(path + "words-output1.xlsx", sheet_name="1", index=False)


# 시트2
df = pd.read_excel(path + "words.xlsx", engine="openpyxl", header=3, sheet_name="2")

for i in range(4):
    for j in range(len(df.values)):
        en_word = str(df.values[j][i * 3 + 1])
        if en_word != "nan":
            result = translator.translate(en_word, dest="ko")
            df[f"Unnamed: {i*3+2}"].loc[j] = result.text
            print(
                str(df.values[j][i * 3])
                + ". "
                + str(df.values[j][i * 3 + 1])
                + ": "
                + str(df.values[j][i * 3 + 2])
            )
            time.sleep(0.2)
        else:
            pass

df.to_excel(path + "words-output2.xlsx", sheet_name="2", index=False)
