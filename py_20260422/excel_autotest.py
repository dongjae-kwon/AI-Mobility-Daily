import pandas as pd

#파일 불러오기
df = pd.read_excel("sample.xlsx")

#중복 제거
df = df.drop_duplicates()

# 정렬 (ex: 가격)
df = df.sort_values(by = "가격")

# 엑셀 저장
df.to_excel("output.xlsx", index=False)