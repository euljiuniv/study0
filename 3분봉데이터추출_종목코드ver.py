from pykiwoom.kiwoom import *
import pandas as pd

# 로그인
kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)

# 종목 코드 리스트
stock_codes = ["003720", "237880", "389030", "003350", "046120", "123690", "025320", "060280", "348350", "027050", "900310", "406820", "074610", "011810", "338220", "439090", "418250", "226320", "214420", "036620"]

# ExcelWriter 객체 생성
writer = pd.ExcelWriter('stock_data_20230828.xlsx')

for stock_code in stock_codes:
    # 종목 이름 가져오기
    stock_name = kiwoom.GetMasterCodeName(stock_code)
    if stock_name is None or stock_name == '':
        print(f'{stock_code}은(는) 존재하지 않는 종목 코드입니다.')
        continue

    # 각 종목의 3분봉 데이터 가져오기
    df = kiwoom.block_request("opt10080",
                              종목코드=stock_code,
                              틱범위="3",
                              수정주가구분=1,
                              output="주식분봉차트조회",
                              next=0)

    # '체결시간' 컬럼을 datetime 형태로 변환
    df['체결시간'] = pd.to_datetime(df['체결시간'])

    # 2023년 8월 8일의 데이터만 필터링
    df = df[df['체결시간'].dt.date == pd.to_datetime('2023-08-28').date()]

    # 체결시간 오름차순으로 정렬
    df = df.sort_values('체결시간')

    # 데이터를 엑셀 파일의 각 시트에 저장
    df.to_excel(writer, sheet_name=stock_name)

# ExcelWriter 객체를 닫고 엑셀 파일 저장
writer.close()
