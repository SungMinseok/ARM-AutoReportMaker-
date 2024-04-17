import pandas as pd
from datetime import datetime
import os

def save_qa_info_to_csv(excel_filename):
    #excel_filename = f'CL_{country}_{project_name}_{date.strftime("%y%m%d")} {patchtype} QA.xlsx'
    try:
    # Load Excel sheet
        df = pd.read_excel(excel_filename, sheet_name='TEST REPORT', header=None)
    except FileNotFoundError :
        print(f'참조 파일이 없습니다. : {excel_filename}')
        return

    # "◆ 테스트 환경 정보" 텍스트를 찾아 해당 행 값 리턴
    target_text = "◆ 테스트 환경 정보"
    environment_row = find_row_with_text(df, target_text)

    # Extract information
    project_name = df.loc[4, 2]
    #start_date = pd.to_datetime(df.loc[6, 2])
    korean_days = ["일", "월", "화", "수", "목", "금", "토"]
    start_date = df.loc[6, 2]
    day_of_week = start_date.strftime("%w")
    start_date = start_date.strftime("%Y.%m.%d") + f"({korean_days[int(day_of_week)]})"
    end_date = df.loc[6, 6]
    day_of_week = end_date.strftime("%w")
    end_date = end_date.strftime("%Y.%m.%d") + f"({korean_days[int(day_of_week)]})"


    # Create a dictionary to hold the data
    qa_info = {
        'NAMEQA': df.loc[4, 2],
        'NAME': str(df.loc[4, 2]).replace(' QA',''),
        'START_DATE': start_date,
        'END_DATE': end_date,
        'MEMBERS': df.loc[7, 2],
        'SERVER_ALPHA': df.loc[environment_row+2, 2],
        'SERVER_LIVE': df.loc[environment_row+4, 2],
        'CLIENT_ALPHA_0': df.loc[environment_row+2, 5],
        'CLIENT_ALPHA_1': df.loc[environment_row+3, 5],
        'CLIENT_LIVE_0': df.loc[environment_row+4, 5],
        'CLIENT_LIVE_1': df.loc[environment_row+5, 5],
        'CLIENT_LIVE_2': df.loc[environment_row+6, 5],
        'CLIENT_LIVE_3': df.loc[environment_row+7, 5],
        'SUCCESS_RATE': "{:.2%}".format(round(df.loc[11, 2], 4)),
        'EXECUTION_RATE': "{:.2%}".format(round(df.loc[11, 3], 4)),
        'PASS_COUNT': df.loc[11, 4],
        'FAIL_COUNT': df.loc[11, 5],
        'NA_COUNT': df.loc[11, 6],
        'NT_COUNT': df.loc[11, 7],
        'NE_COUNT': df.loc[11, 8],
        'TOTAL_COUNT': df.loc[11, 9]
    }

    # Convert the dictionary to a DataFrame and save to CSV
    qa_df = pd.DataFrame.from_dict(qa_info, orient='index', columns=['Value'])
    qa_df.to_csv('qa_info.csv', index_label='Key', encoding='utf-8-sig')
    import os
    #os.startfile('qa_info.csv')

def find_row_with_text(df, target_text):
    for index, row in df.iterrows():
        if target_text in row.values:
            #print(index)
            return int(index)


if __name__  == "__main__" :
    # 호출 예시
    #country = 'TW'
    #date = pd.to_datetime('2023-08-16')
    name = './result/231017/CL_TW_R2M_231017 업데이트 QA.xlsx'
    save_qa_info_to_csv(name)
