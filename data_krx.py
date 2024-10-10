import pandas as pd
from pykrx import stock
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog

# 데이터 가져오기 함수
def get_market_data(market_code):
    today = datetime.today().strftime('%Y-%m-%d')
    df = stock.get_index_ohlcv("1980-01-01", today, market_code)
    return df

# 최신 데이터 업데이트 함수
def update_data():
    global data_kospi, data_kosdaq  # 전역 변수로 선언하여 업데이트할 수 있게 함
    try:
        data_kospi = get_market_data("1001")  # KOSPI 최신 데이터 가져오기
        data_kosdaq = get_market_data("2001")  # KOSDAQ 최신 데이터 가져오기

        # 각 데이터에 등락률 및 등락폭 계산을 다시 수행
        data_kospi = add_calculated_columns(data_kospi)
        data_kosdaq = add_calculated_columns(data_kosdaq)

        messagebox.showinfo("업데이트 완료", "코스피와 코스닥 데이터가 최신 값으로 업데이트되었습니다.")
    except Exception as e:
        messagebox.showerror("업데이트 실패", f"데이터를 업데이트하는 중 오류가 발생했습니다: {str(e)}")

# KOSPI와 KOSDAQ 데이터를 가져옴
data_kospi = get_market_data("1001")
data_kosdaq = get_market_data("2001")

# 등락률 및 등락폭 추가 함수
def add_calculated_columns(data):
    data['등락률_raw'] = data['종가'].pct_change() * 100  # 등락률 계산
    data['등락폭_raw'] = data['종가'].diff()  # 등락폭 계산

    # 포맷 설정 - 표시용 컬럼 추가
    data['등락률'] = data['등락률_raw'].apply(lambda x: f"{x:.2f}%" if pd.notnull(x) else "0.00%")
    data['등락폭'] = data['등락폭_raw'].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "0.00")
    data['거래대금'] = data['거래대금'].apply(lambda x: f"{int(x):,}")
    data['상장시가총액'] = data['상장시가총액'].apply(lambda x: f"{int(x):,}")
    data['등락률_raw'] = data['등락률_raw'].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "0.00")
    data['등락폭_raw'] = data['등락폭_raw'].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "0.00")

    return data

# 각 데이터에 대해 열 순서 재배치 및 계산된 열 추가
data_kospi = add_calculated_columns(data_kospi)[['종가', '등락폭', '등락률', '시가', '고가', '저가', '거래대금', '상장시가총액', '등락률_raw', "등락폭_raw"]]
data_kosdaq = add_calculated_columns(data_kosdaq)[['종가', '등락폭', '등락률', '시가', '고가', '저가', '거래대금', '상장시가총액', '등락률_raw', "등락폭_raw"]]

# 선택된 시장에 따라 데이터를 반환
def get_selected_data():
    selected_market = market_var.get()
    if selected_market == "KOSPI":
        return data_kospi
    elif selected_market == "KOSDAQ":
        return data_kosdaq

# 결과를 Excel 파일로 저장
def save_to_excel(dataframe):
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if file_path:
        try:
            dataframe.to_excel(file_path, index=True)
            messagebox.showinfo("저장 완료", f"파일이 저장되었습니다: {file_path}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"파일 저장 중 오류 발생: {e}")

# 연속 상승/하락 일자 엑셀로 저장 함수
def save_to_excel_1(dataframes, filename='연속_상승_하락_일자_정보.xlsx'):
    # 파일 저장 위치를 사용자에게 묻기
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=filename)

    # 선택한 경로가 없으면 종료
    if not filepath:
        return

    # 데이터프레임을 엑셀 파일로 저장
    with pd.ExcelWriter(filepath) as writer:
        if isinstance(dataframes, list):
            for i, df in enumerate(dataframes):
                sheet_name = f"Sheet{i + 1}"
                if i == 0:
                    sheet_name = "상승_일자"
                elif i == 1:
                    sheet_name = "하락_일자"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            dataframes.to_excel(writer, index=False)


# 조건을 만족하는 가장 최근 일자와 값 및 순위 출력
def show_latest_matching_info():
    try:
        data = get_selected_data()  # 선택된 시장 데이터를 가져옴
        threshold_value = float(entry_value.get())
        selected_criteria = criteria_var.get()
        comparison_operator = comparison_var.get()

        # '등락률'은 원본 숫자 컬럼을 사용 ('등락률_raw')
        if selected_criteria == '등락률':
            selected_criteria = '등락률_raw'

        # 선택된 기준이 데이터 프레임에 포함된 경우
        if selected_criteria in ['종가', '장중 가격', '등락률_raw', '등락폭', '상장시가총액', '거래대금']:
            # '장중 가격'은 '고가'를 의미
            if selected_criteria == '장중 가격':
                selected_criteria = '고가'

            # 데이터 타입을 숫자로 변환 (문자열이 포함될 경우 처리)
            if data[selected_criteria].dtype == 'object':
                # 문자열에서 숫자로 변환하기 위한 정리
                data[selected_criteria] = data[selected_criteria].str.replace(',', '', regex=False).astype(float)

            # 기준 열이 NaN으로 변환된 경우의 처리
            if data[selected_criteria].isna().any():
                raise ValueError("데이터에 숫자로 변환할 수 없는 값이 포함되어 있습니다.")

            # 조건에 맞는 데이터 필터링
            if comparison_operator == '이상':
                filtered_data = data[data[selected_criteria] >= threshold_value]
            elif comparison_operator == '이하':
                if selected_criteria == '고가':
                    # '이하'인 경우 '저가'로 필터링
                    filtered_data = data[data['저가'] <= threshold_value]
                else:
                    filtered_data = data[data[selected_criteria] <= threshold_value]
            else:
                messagebox.showerror("선택 오류", "비교 연산자를 올바르게 선택해 주세요.")
                return
        else:
            messagebox.showerror("선택 오류", "유효한 기준을 선택해 주세요.")
            return

        if not filtered_data.empty:
            latest_date = filtered_data.index[-1].strftime('%Y-%m-%d')
            latest_value = filtered_data.iloc[-1][selected_criteria]
            rank = data[selected_criteria].rank(ascending=False, method='min')[latest_date]

            # '등락률' 포맷팅된 값을 보여줌
            if selected_criteria == '등락률_raw':
                selected_criteria_display = '등락률'
                latest_value_display = f"{latest_value:.2f}%"
            else:
                selected_criteria_display = selected_criteria
                latest_value_display = latest_value

            messagebox.showinfo("조건 만족 가장 최근 일자 정보",
                                f"기준: {selected_criteria_display}\n가장 최근 일자: {latest_date}\n"
                                f"값: {latest_value_display}\n전체 중 {rank:.0f}등")
        else:
            messagebox.showinfo("결과 없음", f"{selected_criteria}가 {threshold_value} {comparison_operator}인 데이터가 없습니다.")
    except ValueError as ve:
        messagebox.showerror("입력 오류", str(ve))
    except Exception as e:
        messagebox.showerror("오류", f"처리 중 오류가 발생했습니다: {e}")

# 오늘자 기준으로 모든 값 조회 (오늘자 제외)
def show_today_info():
    try:
        data = get_selected_data()  # 선택된 시장 데이터를 가져옴
        comparison_operator = comparison_var.get()  # "이상" 또는 "이하" 선택
        today_date = data.index[-1]

        previous_data = data[:-1]

        results = []
        # 조건에 따라 사용할 기준 리스트를 변경
        criteria_list = ['종가', '고가', '등락률', '등락폭', '거래대금', '상장시가총액']

        if comparison_operator == '이하':
            criteria_list = ['종가', '저가', '등락률', '등락폭', '거래대금', '상장시가총액']

        for criteria in criteria_list:
            today_value = data.iloc[-1][criteria.replace('_raw', '')]  # 포맷된 컬럼 대신 원본 사용
            today_value_raw = data.iloc[-1][criteria]  # 원본 숫자 값 사용

            if comparison_operator == '이상':
                filtered_data = previous_data[previous_data[criteria] >= today_value_raw]
            elif comparison_operator == '이하':
                filtered_data = previous_data[previous_data[criteria] <= today_value_raw]

            if not filtered_data.empty:
                latest_date = filtered_data.index[-1].strftime('%Y-%m-%d')
                latest_value = filtered_data.iloc[-1][criteria.replace('_raw', '')]
                days_since_last_cross = (today_date - filtered_data.index[-1]).days
                rank = data[criteria].rank(ascending=False, method='min')[today_date]

                results.append(f"{criteria.replace('_raw', '')}:\n"
                               f"오늘자 값: {today_value}\n"
                               f"가장 최근 일자: {latest_date} (값: {latest_value})\n"
                               f"최근 {days_since_last_cross}일 전 동일한 값 도달\n"
                               f"전체 중 {rank:.0f}등\n")
            else:
                results.append(f"{criteria.replace('_raw', '')}: 결과 없음\n")

        messagebox.showinfo("오늘 기준 정보", "\n".join(results))
    except ValueError:
        messagebox.showerror("입력 오류", "유효한 숫자를 입력해주세요.")

#연도별 종가 등락률, 종가 표기
def calculate_yearly_end_of_year_prices(data, special_year=None, next_year=None):
    # 연도별 연말 마지막 거래일의 종가 계산
    end_of_year_prices = data.resample('YE').apply(lambda df: df['종가'].iloc[-1] if not df.empty else None)

    # 연도별 연초 종가 계산
    start_of_year_prices = data.resample('YE').apply(lambda df: df['종가'].iloc[0] if not df.empty else None)

    # 데이터프레임 생성
    yearly_prices_df = pd.DataFrame({
        '연말종가': end_of_year_prices,
        '연초종가': start_of_year_prices,
    }).dropna()

    # 연말 종가 등락률 계산
    # 특별한 연도 처리
    if special_year and next_year:
        if special_year in yearly_prices_df.index:
            end_of_special_year_price = end_of_year_prices.loc[special_year]
            start_of_special_year_price = start_of_year_prices.loc[special_year]
            yearly_prices_df.loc[special_year, '연말종가등락률'] = (
                                                                        end_of_special_year_price - start_of_special_year_price) / start_of_special_year_price * 100

    # 다른 연도들에 대한 연말 종가 등락률 계산
    yearly_prices_df['연말종가등락률'] = yearly_prices_df['연말종가'].pct_change() * 100

    return yearly_prices_df

# 연도별 일평균 거래대금, 연도별 지수 등락률 값을 구하는 함수의 전처리
def preprocess_data(data):
    print("Preprocessing data...")
    print(f"Initial data types:\n{data.dtypes}")

    # 거래대금의 쉼표 제거 및 숫자형으로 변환
    try:
        if data['거래대금'].dtype == 'object':
            data['거래대금'] = data['거래대금'].str.replace(',', '', regex=False)
            print("Comma removed from 거래대금.")
        data['거래대금'] = pd.to_numeric(data['거래대금'], errors='coerce')
        print("거래대금 converted to numeric.")
    except Exception as e:
        print(f"Error in preprocess_data: {e}")
        raise  # 오류를 다시 발생시켜 디버깅을 이어갈 수 있게 합니다.

    # 숫자형 데이터 확인
    if not pd.api.types.is_numeric_dtype(data['거래대금']):
        raise ValueError("거래대금 열이 숫자형 데이터가 아닙니다.")

    print("Preprocessing complete.")
    return data

# 연도별 일평균 거래대금, 연도별 지수 등락률 값을 구하는 함수
def show_yearly_avg_info():
    try:
        data = get_selected_data()

        # 데이터 전처리
        data = preprocess_data(data)

        # 연도별 일평균 거래대금 계산
        yearly_avg = data.resample('YE').agg({'거래대금': 'mean'})

        # 연도별 일평균 거래대금 등락률 계산
        yearly_avg['거래대금 등락률'] = yearly_avg['거래대금'].pct_change() * 100

        # 첫 연도의 거래대금 등락률을 NaN으로 설정
        first_year = yearly_avg.index[0]
        yearly_avg.at[first_year, '거래대금 등락률'] = pd.NA  # NaN으로 변경

        yearly_avg = yearly_avg.dropna()

        # 연도별 연말 마지막 거래일의 종가와 등락률 계산
        if market_var.get() == "KOSPI":
            end_of_year_prices_df = calculate_yearly_end_of_year_prices(data, 1980, 1980)
        elif market_var.get() == "KOSDAQ":
            end_of_year_prices_df = calculate_yearly_end_of_year_prices(data, 1996, 1996)
        else:
            end_of_year_prices_df = calculate_yearly_end_of_year_prices(data)

        # 연도별 마지막 거래일 추출
        last_trading_dates = data.resample('YE').apply(lambda x: x.index[-1].strftime('%Y-%m-%d'))
        last_trading_dates_1 = last_trading_dates['종가']
        last_trading_dates_df = pd.DataFrame({'마지막 거래일': last_trading_dates_1})

        # 데이터프레임 병합을 위해 인덱스 일관성 유지
        yearly_info = yearly_avg.join(end_of_year_prices_df, how='inner')
        yearly_info = yearly_info.join(last_trading_dates_df, how='inner')

        # 연도 열 추가
        yearly_info['연도'] = yearly_info.index.year
        yearly_info = yearly_info.reset_index(drop=True)

        # 연도별 추가 데이터 처리
        if market_var.get() == "KOSPI":
            # 1980년도 값을 수동으로 추가
            if 1980 not in yearly_info['연도'].values:
                end_1980_price = 106.87  # 1980년도 연말 종가
                start_1980_price = 100  # 1980년도 연초 종가

                new_row = pd.DataFrame({
                    '연도': [1980],
                    '마지막 거래일': [data[data.index.year == 1980].index[-1].strftime('%Y-%m-%d')],
                    '연말종가': [end_1980_price],
                    '연말종가등락률': [(end_1980_price - start_1980_price) / start_1980_price * 100],
                    '거래대금': [data[data.index.year == 1980]['거래대금'].mean()],
                    '거래대금 등락률': [pd.NA]  # NaN으로 변경
                })

                yearly_info = pd.concat([new_row, yearly_info], ignore_index=True)

        elif market_var.get() == "KOSDAQ":
            # 1996년도 값을 수동으로 추가
            if 1996 not in yearly_info['연도'].values:
                end_1996_price = data[data.index.year == 1996]['종가'].iloc[-1]  # 1996년도 연말 종가
                start_1996_price = 1000  # 1996년도 연초 종가

                new_row = pd.DataFrame({
                    '연도': [1996],
                    '마지막 거래일': [data[data.index.year == 1996].index[-1].strftime('%Y-%m-%d')],
                    '연말종가': [end_1996_price],
                    '연말종가등락률': [(end_1996_price - start_1996_price) / start_1996_price * 100],
                    '거래대금': [data[data.index.year == 1996]['거래대금'].mean()],
                    '거래대금 등락률': [pd.NA]  # NaN으로 변경
                })

                yearly_info = pd.concat([new_row, yearly_info], ignore_index=True)

        # 열 순서 조정
        yearly_info = yearly_info[['연도', '마지막 거래일', '연말종가', '연말종가등락률', '거래대금', '거래대금 등락률']]

        # 데이터 포맷 설정
        yearly_info['거래대금'] = yearly_info['거래대금'].apply(lambda x: f"{x:,.0f}")
        if '거래대금 등락률' in yearly_info.columns:
            yearly_info['거래대금 등락률'] = yearly_info['거래대금 등락률'].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "")
        yearly_info['연말종가등락률'] = yearly_info['연말종가등락률'].apply(lambda x: f"{x:.2f}%")
        yearly_info['연말종가'] = yearly_info['연말종가'].apply(lambda x: f"{x:,.2f}")

        # 데이터 저장
        save_to_excel(yearly_info)

        # 결과 메시지
        results = []
        for index, row in yearly_info.iterrows():
            results.append(f"{row['연도']}년:\n"
                           f"마지막 거래일: {row['마지막 거래일']}\n"
                           f"연말 종가: {row['연말종가']}\n"
                           f"연말 종가 등락률: {row['연말종가등락률']}\n"
                           f"일평균 거래대금: {row['거래대금']}\n"
                           f"일평균 거래대금 등락률: {row['거래대금 등락률']}\n")
        messagebox.showinfo("연도별 정보", "\n".join(results))

    except Exception as e:
        messagebox.showerror("오류", f"데이터 처리 중 오류 발생: {e}")

# 사용자 입력을 처리하고 결과를 저장하는 함수
def process_data():
    try:
        data = get_selected_data()  # 선택된 시장 데이터를 가져옴
        threshold_value = float(entry_value.get())
        selected_criteria = criteria_var.get()
        comparison_operator = comparison_var.get()

        # '장중 가격'을 고가/저가로 처리
        if selected_criteria == '장중 가격':
            selected_criteria = '고가' if comparison_operator == '이상' else '저가'

        # 등락률 및 등락폭의 경우, raw 데이터를 사용
        elif selected_criteria == '등락률':
            selected_criteria = '등락률_raw'
        elif selected_criteria == '등락폭':
            selected_criteria = '등락폭_raw'

        # 상장시가총액과 거래대금의 경우, 숫자형 데이터로 변환
        if selected_criteria in ['상장시가총액', '거래대금']:
            if data[selected_criteria].dtype == 'object':  # 문자열일 때만 변환
                data[selected_criteria] = data[selected_criteria].str.replace(',', '').astype(float)

        # 기준에 따른 데이터 필터링
        if comparison_operator == '이상':
            filtered_data = data[data[selected_criteria] >= threshold_value]
        elif comparison_operator == '이하':
            filtered_data = data[data[selected_criteria] <= threshold_value]

        if not filtered_data.empty:
            # 필요한 컬럼만 선택해서 저장
            columns_to_save = ['종가', '등락폭', '등락률', '시가', '고가', '저가', '거래대금', '상장시가총액']
            filtered_data = filtered_data[columns_to_save]
            filtered_data['일자'] = filtered_data.index.strftime('%Y-%m-%d')  # 일자 추가
            filtered_data = filtered_data[['일자'] + columns_to_save]  # 일자를 첫 번째 열로 이동

            # 엑셀 파일로 저장할 경로 선택
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                filtered_data.to_excel(file_path, index=False)  # 인덱스는 저장하지 않음
                messagebox.showinfo("저장 완료", f"조회된 데이터가 엑셀 파일로 저장되었습니다.\n{file_path}")
        else:
            messagebox.showinfo("결과 없음", f"{selected_criteria}가 {threshold_value} {comparison_operator}인 데이터가 없습니다.")
    except ValueError:
        messagebox.showerror("입력 오류", "유효한 숫자를 입력해주세요.")
    except Exception as e:
        messagebox.showerror("오류", f"엑셀 저장 중 오류가 발생했습니다: {str(e)}")

# 연속 상승 및 하락일자 계산 함수
def calculate_consecutive_days(data):
    data['상승'] = data['종가'].diff().apply(lambda x: 1 if x > 0 else -1 if x < 0 else 0)

    # 상승과 하락 그룹 식별
    data['그룹'] = (data['상승'] != data['상승'].shift(1)).cumsum()

    # 그룹별 연속 일수 계산 및 필요한 열 추가
    result = data.groupby('그룹').agg({
        '상승': 'first',
        '종가': ['size', 'first', 'last'],
    }).reset_index(drop=True)

    # 컬럼 이름 변경
    result.columns = ['방향', '연속일수', '시작일종가', '마지막일종가']

    # 시작일자와 마감일자 추가
    result['시작일자'] = data.groupby('그룹').apply(lambda x: x.index.min().strftime('%Y-%m-%d')).reset_index(drop=True)
    result['마감일자'] = data.groupby('그룹').apply(lambda x: x.index.max().strftime('%Y-%m-%d')).reset_index(drop=True)

    # 시작 전일 종가 추가
    # 직전 거래일의 종가를 찾기 위해 shift 사용
    data['시작전일종가'] = data['종가'].shift(1)
    # 직전 거래일의 종가가 없으면 시작일 종가로 대체
    data['시작전일종가'] = data.apply(lambda row: row['종가'] if pd.isna(row['시작전일종가']) else row['시작전일종가'], axis=1)

    # 각 그룹에 대해 시작 전일 종가 추출
    result['시작전일종가'] = data.groupby('그룹').apply(lambda x: x['시작전일종가'].iloc[0]).reset_index(drop=True)

    # 연속일자 등락률 계산
    result['등락률'] = ((result['마지막일종가'] - result['시작전일종가']) / result['시작전일종가']) * 100
    result['등락률'] = result['등락률'].round(2)  # 소수점 둘째 자리까지

    # 등락률에서 NaN 값 처리
    result['등락률'].replace([float('inf'), -float('inf')], 0, inplace=True)  # 무한 값 처리
    result['등락률'].fillna(0, inplace=True)  # NaN을 0으로 대체

    # 상승 결과만 필터링 및 순위 매기기
    상승_result = result[result['방향'] > 0].sort_values(by='연속일수', ascending=False).reset_index(drop=True)
    상승_result.index += 1  # 순위 시작을 1부터

    # 컬럼 순서 변경 및 필요한 정보 포함
    상승_result = 상승_result[[
        '연속일수',
        '시작전일종가',
        '마지막일종가',
        '시작일자',
        '마감일자',
        '등락률'
    ]]

    상승_result.columns = ['연속일수', '시작전일종가', '마지막일종가', '시작일자', '마감일자', '등락률']

    상승_result = result[result['방향'] > 0].sort_values(by='연속일수', ascending=False)
    하락_result = result[result['방향'] < 0].sort_values(by='연속일수', ascending=False)

    return 상승_result, 하락_result


# GUI 추가
def show_consecutive_days_info():
    data = get_selected_data()
    상승_result, 하락_result = calculate_consecutive_days(data)

    상승_top = 상승_result.head(5)
    하락_top = 하락_result.head(5)

    result_text = "연속 상승 일자 Top 5:\n"
    for index, row in 상승_top.iterrows():
        result_text += f"연속 일수: {row['연속일수']}일 | 시작일자: {row['시작일자']} | 마감일자: {row['마감일자']}\n"

    result_text += "\n연속 하락 일자 Top 5:\n"
    for index, row in 하락_top.iterrows():
        result_text += f"연속 일수: {row['연속일수']}일 | 시작일자: {row['시작일자']} | 마감일자: {row['마감일자']}\n"

    messagebox.showinfo("연속 일자 정보", result_text)

    # 결과를 엑셀로 저장
    save_to_excel_1([상승_result, 하락_result])

# 상위/하위 10위 데이터 표시 함수
def show_top_10():
    try:
        data = get_selected_data()  # 선택된 시장 데이터를 가져옴
        selected_criteria = criteria_var.get()
        comparison_operator = comparison_var.get()

        # '장중 가격'을 고가/저가로 처리
        if selected_criteria == '장중 가격':
            selected_criteria = '고가' if comparison_operator == '이상' else '저가'

        # 등락률 및 등락폭의 경우, raw 데이터를 사용
        elif selected_criteria == '등락률':
            selected_criteria = '등락률_raw'
        elif selected_criteria == '등락폭':
            selected_criteria = '등락폭_raw'

        # 상장시가총액과 거래대금의 경우, 숫자형 데이터로 변환
        if selected_criteria in ['상장시가총액', '거래대금']:
            if data[selected_criteria].dtype == 'object':  # 문자열일 때만 변환
                data[selected_criteria] = data[selected_criteria].str.replace(',', '').astype(float)

        # 기준에 따른 상위 또는 하위 10위 데이터 가져오기
        if comparison_operator == '이상':
            top_10_data = data.nlargest(10, selected_criteria)[[selected_criteria]]
        else:
            top_10_data = data.nsmallest(10, selected_criteria)[[selected_criteria]]

        top_10_dates = top_10_data.index.strftime('%Y-%m-%d')  # 상위/하위 10위의 일자

        # 상위/하위 10위 데이터 출력
        top_10_text = "\n".join([f"{date}: {value}" for date, value in zip(top_10_dates, top_10_data[selected_criteria])])
        messagebox.showinfo("상위/하위 10위", f"{selected_criteria}의 {comparison_operator} 순위 10위 값과 일자:\n\n{top_10_text}")
    except Exception as e:
        messagebox.showerror("오류", f"상위/하위 10위 조회 중 오류가 발생했습니다: {str(e)}")

# 상위/하위 100위 데이터를 엑셀로 저장하는 함수
def save_top_100_to_excel():
    try:
        data = get_selected_data()  # 선택된 시장 데이터를 가져옴
        selected_criteria = criteria_var.get()
        comparison_operator = comparison_var.get()

        # '장중 가격'을 고가/저가로 처리
        if selected_criteria == '장중 가격':
            selected_criteria = '고가' if comparison_operator == '이상' else '저가'

        # 등락률 및 등락폭의 경우, raw 데이터를 사용
        elif selected_criteria == '등락률':
            selected_criteria = '등락률_raw'
        elif selected_criteria == '등락폭':
            selected_criteria = '등락폭_raw'

        # 상장시가총액과 거래대금의 경우, 숫자형 데이터로 변환
        if selected_criteria in ['상장시가총액', '거래대금']:
            if data[selected_criteria].dtype == 'object':  # 문자열일 때만 변환
                data[selected_criteria] = data[selected_criteria].str.replace(',', '').astype(float)

        # 기준에 따른 상위 또는 하위 100위 데이터 가져오기
        if comparison_operator == '이상':
            top_100_data = data.nlargest(100, selected_criteria)
        else:
            top_100_data = data.nsmallest(100, selected_criteria)

        # 필요한 컬럼만 선택해서 저장
        columns_to_save = ['종가', '등락폭', '등락률', '시가', '고가', '저가', '거래대금', '상장시가총액']
        top_100_data = top_100_data[columns_to_save]
        top_100_data['일자'] = top_100_data.index.strftime('%Y-%m-%d')  # 일자 추가
        top_100_data = top_100_data[['일자'] + columns_to_save]  # 일자를 첫 번째 열로 이동

        # 엑셀 파일로 저장할 경로 선택
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            top_100_data.to_excel(file_path, index=False)  # 인덱스는 저장하지 않음
            messagebox.showinfo("저장 완료", f"상위/하위 100위 데이터가 엑셀 파일로 저장되었습니다.\n{file_path}")
    except Exception as e:
        messagebox.showerror("오류", f"엑셀 저장 중 오류가 발생했습니다: {str(e)}")

# GUI 설정
root = tk.Tk()
root.title("KOSPI/KOSDAQ 데이터 필터링")

# 시장 선택
market_var = tk.StringVar(value='KOSPI')
tk.Label(root, text="시장 선택:").pack(pady=10)
tk.Radiobutton(root, text="KOSPI", variable=market_var, value='KOSPI').pack(anchor='w')
tk.Radiobutton(root, text="KOSDAQ", variable=market_var, value='KOSDAQ').pack(anchor='w')

# 기준 선택
criteria_var = tk.StringVar(value='종가')
tk.Label(root, text="기준을 선택하세요:").pack(pady=10)
tk.Radiobutton(root, text="종가", variable=criteria_var, value='종가').pack(anchor='w')
tk.Radiobutton(root, text="장중 가격", variable=criteria_var, value='장중 가격').pack(anchor='w')
tk.Radiobutton(root, text="등락률", variable=criteria_var, value='등락률').pack(anchor='w')
tk.Radiobutton(root, text="등락폭", variable=criteria_var, value='등락폭').pack(anchor='w')
tk.Radiobutton(root, text="상장시가총액", variable=criteria_var, value='상장시가총액').pack(anchor='w')
tk.Radiobutton(root, text="거래대금", variable=criteria_var, value='거래대금').pack(anchor='w')

# 기준 값 입력
tk.Label(root, text="기준 값을 입력하세요:").pack(pady=10)
entry_value = tk.Entry(root)
entry_value.pack(pady=5)

# 비교 연산자 선택
comparison_var = tk.StringVar(value='이상')
tk.Label(root, text="비교 연산자 선택:").pack(pady=10)
tk.Radiobutton(root, text="이상/상위", variable=comparison_var, value='이상').pack(anchor='w')
tk.Radiobutton(root, text="이하/하위", variable=comparison_var, value='이하').pack(anchor='w')

# 버튼 추가: 데이터 업데이트
tk.Button(root, text="코스피/코스닥 데이터 최신화", command=update_data).pack(pady=10)

# 버튼 추가: 조회 및 저장
tk.Button(root, text="기준 만족 값 모두 엑셀 저장", command=process_data).pack(pady=10)

# 버튼 추가: 조건 만족 가장 최근 일자 조회
tk.Button(root, text="조건 만족 가장 최근 일자 조회(오늘 포함)", command=show_latest_matching_info).pack(pady=10)

# 버튼 추가: 오늘 기준 조회
tk.Button(root, text="오늘 기준 값 직전 일자 조회", command=show_today_info).pack(pady=10)

# 버튼 추가: 연도별 일평균 거래대금 및 등락률 조회
tk.Button(root, text="연도별 일평균 거래대금 및 연말 종가 등락률 조회", command=show_yearly_avg_info).pack(pady=10)

# GUI에 연속 상승/하락일자 역대 순위를 조회하는 버튼 추가
tk.Button(root, text="연속 상승/하락 일자 조회", command=show_consecutive_days_info).pack(pady=10)

# 버튼 추가: 상위/하위 10위 조회
tk.Button(root, text="상위/하위 10위 조회", command=show_top_10).pack(pady=10)

# 버튼 추가: 상위/하위 100위 엑셀 저장
tk.Button(root, text="상위/하위 100위 엑셀 저장", command=save_top_100_to_excel).pack(pady=10)

# GUI 실행
root.mainloop()

