import pandas as pd
import numpy as np
import os
import re

path = './email'

df = pd.DataFrame(columns = {'기업명', '사업자등록번호', '일일여신조회 단면출력여부', '요청자', '발신메일', '요청시간'})
df = df[['기업명', '사업자등록번호', '일일여신조회 단면출력여부', '요청자', '발신메일', '요청시간']]

files = os.listdir(path)
xls_files = [f for f in files if f.endswith(".xlsx")]

for file in xls_files:
    
    #파일명에서 요청자, 발신메일, 요청시간 추칠
    info = file.split('_')
    request_user = info[1]
    request_email = info[2]
    request_time = info[3]
    request_time = ':'.join([request_time[:13],request_time[13:15],request_time[15:17]])
    
    #파일 불러오기
    fn = os.path.join(path, file)
    tmp = pd.read_excel(fn)
    
    #컬럼수 불일치 처리 (일일여신조회 없는 경우)
    if len(tmp.columns) > 3:
        tmp = tmp.iloc[:,:3]
    elif len(tmp.columns) == 2:
        tmp['일일여신조회 단면출력여부'] = '4'
    
    #컬럼명 불일치 처리
    tmp.columns = ['기업명', '사업자등록번호', '일일여신조회 단면출력여부']
    
    #공백행 처리
    tmp = tmp.dropna(thresh=2, axis = 0)
    
    #사업자등록번호 데이터 정제
    tmp['사업자등록번호'] = tmp['사업자등록번호'].replace('-', '', regex=True)
    tmp['사업자등록번호'] = tmp['사업자등록번호'].astype(str, errors='ignore')
    if len(tmp['사업자등록번호'][0]) > 10:
        tmp['사업자등록번호'] = tmp['사업자등록번호'].str[:-2]
    
    #일일여신조회 데이터 정제
    tmp['일일여신조회 단면출력여부'] = tmp['일일여신조회 단면출력여부'].fillna('')
    tmp['일일여신조회 단면출력여부'] = tmp['일일여신조회 단면출력여부'].astype(str, errors='ignore')
    if len(tmp['일일여신조회 단면출력여부'][0]) > 1:
        tmp['일일여신조회 단면출력여부'] = tmp['일일여신조회 단면출력여부'].str[:-2]
    if len(tmp) > 1 and tmp['일일여신조회 단면출력여부'][0] == '1' and tmp['일일여신조회 단면출력여부'][1] != '4':
        tmp['일일여신조회 단면출력여부'] = '1'
    elif len(tmp) > 1 and tmp['일일여신조회 단면출력여부'][0] == '4' and tmp['일일여신조회 단면출력여부'][1] == '1':
        tmp['일일여신조회 단면출력여부'] = '4'

    if len(tmp['일일여신조회 단면출력여부'][0]) > 1:
        tmp['일일여신조회 단면출력여부'] = '4'
    elif len(tmp['일일여신조회 단면출력여부'][0]) < 1:
        tmp['일일여신조회 단면출력여부'] = '4'
    
    #요청자 정보 추가
    tmp['요청자'] = request_user
    tmp['발신메일'] = request_email
    tmp['요청시간'] = request_time
    df = pd.concat([df, tmp])

#전체 인덱스 정렬 및 리셋
df = df.sort_values(by='요청시간')
df = df.reset_index(drop=True)

df.to_csv(r'd:\excels.csv', encoding='cp949')


