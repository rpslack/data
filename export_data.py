import pandas as pd
import openpyxl
import datetime
import cx_Oracle
import pandas as pd
from DB_Reader import Table_GET
from DB_Reader import Table_GET_TOP100
from DB_Reader import Table_Columns_GET
from DB_Reader import All_Tables
import datetime
import numpy as np

LOCATION = r"d:\instantclient_21_3"
db = cx_Oracle.connect('     ', '          ', '               ')

cursor = db.cursor()



df = pd.read_excel(r'd:\수출마케팅 지원사업 데이터(2021).xlsx', header = 0)
df = df.drop([8154])
df['사업자번호'] = df['사업자번호'].astype(str)
df['사업자번호'].replace('(.*).0', r'\1', regex=True, inplace=True)

df2 = pd.read_excel(r'd:\수출바우처 데이터(2021).xlsx', header = 0)
df2 = df2.drop([3475])
df2['사업자번호'] = df2['사업자번호'].astype(str)
df2['사업자번호'].replace('(.*).0', r'\1', regex=True, inplace=True)

df['사업연도'] = int(2021)
df2['지원연도'] = int(2021)

tdf_c = pd.DataFrame(columns = ['지원년도', '업종', '소재지', '종업원수', '업력',
                              '지원전년도 자산(백만원)', '지원전년도 부채(백만원)', '지원전년도 자본(백만원)', '지원전년도 매출(백만원)', '지원전년도 영업이익(백만원)', '지원전년도 순이익(백만원)',
                              '지원년도 자산(백만원)',  '지원년도 부채(백만원)', '지원년도 자본(백만원)','지원년도 매출(백만원)', '지원년도 영업이익(백만원)', '지원년도 순이익(백만원)'
                               , '기업번호'])

df_1 = df[['사업자번호', '지원사업명']]
df_2 = df2[['사업자번호', '지원사업명']]

tdf = pd.concat([df_1, df_2], axis=0, ignore_index=True)
tdf = pd.concat([tdf, tdf_c], axis=1, ignore_index=True)

tdf.columns = ['사업자번호', '지원사업명', '지원년도', '업종', '소재지', '종업원수', '업력',
                '지원전년도 자산(백만원)', '지원전년도 부채(백만원)', '지원전년도 자본(백만원)', '지원전년도 매출(백만원)', '지원전년도 영업이익(백만원)', '지원전년도 순이익(백만원)',
                '지원년도 자산(백만원)',  '지원년도 부채(백만원)', '지원년도 자본(백만원)','지원년도 매출(백만원)', '지원년도 영업이익(백만원)', '지원년도 순이익(백만원)'
              , '기업번호']

tdf['지원년도'] = 2021
tdf['최근결산'] = ''

depth1 = pd.read_csv(r'd:\표준산업분류코드(중분류).csv', encoding='cp949')
depth1['코드'] = depth1['코드'].apply(lambda x: f'{x:02d}')
dic_depth1 = dict(depth1.values)
depth1 = depth1.set_index('코드')
df = pd.read_csv(r'd:\공단지원_수출.csv', encoding = 'cp949')

df.rename(columns = {'업종' : '업종코드'}, inplace=True)
df['업종'] = ''
df.drop('Unnamed: 0', axis = 1)
df.replace(' ', float('nan'), inplace=True)

for i, row in df.iterrows():
    if pd.isna(row['업종코드']):
        pass
    else:
        if int(row['업종코드'][:2]) >= 1:
            df.loc[i, '업종'] = depth1.loc[row['업종코드'][:2], '항목명']

df2 = df[df.duplicated(['사업자번호'])==True]

for i, row in df2.iterrows():
    idx = tdf[tdf['사업자번호'] == str(row['사업자번호'])].index
    tdf.loc[idx, '업종'] = row['업종']
    tdf.loc[idx, '소재지'] = row['소재지']
    tdf.loc[idx, '종업원수'] = row['종업원수']
    tdf.loc[idx, '업력'] = row['업력']
    tdf.loc[idx, '기업번호'] = row['기업번호']

df = pd.read_csv(r'd:\재무제표_통합정보.csv')

for i, row in df.iterrows():
    idx = tdf[tdf['기업번호'] == row['Code']].index
    for j in list(idx):
        if row['Year'] == 2019:
            tdf.loc[j, '지원전년도 자산(백만원)'] = row['자산'] // 1000000
            tdf.loc[j, '지원전년도 부채(백만원)'] = row['부채'] // 1000000
            tdf.loc[j, '지원전년도 자본(백만원)'] = row['자본'] // 1000000
            tdf.loc[j, '지원전년도 매출(백만원)'] = row['매출'] // 1000000
            tdf.loc[j, '지원전년도 영업이익(백만원)'] = row['영업이익'] // 1000000
            tdf.loc[j, '지원전년도 순이익(백만원)'] = row['순이익'] // 1000000
            tdf.loc[j, '최근결산'] = 2019
        elif row['Year'] == 2020:
            tdf.loc[j, '지원년도 자산(백만원)'] = row['자산'] // 1000000
            tdf.loc[j, '지원년도 부채(백만원)'] = row['부채'] // 1000000
            tdf.loc[j, '지원년도 자본(백만원)'] = row['자본'] // 1000000
            tdf.loc[j, '지원년도 매출(백만원)'] = row['매출'] // 1000000
            tdf.loc[j, '지원년도 영업이익(백만원)'] = row['영업이익'] // 1000000
            tdf.loc[j, '지원년도 순이익(백만원)'] = row['순이익'] // 1000000
            tdf.loc[j, '최근결산'] = 2020
        
df2 = pd.read_csv(r'd:\재무정보_마스터.csv', encoding='cp949')

df3 = df2[['DATA2', 'STAC_YR', 'ETR_CUST_NO', 'IND_PRDT_CD', 'EMPL_CNT', 'FNDT_DT', '지역구분']].copy()
df3['STAC_YR'] = df3['STAC_YR'].fillna(0)
df3['STAC_YR'] = df3['STAC_YR'].astype('int')
df3 = df3.sort_values(by=['STAC_YR'], axis=0, ascending=False)
df3 = df3[df3['STAC_YR'] > 0]
df3.reset_index(drop=True, inplace=True)
df_info = pd.DataFrame(columns = ['사업자번호', '기업번호', '업종', '종업원수', '설립일', '소재지'])
df_info['사업자번호'] = list(df3['DATA2'].unique())
df_info[df_info['사업자번호'] == '서울']

for i, row in df_info.iterrows():
    tmp_dic = dict()
    tmp_df = df3[df3['DATA2'] == row['사업자번호']]
    for j in list(tmp_df['IND_PRDT_CD']):
        tmp_dic['업종'] = ''
        if len(str(j)) > 2:
            tmp_dic['업종'] = j
            break
    for j in list(tmp_df['EMPL_CNT']):
        tmp_dic['종업원수'] = ''
        if j >= 0:
            tmp_dic['종업원수'] = j
            break
    for j in list(tmp_df['FNDT_DT']):
        tmp_dic['설립일'] = ''
        if j >= 0:
            tmp_dic['설립일'] = j
            break
    for j in list(tmp_df['지역구분']):
        tmp_dic['소재지'] = ''
        if len(str(j)) >= 0:
            tmp_dic['소재지'] = j
            break
    df_info.loc[i, '업종'] = tmp_dic['업종']
    df_info.loc[i, '종업원수'] = tmp_dic['종업원수']
    df_info.loc[i, '설립일'] = tmp_dic['설립일']
    df_info.loc[i, '소재지'] = tmp_dic['소재지']        

df_info['CODE'] = ''

for i, row in df_info.iterrows():
    if pd.isna(row['업종']) or row['업종'] == '' or row['업종'][:2] == '80':
        pass
    else:
        if int(row['업종'][:2]) >= 1:
            df_info.loc[i, 'CODE'] = depth1.loc[row['업종'][:2], '항목명']

df_info['업력'] = ''

df_info['설립일'] = df_info['설립일'].astype('str')
df_info['설립일'] = df_info['설립일'].apply(lambda x: x[:-2])
df_info['설립일'] = df_info['설립일'].astype('datetime64[ns]')
df_info.set_index('사업자번호', inplace=True)

t1 = pd.Timestamp.now()
for i, row in tdf.iterrows():
    if int(row['사업자번호']) in list(df_info.index):
        tdf.loc[i, '업종'] = df_info.loc[int(row['사업자번호']), 'CODE']
        tdf.loc[i, '소재지'] = df_info.loc[int(row['사업자번호']), '소재지']
        tdf.loc[i, '종업원수'] = df_info.loc[int(row['사업자번호']), '종업원수']
        tdf.loc[i, '업력'] = pd.Timedelta(t1 - df_info.loc[int(row['사업자번호']), '설립일']) / np.timedelta64(1, 'Y')

tdf.to_csv(r'd:\수출지원사업_220412.csv', encoding='cp949')
tdf = pd.read_csv(r'd:\수출지원사업_220412.csv', encoding='cp949')

fin = pd.read_csv(r'd:\요약재무_goscrap.csv', encoding='cp949')
fin = fin.drop(fin.columns[0], axis=1)

fin['기준연도(Y-3)'] = fin['기준연도(Y-3)'].replace('-', '0')
fin['기준연도(Y-2)'] = fin['기준연도(Y-2)'].replace('-', '0')
fin['기준연도(Y-1)'] = fin['기준연도(Y-1)'].replace('-', '0')

fin['기준연도(Y-3)'] = fin['기준연도(Y-3)'].fillna('0')
fin['기준연도(Y-2)'] = fin['기준연도(Y-2)'].fillna('0')
fin['기준연도(Y-1)'] = fin['기준연도(Y-1)'].fillna('0')

fin['기준연도(Y-3)'] = fin['기준연도(Y-3)'].astype('int')
fin['기준연도(Y-2)'] = fin['기준연도(Y-2)'].astype('int')
fin['기준연도(Y-1)'] = fin['기준연도(Y-1)'].astype('int')

fin = fin[fin['조회 상태'] == '정상']

fin.replace('-', '0', inplace=True)

fin_df = pd.DataFrame(columns = ['사업자번호', '재무년도', '자산(백만원)', '부채(백만원)', '자본(백만원)',
                                 '매출(백만원)', '영업이익(백만원)', '순이익(백만원)'])

for i, row in fin.iterrows():
    if row['기준연도(Y-3)'] == 2019:
        fin_df = fin_df.append({'사업자번호' : row['사업자번호'],
                                '재무년도' : 2019,
                                '자산(백만원)' : row['자산총계(Y-3)'],
                                '부채(백만원)' : row['부채총계(Y-3)'],
                                '자본(백만원)' : row['자본총계(Y-3)'],
                                '매출(백만원)' : row['매출액(Y-3)'],
                                '영업이익(백만원)' : row['영업이익(Y-3)'],
                                '순이익(백만원)' : row['당기순이익(Y-3)']}, ignore_index=True)
    elif row['기준연도(Y-2)'] == 2019:
        fin_df = fin_df.append({'사업자번호' : row['사업자번호'],
                                '재무년도' : 2019,
                                '자산(백만원)' : row['자산총계(Y-2)'],
                                '부채(백만원)' : row['부채총계(Y-2)'],
                                '자본(백만원)' : row['자본총계(Y-2)'],
                                '매출(백만원)' : row['매출액(Y-2)'],
                                '영업이익(백만원)' : row['영업이익(Y-2)'],
                                '순이익(백만원)' : row['당기순이익(Y-2)']}, ignore_index=True)
    elif row['기준연도(Y-1)'] == 2019:
        fin_df = fin_df.append({'사업자번호' : row['사업자번호'],
                                '재무년도' : 2019,
                                '자산(백만원)' : row['자산총계(Y-1)'],
                                '부채(백만원)' : row['부채총계(Y-1)'],
                                '자본(백만원)' : row['자본총계(Y-1)'],
                                '매출(백만원)' : row['매출액(Y-1)'],
                                '영업이익(백만원)' : row['영업이익(Y-1)'],
                                '순이익(백만원)' : row['당기순이익(Y-1)']}, ignore_index=True)
    elif row['기준연도(Y-3)'] == 2020:
        fin_df = fin_df.append({'사업자번호' : row['사업자번호'],
                                '재무년도' : 2020,
                                '자산(백만원)' : row['자산총계(Y-3)'],
                                '부채(백만원)' : row['부채총계(Y-3)'],
                                '자본(백만원)' : row['자본총계(Y-3)'],
                                '매출(백만원)' : row['매출액(Y-3)'],
                                '영업이익(백만원)' : row['영업이익(Y-3)'],
                                '순이익(백만원)' : row['당기순이익(Y-3)']}, ignore_index=True)
    elif row['기준연도(Y-2)'] == 2020:
        fin_df = fin_df.append({'사업자번호' : row['사업자번호'],
                                '재무년도' : 2020,
                                '자산(백만원)' : row['자산총계(Y-2)'],
                                '부채(백만원)' : row['부채총계(Y-2)'],
                                '자본(백만원)' : row['자본총계(Y-2)'],
                                '매출(백만원)' : row['매출액(Y-2)'],
                                '영업이익(백만원)' : row['영업이익(Y-2)'],
                                '순이익(백만원)' : row['당기순이익(Y-2)']}, ignore_index=True)
    elif row['기준연도(Y-1)'] == 2020:
        fin_df = fin_df.append({'사업자번호' : row['사업자번호'],
                                '재무년도' : 2020,
                                '자산(백만원)' : row['자산총계(Y-1)'],
                                '부채(백만원)' : row['부채총계(Y-1)'],
                                '자본(백만원)' : row['자본총계(Y-1)'],
                                '매출(백만원)' : row['매출액(Y-1)'],
                                '영업이익(백만원)' : row['영업이익(Y-1)'],
                                '순이익(백만원)' : row['당기순이익(Y-1)']}, ignore_index=True)


fin_df.replace('-', 0, inplace=True)
fin_df = fin_df.replace(',', '', regex=True)
fin_df[['자산(백만원)', '부채(백만원)', '자본(백만원)', '매출(백만원)', '영업이익(백만원)', '순이익(백만원)']] = fin_df[['자산(백만원)', '부채(백만원)', '자본(백만원)', '매출(백만원)', '영업이익(백만원)', '순이익(백만원)']].astype(float)

for i, row in fin_df.iterrows():
    idx = tdf[tdf['사업자번호'] == row['사업자번호']].index
    for j in list(idx):
        if row['재무년도'] == 2019:
            tdf.loc[j, '지원전년도 자산(백만원)'] = row['자산(백만원)']
            tdf.loc[j, '지원전년도 부채(백만원)'] = row['부채(백만원)']
            tdf.loc[j, '지원전년도 자본(백만원)'] = row['자본(백만원)']
            tdf.loc[j, '지원전년도 매출(백만원)'] = row['매출(백만원)']
            tdf.loc[j, '지원전년도 영업이익(백만원)'] = row['영업이익(백만원)']
            tdf.loc[j, '지원전년도 순이익(백만원)'] = row['순이익(백만원)']
            tdf.loc[j, '최근결산'] = 2020
        elif row['재무년도'] == 2020:
            tdf.loc[j, '지원년도 자산(백만원)'] = row['자산(백만원)']
            tdf.loc[j, '지원년도 부채(백만원)'] = row['부채(백만원)']
            tdf.loc[j, '지원년도 자본(백만원)'] = row['자본(백만원)']
            tdf.loc[j, '지원년도 매출(백만원)'] = row['매출(백만원)']
            tdf.loc[j, '지원년도 영업이익(백만원)'] = row['영업이익(백만원)']
            tdf.loc[j, '지원년도 순이익(백만원)'] = row['순이익(백만원)']
            tdf.loc[j, '최근결산'] = 2019

tdf[tdf['사업자번호']==1010627959]

biz_info = pd.read_csv(r'd:\기업정보_goscrap.csv', encoding='cp949')
biz_info = biz_info.drop(biz_info.columns[0], axis=1)

biz_info = biz_info[['입력번호', '종업원수', '업종코드', '주소', '설립일자']]

biz_info.replace('업종', float('NaN'), inplace=True)
biz_info['설립일자'].replace('-', '', inplace=True)

biz_info = biz_info.fillna('')

biz_info['업종'] = ''
biz_info['소재지'] = ''
biz_info['업력'] = ''

biz_info = biz_info[biz_info['주소'] != '']

biz_info['설립일자'] = biz_info['설립일자'].astype('datetime64[ns]')

t1 = pd.Timestamp.now()
for i, row in biz_info.iterrows():
    biz_info.loc[i, '소재지'] = row['주소'][:2]
    biz_info.loc[i, '업력'] = pd.Timedelta(t1 - row['설립일자']) / np.timedelta64(1, 'Y')
    if len(row['업종코드']) > 2:
        biz_info.loc[i, '업종'] = depth1.loc[row['업종코드'][1:3], '항목명']

biz_info.head()

for i, row in biz_info.iterrows():
    idx = tdf[tdf['사업자번호'] == row['입력번호']].index
    for j in list(idx):
        if tdf.loc[j, '업종'] == float('NaN'):
            tdf.loc[j, '업종'] = row['업종']
        if tdf.loc[j, '소재지'] == float('NaN'):
            tdf.loc[j, '소재지'] = row['소재지']
        if tdf.loc[j, '종업원수'] == float('NaN'):
            tdf.loc[j, '종업원수'] = row['종업원수']
        if tdf.loc[j, '업력'] == float('NaN'):
            tdf.loc[j, '업력'] = row['업력']

tdf = tdf.drop('Unnamed: 0', axis=1)

tdf = tdf.drop('기업번호', axis=1)
tdf = tdf.drop('최근결산', axis=1)

tdf = tdf.sort_values(by=['지원사업명'], axis=0, ascending=True)

tdf.reset_index(drop=True, inplace=True)

tdf['사업자번호'] = tdf['사업자번호'].astype(str)

tdf['종업원수'] = tdf['종업원수'].map('{:,.0f}'.format)
tdf['업력'] = tdf['업력'].map('{:,.0f}'.format)

for i, row in tdf.iterrows():
    tdf.loc[i, '사업자번호'] = row['사업자번호'][:3]+'*******'

tdf[(tdf['소재지'].isna())&(tdf['지원전년도 자산(백만원)'].isna())&(tdf['업종'].isna())]

tdf.to_csv(r'd:\수출지원사업_220413.csv', encoding='cp949')

tdf['소재지'].unique()

tdf2 =tdf[(~tdf['소재지'].isna())&(~tdf['업종'].isna())]

tdf2.reset_index(drop=True, inplace=True)

tdf2['업력'].replace('nan', '', inplace=True)

tdf2.to_csv(r'd:\수출지원사업_220414.csv', encoding='cp949')






