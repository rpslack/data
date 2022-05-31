from email.parser import Parser
from email import policy
from email.parser import BytesParser
import os
import email
from email.message import EmailMessage
from email.header import decode_header
import ctypes
import re

PATH_Dir = './email' #해당폴더에 원하는 eml파일이나 폴더 넣기 eml파일만 잇어야함

# 요월들을 매핑하여 date를 변환하는 함수
def convert_date(date):
    total = []
    
    con_date = date.split(' ')
    
    Month_list = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sept','Oct','Nov','Dec']
    
    for i,month in enumerate(Month_list):        
        if con_date[2] == month:
            Con_month = str(i+1)
            if len(Con_month) == 1:
                Con_month = ''.join([str(0),Con_month])
            temp = f"{con_date[3]}-{Con_month}-{con_date[1]} {con_date[4]}"
            return temp        
    
    
#정보를 텍스트로 쓰는 함수
def Result_report(result):
    with open('Status_Report.txt','a',encoding='utf-8') as f:
        f.writelines(f"{result}\n")

#html내용들을 모조리 삭제하는 함수
def HtmltoText(data):
    data = re.sub(r'&nbsp;','',data)
    data = re.sub(r'</.*?>','\n',data)
    data = re.sub(r'<.*?>', '', data)
    return data
       

#email내의 정보들을 추출하는 함수
def extract_info(target_eml):

    with open(target_eml, 'rb') as fp:
        info = []
        msg = BytesParser(policy=policy.default).parse(fp)
       
        EML_RECEIVE = str(msg['To'])   #수신자
        EML_SENDER = str(msg['From'])
        EML_SUBJECT = str(msg['Subject'])
        EML_DATE = str(msg['Date'])
        if msg['Date'] is not None:
            EML_DATE = convert_date(EML_DATE)
        info.append(f"메일 제목 : {EML_SUBJECT}\n")
        info.append(f"날짜  :   {EML_DATE}\n")
        info.append(f"수신자  :   {EML_RECEIVE}\n")
        info.append(f"발신자  :   {EML_SENDER}\n")
       
        if msg['X-Original-SENDERIP'] is not None:
            EML_SEND_IP = str(msg['X-Original-SENDERIP'])
            info.append(f"X-Original-SendIP  :   {EML_SEND_IP}\n")           

        if msg['X-Originating-IP'] is not None:
            EML_SEND_IP2 = str(msg['X-Originating-IP'])
            info.append(f"X-Originating-IP  :   {EML_SEND_IP2}\n")           
                     
        
        if msg['X-Original-SENDERCOUNTRY'] is not None:
            EML_SEND_COUNTRY = str(msg['X-Original-SENDERCOUNTRY'])
            info.append(f"X-Original-SendCOUNTRY  :   {EML_SEND_COUNTRY}\n")   
           
        try:
            for part in msg.walk():                            # walk visits message
                type = part.get_content_type()
                if type == 'text/html':
                    EML_BODY = str(msg.get_body(preferencelist=('html')).get_content())                    
                    EML_BODY = HtmltoText(EML_BODY)
                elif type == 'text/plain':
                    EML_BODY = str(msg.get_body(preferencelist=('plain')).get_content())
                    
        except Exception as Error:
            print(Error)
            pass
        
        info.append(f"이메일 내용  :   \n{EML_BODY}") # Body
        
        try:
            if "인공지능" not in EML_SENDER:
                request_list = EML_SENDER.split(' ')
                if len(request_list) > 1:
                    request_name = request_list[0]
                    request_email = request_list[1]
                    request_email = re.sub(r'<|>', '', request_email)
                else:
                    request_name = request_list[0]
                    request_email = request_list[0]
                request_time = re.sub(r':', '', EML_DATE)
            else:
                EML_BODY = EML_BODY.strip()
                EML_BODY = re.sub(r'\t', '', EML_BODY)
                body_list = list(map(str, EML_BODY.split('\n')))
                for idx, body_text in enumerate(body_list):
                    if body_text == '보낸사람':
                        request_list = body_list[idx+1].split(' ')
                        request_name = request_list[1]
                        request_name = re.sub(r'"','', request_name)
                        request_email = request_list[2]
                        request_email = re.sub(r'&lt;|&gt;', '', request_email)
                    if body_text == '날짜':
                        request_time = body_list[idx+1]
                        request_time = request_time[2:]
                        request_time = ' '.join([request_time[:10], request_time[-8:]])  
                        request_time = re.sub(r':', '', request_time)
            filename_merge = '_'.join([request_name, request_email, request_time, '.xlsx'])
        except:
            filename_merge = 'returnmail'

    return info, filename_merge


def get_part_filename(msg: EmailMessage):
    try:
        filename =  msg.get_filename()
    
        
        if decode_header(filename)[0][1] is not None:
            filename = decode_header(filename)[0][0].decode(decode_header(filename)[0][1])
            
        return filename    
    except:
        return 'File Error'   

#eml 파일내의 파일들을 추출하여 저장
def extract_attachments(Path, target_eml, filename_merge):
    try:
        msg = email.message_from_file(open(target_eml,encoding='utf-8'))
    except Exception as detail:
        print(detail)       
    attachments=msg.get_payload()    
    fnam_list = []
    
    
    if msg.is_multipart() is True:
        
        for idx, attachment in enumerate(attachments[1:]):
            if get_part_filename(attachment) == 'File Error':
                return 'File Error'
            else:
                fnam = get_part_filename(attachment)
                fnam_list.append(fnam)
                fnam = '_'.join([str(idx), filename_merge])
                attach_file = f"{Path}\{fnam}"
                with open(attach_file, 'wb') as f:
                    f.write(attachment.get_payload(decode=True))
                    Result_report(f"sucess!, extract attachment : {fnam}")
    elif msg.is_multipart() is False:
        return 'No File'       
    else:
        return 'File Error'

    return fnam_list


def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)




def main():

    for root, dirs, files in os.walk(PATH_Dir):
        for file in files:
            list_info = []
            list_Second = []            
            if '_info.txt' not in file:
                target = f"{root}\{file}"
                fnm_txt = f"{root}\{file}_info.txt"            
                list_Second, filename_merge = extract_info(target)
                list_Second = [x for x in list_Second if x]
                list_info.append(f"eml 파일명 : {file}\n")
                try:
                    fnm_name = extract_attachments(root,target,filename_merge)
                    if fnm_name == 'No File':
                        list_info.append(f"{file}   :   첨부파일 없음\n")
                        Result_report(f"{file}  : 첨부파일없음")
                    elif fnm_name == 'File Error':
                        list_info.append(f"첨부 파일 : 파일 형식 에러로 수동 추출 필요\n")   
                        Result_report(f"{file}  :   파일 형식 에러")
                    else:
                        fnm_name = [x for x in fnm_name if x]   #eml 첨부파일명들 가져온 후 list 내 None 삭제
                        for x in fnm_name:
                            #Attach_File_PATH = f"{root}\{x}"
                            list_info.append(f"첨부 파일명 : {x} \n") # 첨부파일 MD5값   :   {getHash(Attach_File_PATH)}

                        
                except Exception as err:
                    Result_report(f"{root}에서 문제 발견 : {err} ")                
                    pass               
             
                with open(fnm_txt, 'w',encoding='utf-8') as f:                                         
                    f.writelines(list_info)
                    f.writelines(list_Second)
                    Result_report(f"Sucess!, create info file : {file}_info.txt")
                    Result_report("---------------------------------------")
            else:
                Result_report(f"중복파일 발견 : {fnm_txt}")
               
  

if __name__ == "__main__":
    main()
    Mbox("완료","작업완료",0)
  
file_nm = '기업철출력(jihyeon@kosmes.or.kr).eml'
cf = os.path.join(PATH_Dir, file_nm)
os.path.isfile(cf)

attach_file = './email\정책자금·사후관리 관련서류 출력_0127..xlsx'
attach_file = ''.join([attach_file[:-5],'(','1',')',attach_file[-5:]])
