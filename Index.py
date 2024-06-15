import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import sys
import google_auth_httplib2
import httplib2
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import HttpRequest
from datetime import datetime, timedelta
import json
import ctypes
import re
import os

# --- 定義 -----------------------------------------------------------------------------------------
SCOPE = "https://www.googleapis.com/auth/spreadsheets"
# 暗号化ファイルの保存場所
ENCODEDLL_PATH = 'encode.so'

class CSpreadSheetCtrl:
    def __init__(self):
        self.sheet_id = None
        self.client_secret = None
        self.cred = None
        self.gsheet = None
        self.sheet_name = None 
    
    def set_sheet_id(self, sheet_id):
        """シートIDの設定"""
        if (sheet_id == None) or (sheet_id == ""):
            self.sheet_id = None
        else:
            self.sheet_id = sheet_id    
    
    def set_client_secret(self, client_secret):
        """JSONファイルの設定"""
        if (client_secret == None) or (client_secret == ""):
            self.client_secret = None
        else:
            self.client_secret = client_secret   
   
    def connect(self, json_data):
        """接続"""
        try:
            if (self.sheet_id == None) or (json_data == None):
                return False
            json_data = json.loads(json_data)
            self.cred = service_account.Credentials.from_service_account_info(json_data, scopes=[SCOPE])
            
            def build_request(http, *args, **kwargs):
                new_http = google_auth_httplib2.AuthorizedHttp(self.cred, http=httplib2.Http())
                return HttpRequest(new_http, *args, **kwargs)

            authorized_http = google_auth_httplib2.AuthorizedHttp(self.cred, http=httplib2.Http())
            service = build("sheets", "v4", requestBuilder=build_request, http=authorized_http)
            self.gsheet = service.spreadsheets()
            return True
        
        except Exception as e:
            print(e, file=sys.stderr)
            return False
    
    def set_data(self, id, name, age, gender, mail): # id,ageを追加済
        """データ設定"""
        try:
            # --- 認証OKかどうか -----------------
            if self.gsheet == None:
                return False
            # --- シート名の設定 -----------------
            self.sheet_name = (datetime.now()+timedelta(hours=9)).strftime('%Y%m%d')
            st.write(self.sheet_name)
            # --- シートの有無確認 & 作成 --------
            result = self.is_exist_sheet(self.sheet_name) 
            if result == False:
                # シートの作成
                result = self.make_sheet(self.sheet_name) 
                st.write(self.sheet_name)
                if result == False:
                    return False
            # --- データ登録 ---------------------
            request = self.gsheet.values().append(
                spreadsheetId=self.sheet_id,
                range=f"{self.sheet_name}!A:E", # A:Eに修正済
                body=dict(values=[[f"'{id:03}", name, age, gender, mail]]), # id,ageを追加済
                valueInputOption="USER_ENTERED",
            ).execute()
            if request:
                return True
            else:
                return False
                
        except Exception as e:
            return False

    def is_exist_sheet(self, sheet_name):
        """シートが存在するかどうか"""
        try:
            spreadsheet = self.gsheet.get(spreadsheetId=self.sheet_id).execute()
            sheet_exists = any(sheet['properties']['title'] == sheet_name for sheet in spreadsheet['sheets'])
            return sheet_exists
        
        except Exception as e:
            print(e, file=sys.stderr)
            return False        
        
    def make_sheet(self, sheet_name):
        """シートの作成"""
        try:
            request_body = {
                'requests': [
                    {
                        'addSheet': {
                            'properties': {
                                'title': sheet_name,
                                'gridProperties': {
                                    'rowCount': 1,
                                    'columnCount': 6 # 6行に修正済
                                }
                            }
                        }
                    }
                ]
            }
            result = self.gsheet.batchUpdate(spreadsheetId=self.sheet_id, body=request_body).execute()
            if result:
                # タイトル行を追加
                titles = [['ID', '名前', '年齢', '性別', 'メール']]
                body = {
                    'values': titles
                }
                self.gsheet.values().update(
                    spreadsheetId=self.sheet_id,
                    range=f"{sheet_name}!A1:E1",
                    body=body,
                    valueInputOption='RAW'
                ).execute()
                return True 
            else:
                return False 
        
        except Exception as e:
            return False        
        
    def get_data_num(self):
        """データ数の取得"""
        try:
            # --- 認証OKかどうか -----------------
            if self.gsheet == None:
                return 0
            # --- シート名の設定 -----------------
            self.sheet_name = (datetime.now()+timedelta(hours=9)).strftime('%Y%m%d')
            # シートが存在するかどうか確認
            sheet_exists = self.is_exist_sheet(self.sheet_name)
            if not sheet_exists:
                result = self.make_sheet(self.sheet_name)
                if not result:
                    return 0, None
            # シートのプロパティを取得して行数を確認
            sheet_metadata = self.gsheet.get(spreadsheetId=self.sheet_id).execute()
            sheets = sheet_metadata.get('sheets', '')
            for sheet in sheets:
                if sheet['properties']['title'] == self.sheet_name:
                    sheet_row_count = sheet['properties']['gridProperties']['rowCount']
                    if sheet_row_count == 1:
                        return 0, self.sheet_name
            # --- スプレッドシート情報の取得 ------
            range_name = f'{self.sheet_name}!A2:A'
            result = self.gsheet.values().get(spreadsheetId=self.sheet_id, range=range_name).execute() 
            values = result.get('values', []) 
            # --- idの取得 ---------------------
            if not values:
                return 0, self.sheet_name
            else:
                # 数値に変換して最大値を取得
                numeric_values = [int(item[0]) for item in values]
                if not numeric_values:
                    return 0, self.sheet_name
                max_value = max(numeric_values)
                return max_value, self.sheet_name
        
        except Exception as e:
            return 0, None
        
def main():
    # urlから暗号化されたパラメータを取得
    try:
        params = st.query_params
        encrypted_data = params["defprm"]
    except KeyError:
        return f"ERROR: {str(e)}"
    
    # 暗号化されたパラメータを復号
    with st.spinner('データを読み込んでいます...'):
        try:
            result,nameid, gmail_adr, gmail_psw, json_data = decrypt_data(encrypted_data)
            if result == None:
                return
        except Exception as e:
            return f"ERROR: {str(e)}"
    st.title("アンケートフォーム test")

    # スプレッドシートコントローラーのインスタンスを作成
    SpreadSheetCtrl = CSpreadSheetCtrl()
    sheet_id = nameid
    # クライアント秘密鍵のJSONファイルを読み込む
    SpreadSheetCtrl.set_sheet_id(sheet_id)
    SpreadSheetCtrl.set_client_secret(json_data)
    # Googleスプレッドシートへの接続
    if SpreadSheetCtrl.connect(json_data):
        last_id, sheet_name = SpreadSheetCtrl.get_data_num()
        if 'submitted' not in st.session_state:
            st.session_state['submitted'] = False
        if 'current_id' not in st.session_state:
            st.session_state['current_id'] = last_id + 1
        display_form(SpreadSheetCtrl, gmail_adr, gmail_psw)
        if st.session_state['submitted']:
            st.success('入力が完了しましたら、このタブを閉じてください')
    else:
        st.error('Google スプレッドシートへの接続に失敗しました。')

def decrypt_data(encrypted_data): 
    current_path = os.getcwd()  # カレントディレクトリのパスを取得
    full_path = os.path.join(current_path, 'Index.dat')  # ファイル名をパスに結合
    with open(full_path, 'r', encoding='utf-8') as file:  # ファイルをutf-8形式で読み込み
        read_dat = file.read().encode('utf-8')  # ファイルの内容を格納
    try: 
        # --- データ結合 --------------------------------
        combined_data = encrypted_data.encode('utf-8') + read_dat
        output_dat = ctypes.c_char_p(combined_data)
        # output_dat = ctypes.create_string_buffer(combined_data)
        # --- DLLをロード -------------------------------
        current_dir = os.path.dirname(os.path.abspath(__file__))
        so_path = os.path.join(current_dir, ENCODEDLL_PATH)
        dll = ctypes.CDLL(so_path)  # DLLをロード
        # --- データ整合性チェック(サイズ) ----------------
        try:
            dll.GetLength.argtypes = [ctypes.c_char_p, ctypes.c_long]
            dll.GetLength.restype = ctypes.c_long
        except:
            return "ERROR A10"
        ret_len1 = ctypes.c_long(0)
        ret_len2 = ctypes.c_long(0)
        ret_len3 = ctypes.c_long(0)
        ret_len4 = ctypes.c_long(0)
        ret_len1 = dll.GetLength(output_dat, ctypes.c_long(0))
        if ret_len1 == 0:
            return "ERROR A11"
        ret_len2 = dll.GetLength(output_dat, ctypes.c_long(1))
        if ret_len2 == 0:
            return "ERROR A12"
        ret_len3 = dll.GetLength(output_dat, ctypes.c_long(2))
        if ret_len3 == 0:
            return "ERROR A13"
        ret_len4 = dll.GetLength(output_dat, ctypes.c_long(3))
        if ret_len4 == 0:
            return "ERROR A14"

        # --- データ整合性チェック(内容) ------------------
        try:
            dll.DecryptString.argtypes = [ctypes.c_char_p, ctypes.c_long, ctypes.c_long, ctypes.c_long, ctypes.c_long, ctypes.c_char_p, ctypes.c_char_p, ctypes.c_char_p, ctypes.c_char_p]
            dll.DecryptString.restype = ctypes.c_long
        except:
            return "ERROR A20"
        nameid = ctypes.create_string_buffer(ret_len1+1)  # 出力バッファ
        gmail_adr = ctypes.create_string_buffer(ret_len2+1)  # 出力バッファ
        gmail_psw = ctypes.create_string_buffer(ret_len3+1)  # 出力バッファ
        json_data = ctypes.create_string_buffer(ret_len4+1)  # 出力バッファ
        dll_result = dll.DecryptString(output_dat, ret_len1, ret_len2, ret_len3, ret_len4, nameid, gmail_adr, gmail_psw, json_data)
        if dll_result != 0:
            return "ERROR A21"
        # 文字サイズまでの長さに変換
        nameid = nameid.raw[:ret_len1].decode('utf-8') 
        gmail_adr = gmail_adr.raw[:ret_len2].decode('utf-8') 
        gmail_psw = gmail_psw.raw[:ret_len3].decode('utf-8') 
        json_data = json_data.raw[:ret_len4].decode('utf-8') 
        return True,nameid, gmail_adr, gmail_psw, json_data
    except Exception as e:
        return f"ERROR: {str(e)}"

def display_form(SpreadSheetCtrl, gmail_adr, gmail_psw):  
    mail = st.text_input('メールアドレス', key="mail_input")
    if not validate_email(mail) and mail:
        st.error("無効なメールアドレスです。正しいメールアドレスを入力してください。")
        return
    number_of_people = st.number_input('入力する人数を選んでください', min_value=1, max_value=10, value=1, step=1)

    profiles = [user_form(i + 1) for i in range(int(number_of_people))]
    submit_button = st.button('送信')

    if submit_button:
        if all(profile['name'] for profile in profiles):
            results = process_form_data(profiles, mail)
            st.session_state['submitted'] = True
            for result in results['profiles']:
                success_message = f"{result['name']}さん 　ID: {result['id']:03} で受け付けました。\n\nスタッフに番号をお伝えください。"
                st.success(success_message)
                if mail:
                    print(gmail_adr)
                    print(gmail_psw)
                    mail_message = f"{result['name']}さん\n\n ID: {result['id']:03} で受け付けました。\n\n スタッフに番号をお伝えください。"
                    send_email(mail, 'IDのご連絡', mail_message, gmail_adr, gmail_psw)           
                # result.pop('id', None) 
                SpreadSheetCtrl.set_data(**result)  
                print(result) # test
        else:
            st.error('全ての名前を入力してください。')

def validate_email(email):    
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(email_regex, email) is not None

def user_form(user_number):
    name = st.text_input(f'名前 {user_number}', key=f'name_{user_number}', placeholder='必須')
    age = st.number_input(f'年齢 {user_number}', min_value=0, max_value=100, step=1, key=f'age_{user_number}', format='%d')
    age = age if age != 0 else '-'
    gender = st.selectbox(f'性別 {user_number}', ['', '男性', '女性', 'その他'], key=f'gender_{user_number}', index=0)
    gender = gender if gender else '-'

    return {'name': name, 'age': age if age != 0 else None, 'gender': gender if gender else None}

def process_form_data(profiles, mail):
    id_start = st.session_state['current_id']
    for index, profile in enumerate(profiles):
        profile['id'] = id_start + index
        profile['mail'] = mail
    st.session_state['current_id'] += len(profiles)  # 更新したIDをセッションステートに保存
    return {'profiles': profiles}

def send_email(recipient_email, subject, message, gmail_adr, gmail_psw):
    
    # MIMETextオブジェクトを作成
    msg = MIMEMultipart()
    msg['From'] = gmail_adr
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # メッセージを追加
    msg.attach(MIMEText(message, 'plain'))

    # GmailのSMTPサーバーに接続
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(gmail_adr, gmail_psw)
        server.sendmail(gmail_adr, recipient_email, msg.as_string())
        server.quit()
        st.success('メールにもIDを送信しましたので、ご確認ください。')
    except Exception as e:
        st.error('メールの送信中にエラーが発生しました。')

if __name__ == '__main__':
    main()
