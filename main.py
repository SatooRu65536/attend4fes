import binascii
import nfc
import os
import datetime
import gspread
import subprocess
from oauth2client.service_account import ServiceAccountCredentials

path = os.path.dirname(__file__)

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_secret.json', scope)
client = gspread.authorize(creds)
target_data = client.open("文化祭用")
main_sheet = target_data.worksheet("名簿")

SE_success = path + 'sounds/popi.mp3'
SE_success_add = path + 'sounds/teretere.mp3'
SE_err = path + 'sounds/err.mp3'
SE_err_write = path + 'sounds/deliriteli.mp3'

class MyCardReader(object):
    def on_connect(self, tag):
        self.idm = binascii.hexlify(tag.identifier).upper().decode('utf-8')
        record(self.idm)

    def read_id(self):
        clf = nfc.ContactlessFrontend('usb')

        try:
            clf.connect(rdwr={'on-connect': self.on_connect})
        finally:
            clf.close()


def record(idm):
    print("IDm : " + str(idm))

    dt_now = datetime.datetime.now()
    time = dt_now.strftime('%H:%M:%S')

    all_data = main_sheet.get_all_values()
    rfid_list = [d[2] for d in all_data]

    print(f'rfid_list {rfid_list}')
    if idm in rfid_list:
        subprocess.Popen(['mpg321', SE_success, '-q'])
        col = rfid_list.index(idm) + 1
        row = len(all_data[col-1]) + 1

        if col < 1:
            print('エラー: 1行目は書き込み禁止です')
            subprocess.Popen(['mpg321', SE_err, '-q'])
            return
        elif row < 5:
            print('エラー: 5列目以内は書き込み禁止です')
            subprocess.Popen(['mpg321', SE_err, '-q'])
            return

        try:
            print(f'追加書き込み中 {col}:{row}')
            main_sheet.update_cell(col, row, time)
            print('完了')
        except:
            print("エラー: スプレッドシートに追加できませんでした")
            subprocess.Popen(['mpg321', SE_err_write, '-q'])

    else:
        subprocess.Popen(['mpg321', SE_success_add, '-q'])
        col = len(rfid_list) + 1
        if 'add' in rfid_list:
            col = rfid_list.index('add') + 1
            print('add がありました')
        elif '' in rfid_list:
            col = rfid_list.index('') + 1
            print('途中に空欄がありました')

        if col == 1:
            print('エラー: 1行目は書き込み禁止です')
            subprocess.Popen(['mpg321', SE_err, '-q'])
            return

        try:
            print(f'追加書き込み中 {col}:3,5')
            main_sheet.update_cell(col, 3, idm)
            main_sheet.update_cell(col, 5, time)
            print('完了')
        except:
            print("エラー: スプレッドシートに追加できませんでした")
            subprocess.Popen(['mpg321', SE_err_write, '-q'])


if __name__ == '__main__':
    cr = MyCardReader()
    while True:
        print('⚡️ NFC checker is running!')
        cr.read_id()
