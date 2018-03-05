import requests
import time
import hashlib
import random

from conn import *

class Sql_Core(object):

    def __init__(self, conn = None):
        self.conn = conn

    # get
    def get_request(self, url, payload, headers, cookies):
        try:
            r = requests.get(url, params=payload, headers=headers, cookies=cookies, timeout=30)
            return r
        except requests.HTTPError as e:
            print(e)
            return None

    # post
    def post_request(self, url, payload, headers, cookies):

        try:
            r = requests.post(url, timeout=30, data=payload, headers=headers, cookies=cookies)
            return r
        except requests.HTTPError as e:
            print(e)
            return None

    # post
    def post_request_file(self, url, payload, file, headers, cookies):
        try:
            r = requests.post(url, data=payload,files=file, headers=headers, cookies=cookies, timeout=30)
            return r
        except requests.HTTPError as e:
            print(e)
        return None

    def get_sms_number_by_phone(self, phone, behavior):
        cursor = ''
        conn = ''
        data = []
        try:
            conn = self.conn
            cursor = conn.cursor()
            sql = "SELECT code FROM `sms` WHERE `phone` =%s AND `behavior` =%s ORDER BY addtime DESC;"
            cursor.execute(sql, (phone, behavior))
            data = cursor.fetchone()
            if len(data) > 0:
                return data[0]
            else:
                return None
        except pymysql.Error as e:
            print("Mysql Error %d: %s" % (e.args[0], e.args[1]))
        finally:
            conn.commit()
            cursor.close()



    def get_user_by_phone(self, phone):
        cursor = ''
        data = []
        conn = ''
        try:
            conn = self.conn
            cursor = conn.cursor()
            sql = "SELECT id FROM `user` WHERE `phone` =%s ;"
            cursor.execute(sql, phone)
            data = cursor.fetchone()
            conn.commit()
        except pymysql.Error as e:
            print("Mysql Error %d: %s" % (e.args[0], e.args[1]))
        finally:
            cursor.close()
        if data is not None:
            return data

        return None


    @classmethod
    def get_range_str(cls):
        str_tmp = "ABCDEFGHIJKLMNOPQRSTUVWSYZabcdefghijklmnopqrstuvwsyz0987654321"
        range_list = []
        for i in range(0, 7):
            range_list.append(random.randint(0, len(str_tmp) - 1))
        str_data = ''
        for i in range(len(range_list)):
            str_data += str_tmp[range_list[i]]
        return str_data

    def create_authorization(self, auth_session, app_id, sid, token, app_secret):
        code = auth_session
        app_id = app_id
        timestamp = str(int(time.time()))
        noncestr = Sql_Core.get_range_str()
        sid = sid
        token = token

        array = ['auth_session' + code,
                 'app_id' + app_id,
                 'timestamp' + timestamp,
                 'noncestr' + noncestr,
                 'sid' + sid,
                 'token' + token]
        array.sort()

        sign = ''
        for arr in array:
            sign += arr

        sign += app_secret
        md5 = hashlib.md5(sign.encode(encoding='utf-8'))
        authorization = 'keker-auth-v1/' + code + '/' + app_id + '/' + timestamp + '/' + noncestr + '/' + sid + '/' + md5.hexdigest()
        return authorization



if __name__ == "__main__":
    pass