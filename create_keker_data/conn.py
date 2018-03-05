import pymysql

def get_conn_keke():
    conn= pymysql.connect(
        host='139.196.44.210',
        port = 3306,
        user='root',
        passwd='jooxTV818H',
        db ='db_keker',
        charset='utf8',
        )
    return conn