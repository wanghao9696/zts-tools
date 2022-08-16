# coding=utf-8
'''
Created on 2018-06-04

@author: charley
'''

import logging
import cx_Oracle
import time
import traceback

import json
import requests

# from WindPy import w
import pandas as pd

Logger = logging.getLogger("PyWSforFZQS")


log_file_name = 'PyWS_log_%s.log' % time.strftime("%Y%m%d",time.localtime(time.time()))
Logger.setLevel(logging.DEBUG)
log_file_handler = logging.FileHandler(log_file_name)
log_file_handler.setLevel(logging.DEBUG)
log_console_handler = logging.StreamHandler()
log_console_handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s [%(funcName)s: %(filename)s,%(lineno)d]-%(levelname)s: %(message)s')
log_file_handler.setFormatter(formatter)
log_console_handler.setFormatter(formatter)

Logger.addHandler(log_file_handler)
Logger.addHandler(log_console_handler)


def update_realtime_price(mark, list, plist):
    # 现货，期货，指数，etf
    url_list = ["http://101.226.180.145:8005/spot",
                "http://101.226.180.145:8005/future",
                "http://101.226.180.145:8005/index",
                "http://101.226.180.145:8005/etf"]

    value = requests.post(url=url_list[int(mark)], data=json.dumps({"list": list, "plist": plist}))

    value_json = json.loads(value.text)
    datalists = value_json['datalist']

    print(datalists)

    pindex = datalists[0]["datetime"]

    prow = {}
    prow["low"] = datalists[0]["Low"]
    prow["high"] = datalists[0]["high"]
    prow["last"] = datalists[0]["last"]
    prow["open"] = datalists[0]["open"]

    return pindex, prow

def UpdateIndexPrice(index_id):
    Logger.info("UpdateIndexPrice invoke begin: %s" % index_id)
    try:
        dbhandle = cx_Oracle.connect('zhywpt', 'zhywpt', '10.29.180.151:2521/fzqsxt')
        dbcursor = dbhandle.cursor()
        sql_string = """SELECT T.JQDATAINDEX, T.WINDINDEX, T.INDEXNAME FROM DJJS_OTCFA_INDEX_LIST T WHERE T.INDEXID = :INDEX_ID AND ROWNUM = 1"""
        dbcursor.execute(sql_string, index_id=index_id)
        (JQ_index_id, wind_index_id, index_name) = dbcursor.fetchone()
        dbcursor.close()
        dbhandle.close()

        mark = 0  # 0:现货, 1:期货, 2:指数, 3:etf
        list = "AU9999.SGEX"
        plist = "Order_book_id,Symbol,Datetime,High,Last,Low,Open"

        pindex, prow = update_realtime_price(mark, list, plist)

        lowprice = prow['low']
        highprice = prow['high']
        closeprice = prow['close']
        openprice = prow['open']
        busidate = str(pindex)[0:10]
        dbhandle = cx_Oracle.connect('zhywpt', 'zhywpt', '10.29.180.151:2521/fzqsxt')
        dbcursor = dbhandle.cursor()
        sql_check_exist = """SELECT COUNT(1) FROM DJJS_OTCFA_PRICE_DATA T WHERE T.SECURITYID = :INDEX_ID  AND BUSIDATE = TO_DATE(:BUSIDATE,'yyyy-MM-dd') AND ROWNUM = 1"""
        dbcursor.execute(sql_check_exist, index_id=index_id, busidate=busidate)
        count = dbcursor.fetchone()[0]
        dbcursor.close()
        dbhandle.close()
        # if (str(prow).count("NaN") != 0):
        #     Logger.info("issue: updating " + index_id + "'s price at " + busidate + r" \r\n" + str(prow))
        #     continue

        if (count > 0):
            dbhandle = cx_Oracle.connect('zhywpt', 'zhywpt', '10.29.180.151:2521/fzqsxt')
            dbcursor = dbhandle.cursor()
            sql_string = """     
            DELETE FROM DJJS_OTCFA_PRICE_DATA T WHERE T.SECURITYID = :index_id  AND BUSIDATE = TO_DATE(:busidate,'yyyy-MM-dd') 
                          """
            dbcursor.execute(sql_string, index_id=index_id, busidate=busidate)
            dbcursor.close()
            dbhandle.commit()
            dbhandle.close()

        dbhandle = cx_Oracle.connect('zhywpt', 'zhywpt', '10.29.180.151:2521/fzqsxt')
        dbcursor = dbhandle.cursor()
        sql_string = """     
        insert into DJJS_OTCFA_PRICE_DATA(SECURITYID,SECURITYNAME,BUSIDATE,LOWPRICE,HIGHPRICE,CLOSEPRICE,OPENPRICE) values(:securityid,:securityname,to_date(:pricedate,'yyyy-mm-dd'),:lowprice,:highprice,:closeprice,:openprice)
                  """
        dbcursor.execute(sql_string, securityid=index_id, securityname=index_name, pricedate=busidate,
                         lowprice=lowprice, highprice=highprice, closeprice=closeprice, openprice=openprice)
        dbcursor.close()
        dbhandle.commit()
        dbhandle.close()

    except Exception as err:
        Logger.error(traceback.format_exc())
        return str(err)

    Logger.info("UpdateIndexPrice invoke end")
    return "OK"


def update_today_price():
    Logger.info("update_today_price invoke begin")
    try:
        dbhandle = cx_Oracle.connect('zhywpt', 'zhywpt', '10.29.180.151:2521/fzqsxt')
        dbcursor = dbhandle.cursor()
        sql_string = """SELECT T.INDEXID FROM ZHYWPT.DJJS_OTCFA_INDEX_LIST T WHERE T.ISMARK = '1'"""
        dbcursor.execute(sql_string)
        index_list = dbcursor.fetchall()
        dbcursor.close()
        dbhandle.close()

        for index_code in index_list:
            UpdateIndexPrice(index_code[0])

    except Exception as err:
        Logger.error(traceback.format_exc())
        return str(err)
    Logger.info("update_today_price invoke end")
    return "OK"


if __name__=='__main__':
    update_today_price()