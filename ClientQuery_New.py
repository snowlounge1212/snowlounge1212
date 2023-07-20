
# Client property query script to utilize Attend's web query
# Generate report for any matching property search

import pandas as pd
import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
import lxml
import html5lib
import datetime
import certifi
import os
import re
import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import datetime
from selenium import webdriver
import jaconv
from selenium.webdriver.support.select import Select

import sys
import ssl
import urllib.request
import time
import numpy as np

import win32com.client

import jaconv
import logging
from logging import getLogger
import argparse
import chromedriver_autoinstall

logger = getLogger(__name__)
logger.setLevel(logging.DEBUG)
logging.basicConfig(filename='Z:\\omata\\ClientQuery\\clientquery.log',
                    format='%(asctime)s : %(levelname)s - %(filename)s - %(message)s',
                    datefmt="%Y-%m-%d %H:%M:%S %z")

def for_sale_data(cat_flg):
    #driver = webdriver.Chrome('Z:\\ChromeDriver\\chromedriver.exe')
    driver = webdriver.Chrome(service=chromedriver_autoinstall.install())

    if cat_flg == "rent":
        df1 = pd.read_html("https://www.yuzawacorp.jp/lease/index.html?kind=&kind2=5&all=on", encoding='utf-8')
    elif cat_flg == "sale":
        df1 = pd.read_html("https://www.yuzawacorp.jp/buy_sell/index.html?kind=&kind2=0&all=on", encoding='utf-8')
    elif cat_flg == "salem":
        df1 = pd.read_html("https://www.yuzawacorp.jp/buy_sell/index.html?kind=&kind2=1&all=on", encoding='utf-8')
    elif cat_flg == "parking":
        df1 = pd.read_html("https://www.yuzawacorp.jp/lease/index.html?kind=&kind2=9&all=on", encoding='utf-8')
    elif cat_flg == "villa":
        df1 = pd.read_html("https://www.yuzawacorp.jp/buy_sell/index.html?kind=&kind2=2&all=on", encoding='utf-8')

    df_for_rent = df1[1].iloc[:, 2:7]

    if cat_flg == "salem":
        df_for_rent['物件名'] = df_for_rent['マンション名 交通'].apply(lambda x: x.split("\t")[0].strip())

    print(df_for_rent)

    #plink_list = []
        #setsubi_list = []
        #if cat_flg == "rent":
        #    driver.get("https://www.yuzawacorp.jp/lease/index.html?kind=&kind2=5&all=on")
        #elif cat_flg == "sale":
        #    driver.get("https://www.yuzawacorp.jp/lease/index.html?kind=&kind2=0&all=on")
        #elif cat_flg == "salem":
        #    driver.get("https://www.yuzawacorp.jp/buy_sell/index.html?kind=&kind2=1&all=on")
        #elif cat_flg == "parking":
        #    driver.get("https://www.yuzawacorp.jp/lease/index.html?kind=&kind2=9&all=on")

    #for i in range(df_for_rent.shape[0]):
    #    plink  = df_for_rent.iloc[i][0]
    #    plink_list.append(driver.find_elements_by_link_text(plink).__str__())

    if cat_flg == "rent":
        response = requests.get("https://www.yuzawacorp.jp/lease/index.html?kind=&kind2=5&all=on")
    elif cat_flg == "sale":
        response = requests.get("https://www.yuzawacorp.jp/lease/index.html?kind=&kind2=0&all=on")
    elif cat_flg == "salem":
        response = requests.get("https://www.yuzawacorp.jp/buy_sell/index.html?kind=&kind2=1&all=on")
    elif cat_flg == "parking":
        response = requests.get("https://www.yuzawacorp.jp/lease/index.html?kind=&kind2=9&all=on")
    elif cat_flg == "villa":
        response = requests.get("https://www.yuzawacorp.jp/buy_sell/index.html?kind=&kind2=2&all=on")

    soup = BeautifulSoup(response.text, 'html.parser')
    table = soup.find_all('table')[1]

    links = []
    madori = ""
    rentp = ""
    data = ""

    for tr in table.find_all("tr"):         #exctract [propertyname: property_link] for furtherer process
        try:
            if cat_flg == "rent"  :
                if tr.find_all('span')[0].text == "商談中":
                    data = "商談中" + tr.find_all('span')[1].text + re.search('\d*?\階?\w?号室', tr.find_all('a').__str__()).group() + ";" + tr.find('a')['href'].replace("..", "https://www.yuzawacorp.jp")
                elif tr.find_all('span')[0].text == "NEW":#
                    if tr.find_all('span')[1].text == "商談中":
                        data = "NEW商談中" + tr.find_all('span')[2].text + re.search('\d*?\階?\w?号室', tr.find_all('a').__str__()).group() + ";" + tr.find('a')['href'].replace("..", "https://www.yuzawacorp.jp")
                    else:
                        data  = "NEW" + tr.find_all('span')[1].text + re.search('\d*?\階?\w?号室', tr.find_all('a').__str__()).group()+ ";" + tr.find('a')['href'].replace("..", "https://www.yuzawacorp.jp")

                elif tr.find_all('span')[0].text == "値下げ":
                    data = "値下げ" + tr.find_all('span')[1].text + re.search('\d*?\階?\w?号室', tr.find_all('a').__str__()).group() + ";" + tr.find('a')['href'].replace("..", "https://www.yuzawacorp.jp")
                else:
                    data = tr.find_all('span')[0].text + re.search('\d*?\階?\w?号室', tr.find_all('a').__str__()).group()+";" + tr.find('a')['href'].replace("..", "https://www.yuzawacorp.jp")
            elif cat_flg == "salem":
                price = re.search('\d+', tr.text.strip('\n').replace(',', '').split('\n')[0]).group()
                rentp = price
                print(price)

                if tr.find_all('span')[0].text in ["商談中","NEW", "値下げ", "NEW商談中", "NEW値下げ" ]:
                    data = tr.find_all('span')[0].text + tr.find_all('span')[3].text + ";" + tr.find('a')['href'].replace("..", "https://www.yuzawacorp.jp")
                else:
                    data = tr.find_all('span')[2].text + ";" + tr.find('a')['href'].replace("..", "https://www.yuzawacorp.jp")

            elif cat_flg == "sale" or cat_flg == "villa":
                if cat_flg == "sale":
                    price = re.search('\d+', tr.text.strip('\n').replace(',', '').split('\n')[0]).group()
                else:
                    #villa_propname = tr.text[re.search('[^\\n]', tr.text).span()[0]:].split('\n\n')[0].replace('\n','')
                    #price = re.search('\d+?万円', villa_propname).group().replace('万円','')
                    #price = re.search('\d+', tr.find_all('td')[2].find_all('span')[1].text).group()
                    price = re.search('\d+\,?\d+', tr.find_all('td')[2].find_all('span')[1].text).group().replace(",","")

                rentp = int(price)
                print(rentp)
                data = tr.text.strip('\n').split('\n')[0].split(price)[0] + "; " + tr.find('a')['href'].replace("..", "https://www.yuzawacorp.jp")

            elif cat_flg == "parking":
                data = tr.text.strip('\n').split('\n')[0] + "; " + tr.find('a')['href'].replace("..", "https://www.yuzawacorp.jp")
                price = re.search('\d+?円', tr.text.strip('\n').replace(',', '').split('\n')[-1]).group().replace("円","").replace(",","")
                rentp = price
                print(price)
            mydata = data.split(";")
            print(mydata)
            salep_header = ['所在地', '地図 ストリートビュー','建物面積', '土地面積','間取り','価格万円']
            driver.get(mydata[1])
            time.sleep(1)
            if cat_flg == "salem":
                madori = driver.find_element('xpath','//*[@id="gotoTop"]/div/div/div/div[2]/article/table[2]/tbody/tr[2]/td[2]/strong/font').text

            elif cat_flg == "sale" or cat_flg == "villa":
                if cat_flg == "sale":
                    madori = driver.find_element('xpath','//*[@id="gotoTop"]/div/div/div/div[2]/article/table[2]/tbody/tr[2]/td[3]/strong/font').text
                else:
                    madori = ""
                rentp = int(price)       #time.sleep(1)

            elif cat_flg == "parking":
                rentp = driver.find_element('xpath','//*[@id="gotoTop"]/div/div/div/div[2]/article/table[2]/tbody/tr[2]/td[3]/strong').text[:-1]


            else:
                madori = driver.find_element('xpath','//*[@id="gotoTop"]/div/div/div/div[2]/article/table[2]/tbody/tr[2]/td[2]/strong/font').text
                rentp = driver.find_element('xpath','//*[@id="gotoTop"]/div/div/div/div[2]/article/table[2]/tbody/tr[2]/td[4]/strong[1]').text

            if cat_flg != "parking":
                padd = tr.text.strip('\n').split('\n')[5].split("ＪＲ")[0]
            elif cat_flg == "parking":
                padd = tr.text.strip('\n').split('\n')[4].split("ＪＲ")[0]
            if padd.find("湯沢町") == 0 and padd.find("南魚沼郡") == -1:  # Yuzawamachi
                padd = "南魚沼郡" + padd
            elif padd.find("湯沢町") == -1 and padd.find("南魚沼市") == -1:  # Minamiuonuma
                padd = "南魚沼市" + padd

            dict_list = [mydata[0], rentp, padd]

            time.sleep(1)
            try:
                if cat_flg == 'parking':
                    fcdict=dict(zip(['物件名', 'rentp','所在地'], dict_list))
                    fcdict['所在地'] = padd
                if cat_flg == 'rent' or cat_flg == "salem":
                    #driver.find_element('partial_link_text','共用設備').click()
                    driver.find_element('xpath', '//*[@id="gotoTop"]/div/div/div/div[2]/article/nav/ul/li[2]/a').click()
                    #driver.get(driver.find_element('xpath', '//*[@id="gotoTop"]/div/div/div/div[2]/article/nav/ul/li[2]/a').parent.current_url)
                    #driver.get(driver.current_url.replace('dtl', 'public').strip('&kind2=5') + '&apartment_id=7')
                    time.sleep(1)
                    df_facility = pd.read_html(driver.current_url, encoding='utf-8')[0].T
                    #if df_facility.isna().sum().sum() > 10:
                    #    driver.get(driver.current_url.replace('dtl', 'public').replace('&kind2=5', '&kind2=8'))
                    #    df_facility = pd.read_html(driver.current_url, encoding='utf-8')[0].T
                    print(df_facility)
                    df_facility['prop'] = mydata[0]
                    df_facility.columns = list(df_facility.iloc[0])
                    fcdict = dict(zip(list(df_facility.columns), list(df_facility.iloc[1])))
                elif cat_flg == "sale" or cat_flg == "villa":
                    salep_detail = driver.find_element('xpath','//*[@id="gotoTop"]/div/div/div/div[2]/article/table[2]/tbody/tr[2]').text.split("\n")

                    salep_table = []

                    for i in range(salep_detail.__len__()):
                        if i <3:
                            salep_table.append(salep_detail[i])
                        elif i == 3:
                            for ii in range(3):
                                salep_table.append(salep_detail[3].split(" ")[ii])

                    fcdict = dict(zip(salep_header, salep_table))

            except:
                fcdict = {}

            #if mydata[0] == "NEWYMハイツ205号室":
            #    print('Beep!!!')

            if mydata[0][:6] == "NEW商談中":
                p_name = mydata[0][6:]
            elif mydata[0][:3] in ["商談中","NEW","値下げ"]:
                p_name = mydata[0][3:]
            elif not mydata[0][:3] in ["商談中","NEW","値下げ"]:
                p_name = mydata[0]

            else: p_name = p_name

            #fcdict = dict(zip(list(df_facility.columns), list(df_facility.iloc[1])))
            if cat_flg == "rent" or "salem":
                fcdict.popitem()
            elif cat_flg == "sale" or "parking":
                fcdict = fcdict
                #if cat_flg == "parking":
                #    fcdict['所在地'] = padd
            #print(fcdict.keys().__len__() == fcdict.values().__len__())
            fcdict['pname'] = p_name
            fcdict['link'] = mydata[1]
            fcdict['madori'] = madori
            fcdict['rentp'] = rentp
            fcdict['price'] = rentp
            print(rentp)
            if cat_flg == "parking":
                fcdict['所在地'] = padd
                fcdict['parking'] = True

            fcdict['area'] = area_detection(fcdict['所在地'])
            #try:
            #    #fcdict['area'] = re.search('南魚沼?.',fcdict['所在地'] ).group()   # case for fcdict['所在地'] =='湯沢町湯沢湯沢1538'
            #    if re.search('南魚沼郡湯沢町?..?三国',fcdict['所在地'] ).group() == "南魚沼郡湯沢町三国" or "南魚沼郡湯沢町大字三国" :
            #        fcdict['area'] = "苗場"
            #    #elif re.search("南魚沼市", fcdict['所在地']).group() == "南魚沼市":
            #    #    fcdict['area'] = re.search("南魚沼市?.", fcdict['所在地']).group()[-2:]

            #except:
            #    if re.search("湯沢町?",fcdict['所在地']).group() == "湯沢町":
            #        mystr = re.search("湯沢町?..?..", fcdict['所在地']).group()
            #        fcdict['area'] = re.search("\D*", mystr).group()[-2:]
            #        if re.search("南魚沼市", fcdict['所在地']).group() == "南魚沼市":
            #            fcdict['area'] = re.search("南魚沼市?.", fcdict['所在地']).group()[-2:]



            print(fcdict['area'])

            try:
                if cat_flg == "rent" :
                    if q_ptype != "parking":
                        m2 = df_for_rent[df_for_rent.物件名 == p_name]['面積間取り'][0]
                    else:
                        m2 = 0
                elif cat_flg == "sale" or "salem":
                    #m2 = df_for_rent[df_for_rent['物件名価格'] == mydata[0] + price + " 万円"]['土地面積建物面積'][0]
                    m2 = 0

            except:
                m2 = 0

            fcdict['m2'] = m2
            time.sleep(2)
            print(fcdict)
            links.append(fcdict)
        except:
            print('failed to append links')

            pass

    pd.DataFrame(links).to_excel("Z:\\omata\\ClientQuery\\df_" + cat_flg.__str__() + "_master.xlsx", engine='openpyxl')
    return("data generated")

def c_query(q_cat):

    df_crq = pd.read_excel("Z:\\omata\\ClientQuery\\ClientQuery.xlsx", engine='openpyxl')

    df_rent = pd.read_excel("Z:\\omata\\ClientQuery\\df_rent_master.xlsx", engine='openpyxl')
    df_sale = pd.read_excel("Z:\\omata\\ClientQuery\\df_sale_master.xlsx", engine='openpyxl')
    df_salem  = pd.read_excel("Z:\\omata\\ClientQuery\\df_salem_master.xlsx", engine='openpyxl')
    df_parking = pd.read_excel("Z:\\omata\\ClientQuery\\df_parking_master.xlsx", engine='openpyxl')

    if q_cat == "sale":

        df_sale_q = df_crq[df_crq['区分'] == "売買"]

    elif q_cat  == "rent":  # rent includes parking

        df_sale_q = df_crq[df_crq['区分'] == "賃貸"]

    # adding address
    #df_rent['address'] = df_rent['所在地'].apply(lambda x: area_detection(x))
    df_rent['address'] = df_rent['area']
    df_sale['address'] = df_sale['area']
    df_salem['address'] = df_salem['area']
    df_parking['address'] = df_parking['area']
    #df_salem['address'] = df_salem['所在地'].apply(lambda x: "苗場" if x[:12] =="新潟県南魚沼郡湯沢町三国" else x)

    print(q_cat)

    if q_cat == "sale":
        df_sale_all = df_salem._append(df_sale)
        df_sale_all = df_sale_all._append(df_parking)

    elif q_cat  == "rent":
        df_sale_all = df_rent._append(df_parking)

    else:
        logger.info(q_cat)
        logger.info('other than "sale" or "rent" category chosen')

    df_sale_all['maxroom'] = df_sale_all['madori'].apply(lambda x: np.nan if x.__str__()[:1] == 'n' else x.__str__()[:1] )

    wb = openpyxl.Workbook()
    mybook = "Z:\\omata\\ClientQuery\\MatchDataResult" + q_cat + format(datetime.date.today(), "%Y%m%d") + ".xlsx"
    wb.save(mybook)

    for i in range(df_sale_q.shape[0]):
        print(df_sale_q.iloc[i])
        df_q = pd.DataFrame(df_sale_q.iloc[i][['名前', '区分', '地域', '物件種別', '間取り', 'こだわり','予算上限','Mail', '携帯']]).transpose()
        q_area = df_sale_q.iloc[i]['地域']     #df_sale_q.iloc[i]['地域'].__str__().replace('湯沢町',"南魚沼郡" )
        q_madori = int(df_sale_q.iloc[i]['madorino']).__str__()
        q_ptype = df_sale_q.iloc[i]['物件種別'].strip()
        print(df_sale_q.iloc[i]['予算上限'])
        if df_sale_q.iloc[i]['予算上限'].__str__() == "nan":
            q_maxp = None
        else:
            q_maxp = int(jaconv.zen2han(df_sale_q.iloc[i]['予算上限'].__str__().replace("万円",""), digit= True))
        if q_cat != "rent":
            if df_sale_all['rentp'].__contains__("万円"):
                df_sale_all['price'] = df_sale_all['rentp'].apply(lambda x: int(x.__str__().replace("万円", "").__str__().replace(",", "")))
            elif df_sale_all['rentp'].__contains__("円"):
                df_sale_all['price'] = df_sale_all['rentp'].apply(lambda x: int(x.__str__().replace("円", "").__str__().replace(",", "")))
            else:
                df_sale_all['price'] = df_sale_all['rentp'].apply(lambda x: int(x.__str__().replace("円","").__str__().replace(",",""))/10000)

        q_cwish = df_sale_q.iloc[i]['こだわり']
        print(q_cwish)
        if q_cwish.__str__() == 'nan':
            q_cwish = ""
        query_str = ""


        if '大浴場' in q_cwish:
            q_onsen = 'なし'
            query_str = "大浴場 != '" + q_onsen +  "'"     # "大浴場 != 'なし'"   gather data, except  'なし'
        if 'ペット' in q_cwish:
            q_pet = '不可'
            if query_str == "":
                query_str = "ペット !='" + q_pet + "'"     # gather data, except '不可'"
            else:
                query_str = query_str + " & ペット !='" + q_pet + "'"     # gather data, except '不可'
        if '駅近' in q_cwish:
            q_nearstation = True
        else:
            q_nearstation = False

        if '駐車場' in q_cwish:
            q_parking = 'なし'
            if query_str == "":
                query_str = "駐車場 != '" + q_parking + "'"
            else:
                query_str = query_str + " & 駐車場 != '" + q_parking + "'"      # gather data except 'なし'

        print(query_str)
        print(q_ptype)

        if q_ptype  in ['別荘','戸建']:
            df_q_1 = df_sale_all[df_sale_all['管理会社'].isna()]   # detached house without maxroom and no property management
            df_q_1 = df_q_1[df_q_1.parking.isna()]
        elif q_ptype  == "アパート":
            df_q_1 = df_sale_all[df_sale_all['その他共用施設'].isna()]
            df_q_1 = df_q_1[df_q_1.parking.isna()]                      # sales data has no 'parking'? check, this is to omit 'master parking' data
            df_q_1 = df_q_1[df_q_1.maxroom == q_madori]
        elif q_ptype == "リゾートマンション":
            df_q_1 = df_sale_all[df_sale_all['その他共用施設'].isna() == False]
            df_q_1 = df_q_1[df_q_1.parking.isna()]                      # sales data has no 'parking'? check
            df_q_1 = df_q_1[df_q_1.maxroom == q_madori]
        elif q_ptype == "駐車場":
            df_q_1 = df_sale_all[df_sale_all['parking'] == True]


        #df_q_1 = df_q_1[df_q_1.maxroom == q_madori]   # query with maxroom no
        if q_area == "苗場":
            df_q_1 = df_q_1[df_q_1.area == "苗場" ]               # Naeba area only if specified

        if query_str == "":
            df_q_2 = df_q_1
        else:
            df_q_2 = df_q_1.query(query_str)
            # Client wishes query
        print(q_maxp)

        if q_maxp == None:
            df_q_3 = df_q_2
        else:
            if q_cat == 'rent':
                df_q_3 = df_q_2[df_q_2.price.apply(lambda x: int(x[:-1].replace(',',""))) <= q_maxp*10000]
            else:
                df_q_3 = df_q_2[df_q_2.price <= q_maxp]


        #df_final = pd.DataFrame(df_q_3)

        if df_q_2.shape[0] == 0:
            df_final = pd.DataFrame(df_q_1)
        else:
            df_final = pd.DataFrame(df_q_2)
        #df_final['ekichika'] = df_final['所在地'].apply(lambda x: True if re.search("湯沢町?大字?湯沢?", x).group() == "湯沢町湯沢" else False)
        if q_nearstation == True:
            #df_final['ekichika'] = df_final['所在地'].apply(lambda x: True if x.__contains__("湯沢町大字湯沢") else False)
            #df_final['ekichika'] = df_final['所在地'].apply(lambda x: True if x.__contains__("湯沢町湯沢") else False)
            df_final = df_final[df_final['area'] == "湯沢"]

        df_final_res = pd.DataFrame(df_sale_q.iloc[i])._append(df_final)

        mysheet = df_q['名前'].values.__str__().replace("\\u3000", " ")[1:-1].strip("'") + df_q['間取り'].values.__str__().replace("\\u3000", " ").strip("[]'") + "_" + df_q['物件種別'].values.__str__().replace("\\u3000", " ")[1:-1].strip("'")
        mysheet = mysheet.__str__().replace(",", "-")
        print(mysheet)

        with pd.ExcelWriter(mybook, mode='a') as writer:
            wb.create_sheet(title = mysheet)
            df_final_res.to_excel(writer, sheet_name=mysheet)
            wb.save(mybook)


def send_mail(resfilepath, category):

    #res_list = check_resfile(resfilepath)
    outlook = win32com.client.Dispatch('Outlook.Application')

    objMail = outlook.CreateItem(0)  # MailItemオブジェクトのID

    # メールの設定
    #objMail.To = "omata-junko@yuzawacorp.jp"
    objMail.To = "info@yuzawacorp.jp"
    #objMail.cc = 'yyy@yyy.com'  # CC
    #objMail.Bcc = 'zzz@zzz.com'  # BCC

    datafileinfo = "(***お客様データの変更、削除、入力は　'\\fileserver\yuzawacorp\Omata\ClientQuery\\ClientQuery.xlsx' より行ってください　)"

    if category == "sale":
        objMail.Subject = '****** ゆざわ商事　顧客カードvs売買物件 マッチング ******　'  # Mailタイトル
        objMail.Body = 'お客様要望データよりマッチングした売買物件を抽出しました、詳細は添付ファイルをご覧ください \n ' + datafileinfo  # Mail本文

    elif category == "rent":
        objMail.Subject = '****** ゆざわ商事　顧客カードvs賃貸物件 マッチング ******　'  # Mailタイトル
        objMail.Body = 'お客様要望データよりマッチングした賃貸物件を抽出しました、詳細は添付ファイルをご覧ください \n ' + datafileinfo  # Mail本文

        #objMail.BodyFormat = 0



    if os.path.exists(resfilepath):

        objMail.Attachments.Add(resfilepath)  # 送付ファイルがある場合はファイルパスで添付

    #objMail.Display(True)  # MailItemオブジェクトを画面表示で確認する
        objMail.Send() # Mailを即時送信outlook = win32com.client.Dispatch('Outlook.Application')
        # objMail = outlook.CreateItem(0) # MailItemオブジェクトのID
        #

def area_detection(add_df):


    if add_df.__str__().__contains__("三国"):
        area_df = "苗場"
    elif add_df.__str__().__contains__("南魚沼市"):
        area_df = re.search("南魚沼市?..", add_df).group()[-2:]
    elif add_df.__str__().__contains__("湯沢町"):
        mystr = re.search("湯沢町?..?..", add_df).group()
        area_df = re.search("\D*", mystr).group()[-2:]


    return area_df

def data_join_villa(df_sale_path, df_villa_path):
    df_sale = pd.read_excel(df_sale_path)
    df_villa = pd.read_excel(df_villa_path)

    df_sale['sub_category'] = "sale"
    df_villa['sub_category'] = "villa"

    df_sale._append(df_villa).to_excel(df_sale_path, encoding='utf-8_sig')

    logging.info("villa data appended to sales data")

    dt = datetime.datetime.fromtimestamp(os.stat(df_sale_path).st_ctime)

    return "villa data append at {}".format(dt)


def main(mode):

    run_mode = mode.__str__().strip("--")
    logger.info("start logging")

    if run_mode == 'data':
        logger.info("generating master data -- rent")
        #for_sale_data('rent')
        logger.info("generating master data -- sale")
        #for_sale_data('sale')
        logger.info("generating master data -- salemansions")
        for_sale_data('salem')
        logger.info("generating master data -- parking")
        for_sale_data('parking')
        logger.info("start process: clientquery matching...")

    elif run_mode == 'villa':
        logger.info("generating master data -- villa")
        for_sale_data('villa')
        data_join_villa("Z:\\omata\\ClientQuery\\df_sale_master.xlsx", "Z:\\omata\\ClientQuery\\df_villa_master.xlsx")
        logger.info("villa data generated and added to sale data")
    else:

        q_category=run_mode

        logger.info("category : {}".format(q_category))
        if run_mode != 'data':
            c_query(run_mode)

        result_file = "Z:\\omata\\ClientQuery\\MatchDataResult" + q_category + format(datetime.date.today(), "%Y%m%d") + ".xlsx"
        logger.info("generated result file to:- {}".format(result_file))
        logger.info("sending mail to info@yuzawacorp.jp")
        send_mail(result_file, q_category)

if __name__ == '__main__':

        arg = sys.argv
        #arg = '--villa'
        #print(arg)
        #main(arg)
        print(arg[1])
        main(arg[1])

