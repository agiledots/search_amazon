import xlrd
import xlwt
import os
import requests
import re

from bs4 import BeautifulSoup


def read_xls(filename):
    # https://blog.csdn.net/wangkai_123456/article/details/50457284

    data = xlrd.open_workbook(filename)
    table = data.sheet_by_index(0) #通过索引顺序获取
    # 获取第一列的数据
    col_values = table.col_values(0)
    return col_values


def write_xls(filename, data):
    # http://www.cnblogs.com/MrLJC/p/3715783.html

    save_path = 'Excel_Workbook.xls'
    if os.path.isfile(save_path) and os.path.exists(save_path):
        os.remove(save_path)

    #
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('My Worksheet')

    font = xlwt.Font() # Create the Font
    font.name = 'Times New Roman'
    font.bold = True
    font.underline = True
    font.italic = True

    style = xlwt.XFStyle() # Create the Style
    style.font = font # Apply the Font to the Style

    # [{'barcode': '4902777026329', 'price': '￥ 2,085', 'asin': 'B014NELCZK'},
    #  {'barcode': '4903333191239', 'price': None, 'asin': 'B00KFVSFX8'}]

    for index, value in enumerate(data):
        # row, col
        # 第index行，第一列
        worksheet.write(index, 0, label = value["barcode"], style=style)
        # 第index行，第二列
        worksheet.write(index, 1, label = value["price"])
        # 第index行，第三列(价格)
        worksheet.write(index, 2, label = value["asin"])

    workbook.save(save_path)


filename = "barcode.xlsx"

# 读取数据
values = read_xls(filename)
print(values)

data = []

url = "https://www.amazon.co.jp/s/ref=nb_sb_noss?__mk_ja_JP=%E3%82%AB%E3%82%BF%E3%82%AB%E3%83%8A&url=search-alias%3Daps&field-keywords={}"
for barcode in values:
    search_url = url.format(barcode)
    #print(search_url)

    header = {'user-agent': 'Mozilla/5.0'}

    response = requests.get(search_url, headers=header)
    html = response.text
    #print(html)

    detail_url = None

    soup = BeautifulSoup(html, "html.parser")
    for link in soup.find_all('a',{'class':'a-link-normal'}):
        detail_url = link.get('href')
        print(link)
        break

    price = None
    asin_code = None
    if detail_url is not None:
        response = requests.get(detail_url, headers=header)
        detail_html = response.text

        soup = BeautifulSoup(detail_html, "html.parser")

        # 正常价格
        for link in soup.find_all(id='priceblock_ourprice'):
            price = link.get_text()
            break

        # セール特価
        if price is None:
            for link in soup.find_all(id='priceblock_dealprice'):
                price = link.get_text()
                break


        # ASIN番号
        for link in soup.find_all('div', {'class': 'pdTab'})[1:2]: # 区第二个元素
            for table in link.findAll('table'):
                # 去表格中第二个ta
                for td in table.findAll('td')[1:2]:
                    asin_code = td.getText()
                    break

        data.append({
            "barcode" : barcode,
            "price": price,
            "asin": asin_code,
        })


print(data)

# 写入数据
write_xls(filename, data)


