#Crawle data from chemical accident database
import requests
import xlwt
from bs4 import BeautifulSoup


def trade_spider(max_pages):
    page = 1001
    workbook = xlwt.Workbook(encoding='utf-8')  #Generate utf-8 encoding workbook
    worksheet = workbook.add_sheet('data', cell_overwrite_ok= True)
    while page <= max_pages:
        url = "https://riscad.aist-riss.jp/acc/" + str(page) + "?lang=en"
        source_code = requests.get(url)
        if source_code.status_code == 200:
            plain_text = source_code.text
            soup = BeautifulSoup(plain_text, 'html.parser')
            col_num = 0
            for link in soup.findAll("span"):
                label = []
                value = []
                label1 = []
                parting = ""
                if link in soup.findAll("span", {'class': 'value finalEvent'}):
                    label1 = []
                elif link in soup.findAll("span", {'class': 'label'}):
                    label1 = [s.strip() for s in link.get_text().split("/")]
                    label1 = " / ".join(label1)
                    if label1 == "Emergency measures":
                        col_num = 18
                else:
                    value = [s.strip() for s in link.get_text().split("/")]
                    value = "/".join(value)
                    value = value.replace("\n", "&")
                    value = value.replace(",", "")
                    worksheet.write(page, col_num, value)
                    col_num += 1
            for link in soup.findAll("td"):
                table = [s.strip() for s in link.get_text().split("/")]
                worksheet.write(page, col_num, table)
                col_num += 1
            for link1 in soup.findAll("span", {'class': 'value finalEvent'}):
                label = [s.strip() for s in link1.get_text().split("/")]
                label = " / ".join(label)
                parting += label
                parting += "&"
            worksheet.write(page, col_num, parting[0:len(parting) - 1])
        else:
            break
        page += 1
        print(page)
    workbook.save('raw_data')

trade_spider(1001)