import re
import requests
import urllib.parse
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from lxml import etree

def getHtmlFrom(url):
	r = requests.get(url, verify=False)
	r.encoding = 'cp950'
	return etree.HTML(r.text)

WORK_BOOK_NAME = "test.xlsx"

wb = load_workbook(WORK_BOOK_NAME)

url = "https://villman.com"
url_page = url + "/Category/Notebook-PCs/%s"
ws = wb.worksheets[0]

row = 0
exist_datas = []
while (True):
	row += 1
	name = ws.cell(row=row, column=1).value
	if name == None:
		break;
	else:
		exist_datas.append(name)

for data in exist_datas:
	print(data)

page = 0
row = 0

# while page < 500:

# 	r = requests.get(url_page % page, verify=False)
# 	html = etree.HTML(r.text)

# 	prod_link_list = html.xpath("*//a[@class='prod_link']")


# 	i = get_column_letter(1)
# 	ws.column_dimensions[i].width = 140

# 	for a in prod_link_list:
# 		row += 1

# 		name = a.text
# 		ws.cell(row=row, column=1).value = name
# 		prod_link = a.get("href")
# 		ws.cell(row=row, column=2).value = '=HYPERLINK("{}", "{}")'.format(url + prod_link, "Link")

# 	page += 50

# wb.save(WORK_BOOK_NAME)


# # page amount catch
# # page_amount = int(html.xpath("*//div[@class='insider_right_t']/h2/em")[0].text.replace("\xa0", ""))

# # result_list = []

# # 	# html = "https://case.104.com.tw/postcase_list.cfm?cat=0&area=0&role=0&iType=2&caseno=%E5%AE%A4%E5%85%A7%E8%A8%AD%E8%A8%88&cat_s=0&money=&enddays=&orderby=0&page=2&other=&otherVal=&casetype=0&begin=0&cfrom=clist&IDNO=0"
# # for i in range(1, page_amount+1):

# # 	dl_list = html.xpath("*//div[@class='caselist']//dl")
# # 	for dl in dl_list:
# # 		result = {}
# # 		# print(dl[0][0].attrib["href"])
# # 		result["title"] = dl[0][0].text
# # 		result["href"] = "https://case.104.com.tw/%s" % dl[0][0].attrib["href"]
# # 		result["budget"] = dl[1].text
# # 		result["last_online"] = dl[2][0].attrib['src']
# # 		result["deadline"] = dl[3][0].attrib['src']
# # 		result["views"] = "".join(dl[4].text.split())
# # 		result["proposal"] = (dl[5][0].text + dl[5].text).replace("\r", "").replace("\n", "").replace(" ", "")
# # 		result["context"] = dl[6].text
# # 		requirement = dl[7][0].text
# # 		for element in dl[7]:
# # 			if element.text == requirement:
# # 				requirement += ": " + element.text
# # 			else:
# # 				requirement += ", " + element.text
# # 		result["requirement"] = requirement
# # 		result_list.append(result);
# # 		# print("title, %s" % result["title"])
# # 		# print("href, %s" % result["href"])
# # 		# print("budget, %s" % result["budget"])
# # 		# print("last_online, %s" % result["last_online"])
# # 		# print("deadline, %s" % result["deadline"])
# # 		# print("views, %s" % result["views"])
# # 		# print("proposal, %s" % result["proposal"])
# # 		# print("context, %s" % result["context"])
# # 		# print("requirement, %s" % result["requirement"])
# # 		# print()
# # 		# print()
# # 		# print()
# # 	# print("page %d" % i)
# # 	url = url + "&page=%d" % i;

# # return result_list

