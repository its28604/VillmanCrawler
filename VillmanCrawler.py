import re
import requests
import urllib.parse
import openpyxl
from openpyxl.utils import get_column_letter
from lxml import etree

cpu_list = [""]

def getHtmlFrom(url):
	r = requests.get(url, verify=False)
	r.encoding = 'cp950'
	return etree.HTML(r.text)

def getPanelSize(text):
	for i in range(12, 17):
		result = re.search("(?<!\d)%d(?!\d)" % i, text)
		if result is not None:
			return "%d-inch" % i
	return None


wb = openpyxl.Workbook()
url = "https://villman.com/Category/Notebook-PCs/%s"

# brands = ["Acer", "Acer_Nitro", "Acer_Porsche_Design", "Acer_Predator", "Acer_Travelmate", "Alienware", "Apple", "ASUS", "Asus_ExpertBook", "Asus_ROG", "Asus_TUF", "CHUWI", "Dell", "Dell_Gaming", "Dell_Vostro", "Dell_XPS", "Gigabyte", "HP", "Huawei", "Intel", "Lenovo", "Lenovo_Legion", "Lenovo_Thinkpad", "Lenovo_Yoga", "LG", "Microsoft", "MSI", "MSI_Modern", "Porsche", "Samsung" ]
panel_pattern = "\d+(\.\d)?(in|-inch|-in)? (HD|FHD|IPS) ?(IPS|Touchscreen|LCD|Display)?"


page = 0
row = 0

while page < 500:

	r = requests.get(url % page, verify=False)
	html = etree.HTML(r.text)

	prod_link_list = html.xpath("*//a[@class='prod_link']")

	ws = wb.worksheets[0]

	i = get_column_letter(1)
	ws.column_dimensions[i].width = 140

	for a in prod_link_list:
		row += 1

		name = a.text
		ws.cell(row=row, column=1).value = name
		
		panel = getPanelSize(a.text)
		if panel is not None:
			ws.cell(row=row, column=2).value = panel

		# modelname_pattern = "(?<=%s).+?(?=[, |-]*%s)" % (brand, panel)
		# s = re.search(modelname_pattern, a.text)
		# modelname = s.group(0)
		# ws.cell(row=row, column=2).value = modelname
		# ws.column_dimensions[1].width = 60

	page += 50

wb.save("test.xlsx")


# page amount catch
# page_amount = int(html.xpath("*//div[@class='insider_right_t']/h2/em")[0].text.replace("\xa0", ""))

# result_list = []

# 	# html = "https://case.104.com.tw/postcase_list.cfm?cat=0&area=0&role=0&iType=2&caseno=%E5%AE%A4%E5%85%A7%E8%A8%AD%E8%A8%88&cat_s=0&money=&enddays=&orderby=0&page=2&other=&otherVal=&casetype=0&begin=0&cfrom=clist&IDNO=0"
# for i in range(1, page_amount+1):

# 	dl_list = html.xpath("*//div[@class='caselist']//dl")
# 	for dl in dl_list:
# 		result = {}
# 		# print(dl[0][0].attrib["href"])
# 		result["title"] = dl[0][0].text
# 		result["href"] = "https://case.104.com.tw/%s" % dl[0][0].attrib["href"]
# 		result["budget"] = dl[1].text
# 		result["last_online"] = dl[2][0].attrib['src']
# 		result["deadline"] = dl[3][0].attrib['src']
# 		result["views"] = "".join(dl[4].text.split())
# 		result["proposal"] = (dl[5][0].text + dl[5].text).replace("\r", "").replace("\n", "").replace(" ", "")
# 		result["context"] = dl[6].text
# 		requirement = dl[7][0].text
# 		for element in dl[7]:
# 			if element.text == requirement:
# 				requirement += ": " + element.text
# 			else:
# 				requirement += ", " + element.text
# 		result["requirement"] = requirement
# 		result_list.append(result);
# 		# print("title, %s" % result["title"])
# 		# print("href, %s" % result["href"])
# 		# print("budget, %s" % result["budget"])
# 		# print("last_online, %s" % result["last_online"])
# 		# print("deadline, %s" % result["deadline"])
# 		# print("views, %s" % result["views"])
# 		# print("proposal, %s" % result["proposal"])
# 		# print("context, %s" % result["context"])
# 		# print("requirement, %s" % result["requirement"])
# 		# print()
# 		# print()
# 		# print()
# 	# print("page %d" % i)
# 	url = url + "&page=%d" % i;

# return result_list

