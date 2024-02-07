from DrissionPage import ChromiumPage
import time


def drop_down(page):
    """执行页面向下滚动的操作，实际上就等同于在浏览器的控制台上执行一串js代码"""
    for x in range(1, 8, 2):
        # 延时操作
        time.sleep(1)
        # 1/5 3/5 5/5
        j = x / 9
        # document.documentElement.scrollTop 制定滚动条位置
        # document.documentElement.scrollHeight 获取浏览器页面的最大高度
        js = 'document.documentElement.scrollTop = document.documentElement.scrollHeight * %f' % j
        # 运行 js 代码
        page.run_js(js)


import pandas as pd
from openpyxl import load_workbook
import os
import time
def save_to_xls(excel_file, rowData):

    # 如果文件不存在，创建一个空的Excel文件
    if not os.path.exists(excel_file):
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            # 创建一个空的DataFrame作为占位符，以便创建Excel文件
            empty_df = pd.DataFrame()
            empty_df.to_excel(writer, sheet_name='Sheet1', index=False)
    # 使用openpyxl打开现有的Excel文件
    wb = load_workbook(excel_file)
    ws = wb['Sheet1']
    # 在最后一行下方追加新行
    ws.append(rowData)
    # 保存Excel文件
    wb.save(excel_file)


save_path = "base_info.xlsx"
t_title = ["detail_link", "pic_url_str", "price", "bedrooms_num", "Bathrooms", "receptions", "address"]
save_to_xls(save_path, t_title)

# 创建页面对象，并启动或接管浏览器
page = ChromiumPage()

"https://www.zoopla.co.uk/for-sale/property/london/?q=london&results_sort=newest_listings&search_source=home&pn=1"
for pn in range(1, 6):
    page.get(
        'https://www.zoopla.co.uk/for-sale/property/london/?q=london&results_sort=newest_listings&search_source=home&pn=%d' % pn)
    drop_down(page)
    items = page.s_eles('xpath://*[@id="main-content"]/div[2]/div[3]/div/section/div[2]/div[3]/div')
    for item in items:
        detail_link = item.s_ele('tag:a').link
        print(detail_link)
        pic_url_str = ""
        pics = item.s_eles("tag:picture")
        for pic_item in pics:
            pic_link = pic_item.s_ele("tag:source").attr("srcset").split(" ", 1)[0].split(":p", 1)[0]
            pic_url_str = pic_url_str + pic_link + ";"
        price = item.ele("@data-testid=listing-price").text.replace("\"", "")
        if "," in price:
            price = price.replace(",", "")
        if "\"" in price:
            price = price.replace("\"", "")

        beds_info_list = []
        beds_lis = item.s_ele("tag:ul@@role=list").s_eles("tag:li")
        for bed_item in beds_lis:
            bed_inner_spans = bed_item.s_eles("tag:span")
            for bed_inner_span in bed_inner_spans:
                beds_info_list.append(bed_inner_span.text.strip())

        bedrooms_num, Bathrooms, receptions = "", "", ""
        for i in range(0, len(beds_info_list), 2):
            if beds_info_list[i] == "Bedrooms":
                bedrooms_num = beds_info_list[i + 1]
            if beds_info_list[i] == "Bathrooms":
                Bathrooms = beds_info_list[i + 1]
            if beds_info_list[i] == "Living rooms":
                receptions = beds_info_list[i + 1]
        address = item.s_ele("tag:address").text
        if "\"" in address:
            address = address.replace("\"", "")
        data_list = [detail_link, pic_url_str, price, bedrooms_num, Bathrooms, receptions, address]
        save_to_xls(save_path, data_list)
