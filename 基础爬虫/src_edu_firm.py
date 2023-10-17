from bs4 import BeautifulSoup
import requests


# 获取数据页数
response = requests.get('https://src.sjtu.edu.cn/list/firm/new/')
soup = BeautifulSoup(response.text, 'html.parser')
# 倒数第三个 <a> 标签为总页数
page_num = int(soup.find_all('a')[-3].text)
print(f"# [Info] 共 {page_num} 页数据\n")
index = 0

with open("教育漏洞报告平台_授权学校列表.txt", 'w', encoding='utf-8') as f:
    for i in range(1, page_num + 1):
        response = requests.get(f'https://src.sjtu.edu.cn/list/firm/new/?page={i}')
        if response.status_code == 200:
            print(f"\n# [Info] 正在爬取第 {i} 页数据")
            soup = BeautifulSoup(response.text, 'html.parser')
            # 找到包含学校信息的表格
            school_table = soup.find('table', class_='am-table')
            # 遍历表格中的每一行（跳过表头）
            for row in school_table.find_all('tr', class_='row'):
                # 提取学校名称
                index += 1
                school_name = row.find('a').text
                print(index, school_name)
                f.write(str(index) + "\t" + school_name + "\n")

print(f"\n# [Info] 共 {index} 所授权学校")
