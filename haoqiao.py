# --- coding: utf-8 ---

import requests
import time
from bs4 import BeautifulSoup
import xlrd
from lxml import etree
import sys
reload(sys)
sys.setdefaultencoding('utf8')

flag = True
a = 1
b = 100
path = ''

while flag:
    try:
        s = raw_input('输入运行区间[格式(a,b)]:'.encode("gbk"))
        if s != '':
            t =eval(s)
            a = int(t[0])
            b = int(t[1])
        else:
            pass
        flag = False
    except Exception, e:
        flag = True
        print 'Retry!'

while not flag:
    try:
        s = raw_input('输入文件路径:'.encode("gbk"))
        path = s.replace('\n', '')
        flag = True
    except Exception, e:
            flag = False
            print 'Retry!'

loginurl = 'http://oa.haoqiao.cn'
codeurl = 'http://oa.haoqiao.cn/verify_code'

loginpage = requests.get(loginurl)
soup = BeautifulSoup(loginpage.text, "lxml")
src = soup.find('script').string.split('+')[1]
codeurl = codeurl + src.replace('\'','') + str(long(time.time()*100)) #from login page get verify_code url

page = requests.get(codeurl)
f = open('./code.png', 'wb')
f.write(page.content)
f.close()
code = raw_input('Input verify_code:')

'''
1)save the code.
2)login, use Session store the cookies
'''
data = {
    'username': '*****', #username and password 
    'paw': '****',
    'url': 'http://oa.haoqiao.cn/'
}
data['verify'] = str(code).replace('\n', '')
se = requests.Session()
checkurl = 'http://oa.haoqiao.cn/oa/login'
reps = se.post(checkurl, cookies =requests.utils.dict_from_cookiejar(page.cookies), data =data)
# cookies = {
#     'PHPSESSID': 'o240as035kqf6md9elq9qor9u2',
#     '_LOGINKEY_HQ' : '2495tEtLAmtPrDU0znjnKdft8qlgBjz9pbRTjc5dhIQ8CWpejL025TX9hEmemVdXmjfwvDgokdTkekddtFYBQScqKUIH5273qkXTFnYxbjR93B%2B%2F7mJGnUXwIhpr%2FZDQWMKg3J0',
#
# }

# Get the data

data = xlrd.open_workbook(path) # excel
table = data.sheet_by_index(0)
file1 = open('./1.txt', 'w')  #txt
for i in range(table.nrows)[a:b]: #duoshaot
    # h = table.cell(i, 1,).value
    city_id = region = country = 'Default'
    try:
        city_id = int(table.cell(i, 3).value)
        finalurl = 'http://oa.haoqiao.cn/backstage/city_edit?id='
        finalurl = finalurl + str(city_id)
    except Exception:
        print '#Wrong id {0} at row {1}!!!'.format(city_id, i+1)
        pass

    try:
        reps = requests.post(finalurl, cookies =se.cookies)
        selector = etree.HTML(reps.text)
        region = selector.xpath('//select[@name="region"]/option[@selected=""]')
        if len(region) >0 :
            region = region[0].text
        else:
            region = 'Region not selected.'
    except Exception:
        refion = 'Exception!'
        pass
    try:
        reps = requests.post(finalurl, cookies =se.cookies)
        selector = etree.HTML(reps.text)
        country = selector.xpath('//select[@name="country"]/option[@selected=""]')
        if len(country) >0 :
            country = country[0].text
        else:
            country = 'country not selected.'
    except Exception:
        country = 'Exception!'
        pass
    file1.write('id:{0}, region: {1}, country: {2}\n'.format(city_id, region, country))
    if (i+1)%10 == 0:
        print '{0} rows have been checked.'.format(i)

file1.close()
