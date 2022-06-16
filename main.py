# Coding By Danyhug at 2022-6-10 15:45:28
# 方正教务成绩查询
# https://github.com/Danyhug/zfCCJ
import time
import requests
import os

# 书写你的cookie
cookie = 'JSESSIONID=49752FE719E9B9401CD1F0EABF3A5CDC; route=6db69d195f41a086888368aa6174353b'
# 书写教学综合信息服务平台的域名，我们学校的是http://jw.xxxxx.edu.cn:8111
domain = 'http://jw..edu.cn:8111/'


try:
    os.mkdir('pdf')
except:
    pass
finally:
    os.chdir('pdf')

def query(stuID: int, **kwargs) -> bool:
    """
    查询成绩
    :param num: 功能选择，1或2
    :param stuID: 学号
    :param kwargs: 起始区间学号
    :return: 返回是否查询成功
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.5005.124 Safari/537.36 Edg/102.0.1245.41",
        'cookie': cookie
    }
    url = domain + '/bysxxcx/xscjzbdy_dyZgkydxList.html?gnmkdm=N558020&su={}'
    data = {
        'gsdygx': '12787-zw-mrgs',
        'ids': stuID,  # 学号
        'dyrq': '2002-03-28',
        'btmc': '学业成绩单',
        'cjdySzxs': 'dyrq,btmc,bwnr,dyfsdkc,xdydjksxm,bjgbdyxxkcxz'
    }

    # 拼接url
    url = url.format(stuID)
    res = requests.post(url=url, data=data, headers=headers).text
    if res.find('成功') != -1:
        # 20105010550.pdf
        downloadName = res.split('_')[1].split('#')[0]

        url = '{}/templete/scorePrint/score_{}'.format(domain, downloadName)

        fileBlob = requests.get(url, headers=headers).content
        with open('{}.pdf'.format(stuID), 'wb+') as f:
            f.write(fileBlob)
            print('写入成功')
    else:
        print('获取成绩出错', url, res)
        return False


while True:
    print('选择功能：')
    print('1. 查询某同学成绩')
    print('2. 查询某区间成绩')

    n = input('我选择功能为：')

    if n == '1':
        print('已选择查询某同学成绩')

        # 学号
        stuID = int(input('请输入该同学的学号：'))
        query(stuID=stuID)

    elif n == '2':
        print('已选择查询某区间成绩')

        start = int(input('输入区间学号开始值：'))
        end = int(input('输入区间学号结束值：')) + 1

        # 循环查询
        for id in range(start, end):
            print('当前查询的学号为', id)
            time.sleep(2)
            query(stuID=id)
