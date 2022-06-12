"""
    该文件的作用是将pdf文件转换为excel文件
    info默认不用改
    courses为你当前学期的课程，每一门都加入到元组中！
    识别到的课程顺序可能有误，但成绩不会错，不影响总分和平均分
    总分名次和平均分请自行使用excel计算
"""
import pdfplumber
import os
from openpyxl import Workbook

# 个人信息
info = (
    '学号',
    '姓名',
    '性别'
)
# 本学期的所有科目，必须要PDF完全一致，否则不能识别！
courses = (
    '形势与政策'
)

stuArr = []
stuScoreArr = []
workbook = Workbook()

def getText(fileName: str) -> str:
    """
    获取PDF文本
    :param fileName:
    :return:
    """
    with pdfplumber.open(fileName) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
        return text


def getScore(obj) -> list:
    """
    获取成绩
    :return:
    """
    global courses
    stuScoreArr = []
    resScore = list()
    for v in obj:
        for course in courses:
            # 如果在pdf文本中查到了课程
            if v.find(course) != -1:
                # Python程序设计 专基 5.0 99 网络安全 专必 4.0 96
                # 先将课程位置确定
                site = v.find(course)  # 文字位置
                courseInfo = v[site:]  # 从课程位置开始切

                t = courseInfo.split(' ')
                # 将课程和成绩截取出来
                c = t[0]  # 课程名
                try:
                    s = t[3]  # 成绩
                except:
                    s = 0
                resScore.append({
                    'course': c,
                    'score': s
                })

    # 遍历科目，主要是为了保证科目和成绩一一对应
    for c in courses:
        flag = False
        # 遍历成绩
        for s in resScore:
            # 如果成绩科目和当前科目相同
            if c == s['course']:
                stuScoreArr.append(int(s['score']))
                flag = True
                break
        if not flag:
            stuScoreArr.append('无')
    return stuScoreArr

def getInfo(obj) -> list:
    # 获取个人信息
    stuArr = []
    stu = obj[2].split(' ')
    stuArr.append(stu[1])
    stuArr.append(stu[3])
    stuArr.append(stu[5])

    return stuArr

def createExcel(res):
    global courses
    global info
    global workbook

    sheet = workbook.active
    sheet.title = '成绩分析 - By Danyhug'
    sheet.append(info + courses)

    # 先添加信息
    #print(res)
    for data in res:
        sheet.append(data)
    # 再添加课程信息
    #sheet.append()
    os.chdir('..')
    workbook.save('成绩单.xlsx')

res = []
os.chdir('pdf')
for file in os.listdir():
    # 只要是pdf文件，就进行分析
    if file.find('.pdf') != -1:
        print(file)
        text = getText(file)
        # print(text)
        # 项个数
        obj = text.split('\n')
        try:
            res.append(getInfo(obj) + getScore(obj))

            # 最终信息
            print({
                '学生信息': getInfo(obj),
                '成绩信息': getScore(obj)
            })
        except:
            print('** 文件识别出错，文件名为：', file)
        print('====================')

try:
    createExcel(res)
    print('文件保存成功！好用别忘记给个star')
except PermissionError:
    print('请关闭excel后再运行！')