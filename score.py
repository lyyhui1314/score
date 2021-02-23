from selenium import webdriver
import time
from aip import AipOcr
from bs4 import BeautifulSoup
from PIL import Image
import wordcloud
import numpy as np
import collections
import jieba
from pyecharts import Line,EffectScatter,Overlap,Pie,Grid
import csv




SR = 1  #scaling ratio 缩放比例

def getCaptcha(driver):
    APP_ID = '15998954'
    API_KEY = 'AphX0AAnZaiE28CgiK2Gt2RF'
    SECRET_KEY = 'Ehs3ILAHy5rhD3GKCAsG3zNci8RTDYFX'

    # 调用远程端口
    client = AipOcr(APP_ID, API_KEY, SECRET_KEY)

    driver.save_screenshot('save.png')
    imgelement = driver.find_element_by_id('vchart')  # 定位验证码
    location = imgelement.location  # 获取验证码x,y轴坐标
    size = imgelement.size  # 获取验证码的长宽
    rangle = (int(location['x']), int(location['y']), int(location['x'] + size['width']),
              int(location['y'] + size['height']))  # 写成我们需要截取的位置坐标
    i = Image.open('save.png')  # 打开截图
    i = i.resize((int(i.size[0]/SR), int(i.size[1]/SR)), Image.ANTIALIAS)
    frame4 = i.crop(rangle)  # 使用Image的crop函数，从截图中再次截取我们需要的区域
    frame4.save('Code.png')  # 保存我们接下来的验证码图片

    # 读取图片
    file = open('Code.png', 'rb')
    content = file.read()
    file.close()

    options = {}
    options["language_type"] = "CHN_ENG"
    options["detect_direction"] = "true"
    options["detect_language"] = "true"
    options["probability"] = "true"

    #保存识别结果
    result = client.basicAccurate(content, options)
    #保存验证码识别结果
    strr = result['words_result'][0]['words']
    str = ''
    #验证码为字母和数字
    for item in strr:
        if item >= 'a' and item <= 'z' or item >= 'A' and item <= 'Z' or item >= '0' and item <= '9':
            str += item
    return str

def login(username, password):
    driver = webdriver.Chrome(executable_path=r'D:\\chromedriver.exe')

    # 打开登陆网站
    loginUrl = 'http://210.44.113.70/'
    driver.get(loginUrl)
    # driver.maximize_window()
    # 找到输入框
    name_input = driver.find_element_by_name("zjh")  # 找到用户名的框框
    pass_input = driver.find_element_by_name("mm")  # 找到输入密码的框框
    login_button = driver.find_element_by_id("btnSure")  # 找到登录按钮

    # 清空输入框后输入内容
    name_input.clear()  # 清空用户名框
    name_input.send_keys(username)  # 输入用户名
    time.sleep(0.5)
    pass_input.clear()  # 清空密码框
    pass_input.send_keys(password)  # 输入密码
    time.sleep(0.5)
    captcha = getCaptcha(driver)  # 获取验证码
    captcha_input = driver.find_element_by_name('v_yzm')  # 找到验证码框
    captcha_input.clear()  # 清空验证码输入框
    captcha_input.send_keys(captcha)  # 输入验证码
    # 点击登陆
    login_button.click()
    return driver

def is_chinese(uchar):
    """判断一个unicode是否是汉字"""
    if uchar >= u'\u4e00' and uchar <= u'\u9fa5':
        return True
    else:
        return False

def grab(driver):
    # 转到综合查询页面
    top_url = 'http://210.44.113.70/gradeLnAllAction.do?type=ln&oper=qb'
    driver.get(top_url)
    #转到按属性查询页面
    driver.find_element_by_class_name('font1').click()
    #转到iframe框架
    driver.switch_to.frame('lnsxIfra')
    #写入文件
    content = driver.page_source
    file=open('grade.html','wb')
    file.write(content.encode('utf-8'))
    file.close()
    # 读取文件
    file = open('grade.html', 'r', encoding='utf-8')
    content = file.read()
    file.close()

    soup = BeautifulSoup(content, 'html.parser')
    titleTop = soup.select('.odd')
    # 保存学科名和学科成绩，列表元素为字典
    list = []
    for item in titleTop:
        #字典键为学科名， 学科成绩，学科成绩分类
        dic = {'course': '', 'grade': '', 'stype':''}
        #获取学科名
        str = item.find_all('td')[2].string
        #去除无用字符
        strr = ''
        for i in str:
            if is_chinese(i):
                strr += i
        dic['course'] = strr
        #获取学科成绩，去除无用字符
        str = item.find_all('td')[6].find('p').string
        dic['grade'] = str[0:4]
        list.append(dic)
    return list

def arrange(dicList):
    arrList = []    #保存学科名和成绩列表，元素为字典，和各成绩分类数量列表，元素为整数
    num = [0, 0, 0, 0, 0, 0]
    for item in dicList:
        if item['grade'] == '优秀\xa0':
            item['grade'] = '95'
            item['stype'] = '优秀'
        elif item['grade'] == '良好\xa0':
            item['grade'] = '80'
            item['stype'] = '良好'
        else:
            if float(item['grade']) == 100:
                num[5] += 1
                item['stype'] = '满分'
            elif float(item['grade']) >= 80:
                num[4] += 1
                item['stype'] = '优秀'
            elif float(item['grade']) >= 70:
                num[3] += 1
                item['stype'] = '良好'
            elif float(item['grade']) >= 60:
                num[2] += 1
                item['stype'] = '及格'
            elif float(item['grade']) >= 0:
                num[1] += 1
                item['stype'] = '不及格'
            else:
                num[0] += 1
                item['stype'] = '异常'
    arrList.append(dicList)
    arrList.append(num)
    return arrList

def savaCSV(dicList):
    with open('wushuang.csv','a',encoding='utf-8',newline='') as f:
        #拿到了编辑器（笔）
        writer=csv.writer(f)
        for comment in dicList:
            writer.writerow(comment.values())
        f.close()

def wordcloud_icon(dicList):
    #选出所有的学科名
    i = 0
    grade = []
    comments = ''
    for item in dicList:
        if i < len(dicList):
            grade.append(list(dicList[i].values())[0])
            i += 1
        pass
    #将字符串连接起来
    for i in range(len(grade)):
        comments+=grade[i]
    #精准分词
    grades = []
    words = jieba.cut(comments, cut_all=False)
    #将长度大于二的保存在words[]中，并进行频率统计，统计学科起名规律
    for word in words:
        if len(word) > 2:
            grades.append(word)
    frequ = collections.Counter(grades)
    # top100=cipin.most_common(200)
    # 导入图片并处理
    image = np.array(Image.open('img1.jpg'))
    wordclouds = wordcloud.WordCloud(
        font_path='C:/Windows/Fonts/STXINGKA.TTF',
        mask=image,
        max_font_size=100
    )
    # 生成词云
    wordclouds.generate_from_frequencies(frequ)
    # 获取背景图片的配色方案
    wordColor = wordcloud.ImageColorGenerator(image)
    # 设置颜色为背景图片的颜色
    wordclouds.recolor(color_func=wordColor)
    wordclouds.to_file('word.png')
    pass

def view(numList):
    line = Line('折线图', width=2000)
    atter = ['数据异常', '不及格', '及格', '良好', '优秀', '满分']
    v1 = [numList[0], numList[1], numList[2], numList[3], numList[4], numList[5]]
    line.add('最高成绩', atter, v1,
             mark_point=['max'],  # 标点最大值
             mark_line=['average'])  # 虚线位置是平均分
    line.add('最低成绩', atter, v1,
             mark_point=['min'],
             legend_pos='20%')
    es = EffectScatter()  # 调用闪烁点
    es.add('', atter, v1, effect_scale=8)
    # 调用合并函数，再一个图表上输出
    overlop = Overlap()
    overlop.add(line)
    overlop.add(es)

    pie = Pie('饼图', title_pos='80%')
    pie.add('', atter, v1,
            radius=[60, 30],  # 控制内外半径的
            center=[65, 50],
            legend_pos='80%',
            legend_orient='vertical')
    # 显示两个示例图的调用函数，不然会覆盖
    grid = Grid()
    grid.add(overlop, grid_right='50%')
    grid.add(pie, grid_left='60%')
    grid.render('abc.html')

def main():
    #分别输入用户名和密码
    driver = login("", "")
    dicList = grab(driver)
    arrList = arrange(dicList)
    savaCSV(arrList[0])
    wordcloud_icon(arrList[0])
    view(arrList[1])

if __name__ == "__main__":
    main()