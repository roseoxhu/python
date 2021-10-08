#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
各城市每日推广数据统计排行
统计指标：
总注册量 totalRegist
下载登录 totalDownload
下载率   downloadRate
订单数（ToC/ToB）orderCount,order2BCount
客服代注册   adminRegist
工程师邀请   engineerInvite
用户邀请     userInvite
推广注册合计 totalInvite
当日排行

大城市：北京,上海,深圳,广州,武汉,天津,沈阳,厦门,西安,重庆,成都,贵阳
'''

import os, sys, time

import logging
import json
import random
import uuid

# pip install -i https://pypi.douban.com/simple/ jinja2
# https://pypi.org/project/Jinja2/
from jinja2 import Environment, FileSystemLoader

# pip install imgkit
# pip install pdfkit
# yum -y install wkhtmltopdf
# https://wkhtmltopdf.org/downloads.html
# https://www.cnpython.com/pypi/imgkit
import imgkit

# pip install -i http://pypi.douban.com/simple/ --trusted-host pypi.douban.com requests
import requests
# pip install --trusted-host pypi.douban.com qiniu
# Successfully installed qiniu-7.5.0
import qiniu


class Stats2img():
    # bigCities = ("北京", "上海", "深圳")
    bigCities = ("北京", "上海", "深圳", "广州", "武汉", "天津", "沈阳", "厦门", "西安", "重庆","成都", "贵阳")
    
    # statIndicators = ('totalRegist','orderCount')
    statIndicators = ('totalRegist','totalDownload','downloadRate','orderCount','order2BCount','adminRegist','engineerInvite','userInvite','totalInvite')
    
    # imgkit 可执行文件，需要事先下载安装
    # 根据需要换成Linux 可执行文件路径 /path/to/imgkit/wkhtmltox/bin/wkhtmltoimage
    wkhtmltoimage = 'D:\\imgkit\\wkhtmltox\\bin\\wkhtmltoimage.exe'
    wkhtmltopdf = 'D:\\imgkit\\wkhtmltox\\bin\\wkhtmltopdf.exe'

    def __init__(self):
        # 列表推导式语法：[ 表达式 for 变量 in 可迭代对象 if 真值表达式 ]  
        self.dataList = [self.__init_item(city) for city in self.bigCities]
        # print(dataList)
        # logging.debug(self.dataList)

        # 按统计指标排序: 倒序
        self.sortedList = sorted(self.dataList, key=lambda item: item['totalRegist'], reverse=True)
        logging.debug(self.sortedList)

    def __init_item(self, cityName):
        # item = {'cityName': cityName}
        # for key in self.statIndicators:
        #     if key == 'downloadRate':
        #         item[key] = format(random.random(), '.2%')
        #     else:
        #         item[key] = random.randint(0, 100)
        
        # 字典推导式语法：{ 键表达式：值表达式 for 变量 in 可迭代对象 [if 真值表达式] }
        exclude_fields = ('downloadRate', 'totalInvite') # 不需要在for循环中初始化值的字段元组
        item = { key: random.randint(0, 100) for key in self.statIndicators if key not in exclude_fields}
        item['downloadRate'] = format(random.random(), '.2%')
        item['totalInvite'] = item['adminRegist'] + item['engineerInvite'] + item['userInvite']
        item['cityName'] = cityName
        return item

    def generate_html(self, templateFie, outputFile):
        logging.debug("templateFile=%s, outputFile=%s" % (templateFile, outputFile))
        # import locale, datetime
        # locale.setlocale(locale.LC_CTYPE, 'Chinese')
        # print(datetime.datetime.now().strftime('%Y年%m月%d日'))
        # statDate = time.strftime(u'%Y年%m月%d日', time.localtime())
        # UnicodeEncodeError: 'locale' codec can't encode character '\u5e74' in position 2: Illegal byte sequence
        statDate = time.strftime('%Y{y}%m{m}%d{d}', time.localtime()).format(y='年', m='月', d='日')
        logging.debug("Stat Date: %s" % statDate)
        logging.debug("os.path.abspath(__file__): %s" % os.path.abspath(__file__))
        logging.debug("os.path.dirname(parent): %s" % os.path.dirname(os.path.abspath(__file__)))
        
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        logging.debug("os.getcwd() = %s" % os.getcwd())

        env = Environment(loader=FileSystemLoader('./'))  # 加载模板
        template = env.get_template(templateFie)
        # template.stream(body).dump('result.html', 'utf-8')
        # with open(outputFile, 'w+') as fout: # 在windows下面，新文件的默认编码是gbk
        # UnicodeEncodeError: 'gbk' codec can't encode character '\ufeff' in position 0: illegal multibyte sequence
        with open(outputFile, 'w+', encoding='utf-8') as fout:
            html_content = template.render(dataList=self.sortedList, statDate=statDate)
            fout.write(html_content)  # 写入模板 生成html
    
    def generate_image(self, inputFile, outputFile):
        logging.debug("inputFile=%s, outputFile=%s" % (inputFile, outputFile))
        # 此处路径需要绝对路径，不能为软连接
        # config = imgkit.config(wkhtmltoimage='/usr/local/wkhtmltox/bin/wkhtmltoimage')
        config = imgkit.config(wkhtmltoimage = self.wkhtmltoimage)
        # imgkit.from_string(), imgkit.from_url()
        options = {
            'width': 1000,
            'height': 600,
            'encoding': 'UTF-8',
        }
        # imgkit.from_file(inputFile, outputFile, config=config, options=options)
        imgkit.from_file(inputFile, outputFile, config=config)
    
    def generate_pdf(self, inputFile, outputFile):
        logging.debug("inputFile=%s, outputFile=%s" % (inputFile, outputFile))
        # 此处路径需要绝对路径，不能为软连接
        # config = imgkit.config(wkhtmltoimage='/usr/local/wkhtmltox/bin/wkhtmltopdf')
        config = imgkit.config(wkhtmltoimage = self.wkhtmltopdf)
        # imgkit.from_string(), imgkit.from_url()
        imgkit.from_file(inputFile, outputFile, config=config)

    def qiniu_upload(self, bucket, outputFile):
        access_key = 'qiniu-access-key'
        secret_key = 'qiniu-secret-key'
        auth = qiniu.Auth(access_key, secret_key)

        # namespace = uuid.NAMESPACE_URL uuid.NAMESPACE_DNS
        # uuid.uuid1([node[, clock_seq]])  : 基于时间戳
        # uuid.uuid3(namespace, name) : 基于名字的MD5散列值
        # uuid.uuid4() : 基于随机数
        # uuid.uuid5(namespace, name) : 基于名字的SHA-1散列值
        # uuid.uuid5(uuid.NAMESPACE_DNS, 'python.org')
        key = str(uuid.uuid4()) 
        # 生成上传 token，可以指定过期时间等
        token = auth.upload_token(bucket, key, 600)
        result, info = qiniu.put_file(token, key, outputFile)
        print(info)
        assert result['key'] == key
        assert result['hash'] == qiniu.etag(outputFile)

        # 参考七牛文档，有两种方式构造base_url的形式(服务端实现)
        base_url = 'http://api.foo.com/qiniuToken?key=%s&bucket=%s' % (key, bucket)
        # 可以设置token过期时间
        private_url = auth.private_download_url(base_url, expires=3600)
        print(private_url)

        response = requests.get(private_url)
        result = json.loads(response.text)
        return result.get('url')

    def send_img(self, title, img_url, access_token):
        # @see https://developers.dingtalk.com/document/robots/custom-robot-access
        notify_url = 'https://oapi.dingtalk.com/robot/send?access_token={}'.format(access_token)
        header = {
            "Content-Type": "application/json",
            "Charset": "utf-8"
        }
        data = {
            "msgtype": "markdown",
            "markdown": {
                "title": title,
                "text": "#### 各个城市统计排行 \n" +
                        "> \n" +
                        "> ![screenshot](" + img_url + ")\n" +
                        "> ###### 10点10分发布 \n"
            },
            "at": {
                "atMobiles": [
                    "13800138000"
                ],
                "isAtAll": True
            }
        }

        print(json.dumps(data))
        # response = requests.post(notify_url, headers=header, json=data)
        # print(response.text)


if __name__ == "__main__":
    logging.basicConfig(level=logging.NOTSET, 
        format='%(asctime)s - %(filename)s[line:%(lineno)d/%(thread)d] - %(levelname)s: %(message)s') # 设置日志级别
    bigCities = ','.join(Stats2img.bigCities)
    logging.debug("Big Cities: %s" % bigCities)
    logging.debug("Stat Indicators: %s" % ','.join(Stats2img.statIndicators))

    stats2img = Stats2img()
    templateFile = 'stats-template.html'
    outHtmlFile = "stats-%s.html" % time.strftime("%Y%m%d", time.localtime())
    stats2img.generate_html(templateFile, outHtmlFile)
    
    outImgFile = "stats-%s.jpg" % time.strftime("%Y%m%d", time.localtime())
    stats2img.generate_image(outHtmlFile, outImgFile)

    outPdfFile = "stats-%s.pdf" % time.strftime("%Y%m%d", time.localtime())
    stats2img.generate_pdf(outHtmlFile, outPdfFile)

    logging.debug("上传图片")
    # img_url = stats2img.qiniu_upload('bucket_foo', outImgFile)
    img_url = 'http://img.foo.com/images/stats-sample.jpg'
    
    if len(sys.argv) > 1 and sys.argv[1] == 'prod':
        title = '每日数据统计'
        access_token = 'prod-access-token'
    else:
        title = '每日数据统计-测试'
        access_token = 'test-access-token'
    
    logging.debug("发送图片")
    stats2img.send_img(title, img_url, access_token)
