# -*- coding:utf-8 -*-
import urllib
import urllib2
import re
import cookielib


#163邮箱类
class MAIL:

    #初始化
    def __init__(self):
        #获取登录请求的网址，也就是上边提到的请求网址
        self.loginUrl = "https://mail.163.com/entry/cgi/ntesdoor?style=-1&df=mail163_letter&net=&language=-1&from=web&race=&iframe=1&product=mail163&funcid=loginone&passtype=1&allssl=true&url2=https://mail.163.com/errorpage/error163.htm"
        #设置代理，以防止本地IP被封
        #self.proxyUrl = "http://202.106.16.36:3128"
        #初始化sid码
        self.sid = ""
        #第一次登陆所需要的请求头request headers，这个在消息头里的请求头有
        self.loginHeaders = {
            'Host':"mail.163.com",
            'User-Agent':"Mozilla/5.0 (Windows NT 6.1; WOW64; rv:45.0) Gecko/20100101 Firefox/45.0",
            'Accept':"text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            'Accept-Language':"zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3",
            'Accept-Encoding':"gzip, deflate, br",
            'Referer':"http://mail.163.com/",
            'Connection':"keep-alive",
        }
        #设置用户名和密码，输入自己的账号密码
        self.username = ''
        self.pwd = ''
        #post所包含的参数也就是参数里的表单数据
        self.post = {
            'savelogin':"0",
            'url2':"http://mail.163.com/errorpage/error163.htm",
            'username':self.username,
            'password':self.pwd
        }
        #对post编码转换
        self.postData = urllib.urlencode(self.post)
        #设置代理
        #self.proxy = urllib2.ProxyHandler({'http':self.proxyUrl})
        #设置cookie对象，会在登录后获取登录网页的cookie
        self.cookie = cookielib.LWPCookieJar()
        #设置cookie处理器
        self.cookieHandler = urllib2.HTTPCookieProcessor(self.cookie)
        #设置登录时用到的opener，相当于我们直接打开网页用的urlopen
        self.opener = urllib2.build_opener(self.cookieHandler,urllib2.HTTPHandler)


    #模拟登陆并获取sid码
    def loginPage(self):
        #发出一个请求
        request = urllib2.Request(self.loginUrl,self.postData,self.loginHeaders)
        #得到响应
        response = self.opener.open(request)
        #需要将响应中的内容用read读取出来获得网页代码，网页编码为utf-8
        content = response.read().decode('utf-8')
        #打印获得的网页代码
        sidpattern = re.compile('sid=(.*?)&', re.S)
        # 获取并储存sid码，打印出来
        result = re.search(sidpattern, content)
        self.sid = result.group(1)
        print self.sid

        # 通过sid码获得邮箱收件箱信息
    def messageList(self):
            # 重定向的网址，用获取到的sid码替换
            listUrl = 'http://mail.163.com/js6/s?sid=%s&func=mbox:listMessages&TopTabReaderShow=1&TopTabLofterShow=1&welcome_welcomemodule_mailrecom_click=1&LeftNavfolder1Click=1&mbox_folder_enter=1' % self.sid
            # 新的请求头
            Headers = {
                'Host': "mail.163.com",
                'User-Agent': "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:45.0) Gecko/20100101 Firefox/45.0",
                'Accept': "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                'Accept-Language': "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3",
                'Referer': "https://mail.163.com/js6/main.jsp?sid=%s&df=mail163_letter" % self.sid,
                'Connection': "keep-alive"
            }
            # 发出请求并获得响应
            request = urllib2.Request(listUrl, headers=Headers)
            print listUrl
            response = self.opener.open(request,'rb')
            # 提取响应的页面内容
            content = response.read().decode('utf-8')
            print content
            return content
            # 获取邮件信息

    def getmail(self):
        # 先获得收件箱列表页面内容
        messages = self.messageList()
        # 信息提取的正则表达式
        pattern = re.compile(
            '<string name="from">"(.*?)".*?name="to">(.*?)<.*?name="subject">(.*?)<.*?name="sentDate">(.*?)<.*?name="receivedDate">(.*?)</date>',
            re.S)
        # re模块中的findall会找出所有匹配的字符串，返回一个列表
        mails = re.findall(pattern, messages)
        # 遍历列表输出中相应项的内容，每个(.*?)对应了相应的项
        for mail in mails:
            print '-' * 50
            print '发件人', mail[0], '主题', mail[2], '发送时间', mail[3]
            print '收件人', mail[1], u'接收时间', mail[4]

#生成邮箱爬虫对象
mail = MAIL()
#调用loginPage方法来获取网页内容
mail.loginPage()
mail.messageList()
mail.getmail()
