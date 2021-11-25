#reportGenerator
#autor:Jade
#创建日期:2020/9/1


import urllib.request
from urllib import parse
from bs4 import BeautifulSoup



def getTopicDetails(topic):
    url = 'https://s.weibo.com/weibo?q={}&Refer=index'.format(parse.quote(topic))
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:79.0) Gecko/20100101 Firefox/79.0'}
    req = urllib.request.Request(url=url,headers=headers)
    response = urllib.request.urlopen(req)
    html = response.read()
    bs = BeautifulSoup(html, 'html.parser')
    spanList = bs.find_all('span')
    readCount = (spanList[0].string)[2:]
    commentCount = (spanList[1].string)[2:]

    return readCount,commentCount