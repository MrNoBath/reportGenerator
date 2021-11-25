#reportGenerator
#autor:Jade
#创建日期:2020/9/1

import xlwt,xlrd,os
import mySpider
from xlutils.copy import copy


def topicsHandler(location, index):   #读取并处理excel文件中话题
    workbook = xlrd.open_workbook(location)
    wb = copy(workbook)
    sheet = workbook.sheet_by_index(index)
    st = wb.get_sheet(index)
    rowCount = sheet.nrows
    sentence = ''
    totalReadCount = 0
    totalCommentCount = 0
    # 拼接链接
    for i in range(1, rowCount):
        topic = sheet.cell(i, 0).value
        topicDeTails = mySpider.getTopicDetails(topic)
        readCount = topicDeTails[0]
        commentCount = topicDeTails[1]
        st.write(i,1,readCount)
        st.write(i,2,commentCount)
        if(i==rowCount-1):
            sentence = sentence + topic + '阅读量' + readCount + '，讨论量' + commentCount + '。'
        else:
            sentence = sentence + topic + '阅读量' + readCount + '，讨论量' + commentCount + '；'
        #阅读量求和
        if (readCount.endswith('万')):
            totalReadCount = totalReadCount + float(readCount[:-1])
        elif (readCount.endswith('亿')):
            totalReadCount = totalReadCount + (10000 * float(readCount[:-1]))
        else:
            totalReadCount = totalReadCount + (float(readCount.strip())/10000)
            # 讨论量求和
        if (commentCount.endswith('万')):
            totalCommentCount = totalCommentCount + (10000 * float(commentCount[:-1]))
        else:
            totalCommentCount = totalCommentCount + float(commentCount)

    os.remove(location)
    wb.save(location)

    return(sentence,totalReadCount,totalCommentCount)