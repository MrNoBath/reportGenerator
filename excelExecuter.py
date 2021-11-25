#reportGenerator
#autor:Jade
#创建日期:2020/9/1

import pandas as pd
import xlrd
import topicExe

def excelExe(dataPath,topicPath,index):
    media_types = ('报纸', '网络', '客户端', '微博', "微信", '论坛', '问答')
    web_media_types = ('报纸', '网络')
    web_media_number = 0  # 电子报+网络媒体数
    app_media_number = 0  # 手机端媒体数
    traditionnal_media_articles = 0  # 传统媒体发文量
    exposure_number = 0  # 曝光量
    weibo_post_number = 0  # 微博发布量
    weibo_read_number = 0 #微博阅读量
    weibo_total_interact_number = 0 #微博总交互量
    weibo_fans_max = 0  # 微博最大粉丝量
    weibo_fans = 0
    weibo_interact_number = 0  # 微博互动量
    weibo_migu_autors = 0  # 微博咪咕账号数
    weibo_migu_articles = 0  # 微博咪咕发文数
    weibo_mobile_authors = 0  # 微博移动账号数
    weibo_mobile_articles = 0  # 微博咪咕发文数
    wechat_authors = 0  # 微信作者数
    wechat_articles = 0  # 微信篇数
    wechat_reads = 0  # 微信阅读量
    wechat_migu_authors = 0  # 微博咪咕账号数
    wechat_migu_articles = 0  # 微博咪咕发文数
    wechat_mobile_authors = 0  # 微博移动账号数
    wechat_mobile_articles = 0  # 微博咪咕发文数
    bbs_media_number = 0  # 论坛媒体数
    bbs_article_number = 0  # 论坛文章数
    bbs_reads = 0
    QA_media_number = 0  # 问答媒体数
    QA_article_number = 0  # 问答文章数



    for media_type in media_types:
        try:
            df = pd.read_excel(dataPath, sheet_name=media_type)

        except xlrd.biffh.XLRDError:
            pass
        else:
            if media_type in web_media_types:
                web_media_number = web_media_number + len(df['媒体名称'].unique())
                traditionnal_media_articles = traditionnal_media_articles + df.shape[0]
                exposure_number = exposure_number + 5.12 * (len(df['媒体名称']))
            if media_type == '客户端':
                app_media_number = app_media_number + len(df['媒体名称'].unique())

                traditionnal_media_articles = traditionnal_media_articles + df.shape[0]
                exposure_number = exposure_number + 0.9 * (len(df['媒体名称']))
            if media_type == '微博':
                df_post = df[~df["标题"].str.contains(r"//@")]
                weibo_post_number = df_post.shape[0] + df_post["转发量"].sum()
                weibo_fans = df["粉丝量"].sum()
                weibo_fans_old = pd.pivot_table(df, index=['作者'], values=['粉丝量'], aggfunc='max')  # 透视最大粉丝量
                weibo_fans_max = weibo_fans_old['粉丝量'].sum()
                weibo_forward_number = df['转发量'].sum()
                weibo_comment_number = df['评论量'].sum()
                weibo_like_number = df['点赞量'].sum()
                weibo_interact_number = weibo_forward_number + weibo_comment_number + weibo_like_number
                df['作者'] = df['作者'].apply(str)  # 把“作者”列的格式全改为str
                df_migu = df[df['作者'].str.contains('咪咕')]

                weibo_migu_autors = len(df_migu["作者"].unique())
                weibo_migu_articles = df_migu.shape[0]

                df_mobile = df[df['作者'].str.contains('移动')]

                weibo_mobile_authors = len(df_mobile["作者"].unique())
                weibo_mobile_articles = df_mobile.shape[0]
            if media_type == "微信":
                wechat_authors = len(df['作者'].unique())
                wechat_articles = df.shape[0]
                wechat_reads = df['阅读量'].sum()
                df['作者'] = df['作者'].apply(str)  # 把“作者”列的格式全改为str
                df_migu = df[df['作者'].str.contains('咪咕')]
                wechat_migu_authors = len(df_migu["作者"].unique())
                wechat_migu_articles = df_migu.shape[0]
                df_mobile = df[df['作者'].str.contains('移动')] 
                wechat_mobile_authors = len(df_mobile["作者"].unique())
                wechat_mobile_articles = df_mobile.shape[0]
            if media_type == "论坛":
                bbs_media_number = len(df["媒体名称"].unique())
                bbs_article_number = df.shape[0]
                bbs_reads = df["阅读量"].sum()

            if media_type == "问答":
                QA_media_number = len(df["媒体名称"].unique())
                QA_article_number = df.shape[0]


    topicInfo = topicExe.topicsHandler(topicPath,index)
    weibo_read_number = weibo_fans*1.98/10000/10000+ topicInfo[1]
    weibo_total_interact_number = weibo_interact_number+topicInfo[2]
    weibo_read_numer_old = weibo_fans_max/10000+topicInfo[1]

    report_content = "一、新闻媒体：\n- 共有%d家网络媒体及%d个新闻客户端，发表报道%d篇，总曝光量%f万。\n" \
                     % (web_media_number, app_media_number, traditionnal_media_articles, exposure_number) + \
                     "- 微博，共有相关内容%d条，事件阅读量%f万(%f万)，交互量%f。咪咕系，共有%d个官方号发表%d条；移动系，共有%d个账号发表%d条。话题，" \
                     % (weibo_post_number, weibo_read_number,weibo_read_numer_old,weibo_total_interact_number,weibo_migu_autors, weibo_migu_articles, weibo_mobile_authors,
                        weibo_mobile_articles) +topicInfo[0]+\
                     "\n- 微信，共有%d个公号，发表相关文章%d篇，共产生阅读量%d。咪咕系，共有%d个公号发表%d篇；移动系，共有%d个公号发表%d篇。" \
                     % (wechat_authors, wechat_articles, wechat_reads, wechat_migu_authors, wechat_migu_articles,
                        wechat_mobile_authors, wechat_mobile_articles)+\
                     "\n- 论坛，共有%d个论坛，产生相关帖子%d篇，共产生阅读量%d。" % (bbs_media_number, bbs_article_number, bbs_reads) + \
                     "\n- 问答，共有%d个问答平台，产生相关问题%d条。" % (QA_media_number, QA_article_number)

    print(report_content)