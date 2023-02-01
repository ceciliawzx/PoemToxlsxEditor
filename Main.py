import numpy as np
import pandas as pd
import re

"""
This is a program to transfer string to poem names, times and authors
and send the output to an excel file. 
Appending contents to the existed file should use the Append file instead.
"""

rawString = """
1、《苦昼短》（唐）李贺
29．《白梅》（元）王冕
2．《白头吟》（两汉）卓文君
30．《哥舒歌》（唐）西鄙人
31、《李延年歌》（两汉）李延年
3．《诀别书》（两汉）卓文君
32、《兵车行》（唐）杜甫
4、《临江王节士歌》（唐）李白
5、《鹧鸪天·、枫落河梁野水秋》（宋）苏庠
33、《塞下曲六首》（唐）李白
34.《七律·忆重庆谈判》（现代）毛泽东
6、《浣溪沙·堤上游人逐画船》（宋）欧阳修
7．《临江仙·浪滚长江东逝水》（明）杨慎
35．《闻鹧鸪》（清）尤侗
8．《临江仙·送钱穆父》（宋）苏轼
36．《喜见外弟又言别》（唐）李益
.《春江花月夜》（唐）张若虚
0、《李凭箜篌引》（唐）李贺
37．《州娄秀才寓居开元寺早秋月夜中见寄》（唐）柳宗元
1．《将进酒》（唐）李白
2．《行路难·其一》（唐）李白
38．《虞美人·银床淅沥青梧老》（清）纳兰性德
3．《苔》（清）袁枚
39．《鹧鸪天·送人》（宋）辛弃疾
4．《侍太子坐诗》（两汉）曹植
40．《鹧鸪天·代人赋》（宋）辛弃疾
5．《送李录事兄归襄邓》（唐）刘长卿
升．《鹧鸪天·彩袖殷勤捧玉钟》（宋）晏几道
．《逢雪宿芙蓉山主人》（唐）刘长卿
42．《鹧鸪天·西都作》（宋）朱敦儒
.《送灵澈上人》（唐）刘长卿
43．《十一月四日风雨大作二首》（宋）陆游
.《宿假湖》（唐）李白
.《沉醉东风·渔夫》（元）白朴
.《马嵬》（清）袁枚
《玉楼春·东山探梅》（宋）刘镇
《惜春词》（唐）温庭筠
《潇湘神·斑竹枝》（唐）刘禹锡
《虞美人·听雨》（宋）蒋捷
《一剪梅·舟过吴江》（宋）蒋捷
《醉赠刘二十八使君》（唐）白居易
《酬乐天扬州初逢席上见赠》（唐）刘禹锡
《同儿辈赋未开海棠》（金）元好问
54．《早梅》（唐）张谓
82．《老将行》（唐）王维
55．《苑中遇雪应制》（唐）宋之问
83．《唐多令·柳絮》（清）曹雪芹
56．《江上秋怀》（唐）李白
84．《唐多令·喵别》（宋）吴文英
57．《御街行·秋日怀旧》（宋）范仲淹
85．《唐多令·芦叶满汀洲》（宋）刘过
58．《乡思》（宋）李觏
86．《上李邕》（唐）李白
59．《渔家傲·秋思》（宋）范仲淹
87．《蝶恋花·伫倚危楼风细细》（宋）柳永
60．《苏幕遮·怀旧》（宋）范仲淹
88．《杨柳枝词九首》（唐）刘禹锡
61．《江上渔者》（宋）范仲淹
89、《临江仙·柳絮》（清）曹雪芹
2．《渔家傲·天接云涛连晓雾》（宋）李清照
90、《谢赐珍珠》（唐）江采萍
3．《夏日绝句》（宋）李清照
91．《望江南·梳云洗罢》（唐）温庭筠
64．《一剪梅·红藕香残玉簟秋》（宋）李清照
92．《视刀环歌》（唐）刘禹锡
65．《渔家傲·画鼓声中昏又晓》（宋）晏殊
93．《虞美人·寄公度》（宋）舒亶
66．《采桑子·时光只解催人老》（宋）晏殊
94．《赠范晔诗》（南北朝）陆凯
67．《蝶恋花·槛菊愁烟兰泣露》（宋）晏殊
95．《蝶恋花·辛苦最怜天上月》（清）纳当性
8．《鹊踏枝·萧索清秋珠泪坠》（五代）冯延已
96．《望江南·昏鸦尽》（清）纳兰性德
69．《鹊踏枝·雅道闲情抛掷久》（五代）冯延已
97．《卜算子·不是爱风坐》（宋）严蕊
70．《戏答元珍》（宋）欧阳修
98．《摊破浣溪沙·莲昌香销翠叶残》（五代）李煜
71．《戏赠丁判官》（宋）欧阳修
99．《少年行四首》（唐）王维
72．《蝶恋花·庭院深深深几许》（五代/宋代）冯延巳/欧阳修
100．《临江仙·高咏楚词酬午日》（宋）陈与义
73．《十二月十五夜》（清）袁枚
74、《春风》（清）袁数
75．《所见》（清）袁枚
76．《更漏子·玉炉香》（唐）温庭筠
77．《忆东山二首》（唐）李白
78．《金陵城西楼月下吟》（唐）李白
79．《鹤冲天·黄金榜上》（宋）柳永
80．《雨霖铃·寒蝉凄切》（宋）柳永
1．《清平乐·别来春半》（五代）李煜
"""

# Split the string into tokens required.
namesList = []
timeList = []
authorList = []
lines = rawString.splitlines()
for line in lines:
    if len(line) == 0:
        continue
    tokens = re.split('[《》（）]', line)
    if len(tokens) != 5:
        print("not in reg: line = " + line)
        continue
    namesList.append(tokens[1])
    timeList.append(tokens[3])
    authorList.append(tokens[4])


# DataFrame for pandas
df = pd.DataFrame({
    '诗名': namesList,
    '朝代': timeList,
    '作者': authorList
})
df.index += 1

# Write to file
writer = pd.ExcelWriter('file.xlsx')
df.to_excel(writer, sheet_name='诗词', index_label='序号', na_rep='NaN')

# To reformat the column width. This loop somehow doesn't work.
for column in df:
    # column_width = max(df[column].astype(str).map(len).max(), len(column))
    # print(column_width)
    col_idx = df.columns.get_loc(column)
    writer.sheets['诗词'].set_column(col_idx, col_idx, 20)

writer.close()
