import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

airport=pd.read_excel('airportQ4.xlsx',sheet_name='Sheet1')
airport.fillna('0',inplace=True)

#airport['SBU'].fillna('暂无')
airport1=pd.read_excel('airportQ4.xlsx',sheet_name='Sheet1')
name=pd.read_excel('airportQ4.xlsx',sheet_name='name')
print('原数据量')
print(airport1.shape)
#airport.info()
print(airport.head(1))
print('---删除退票及国际机票数量---')
#airport.reset_index()
#print(airport.describe())
# print(airport.shape)

#删除不用统计的SBU
# SBUDEL=airport[airport.SBU.str.contains('融信云')]
# print(SBUDEL.shape[0])
# airport.drop(SBUDEL.index,inplace=True)

#删除退票的订单
Refund=airport[airport.订单状态.str.contains('退票')]
print(Refund.shape[0])
airport.drop(Refund.index,inplace=True)

#商务舱情况
Business=airport[airport.物理舱位.str.contains('公务舱')]
#Business.groupby('乘机人').agg({'订单状态': 'count', '折扣':'mean','实收/实付':'sum'})
# Business['乘机人']agg({'订单状态': 'count', '折扣':'mean','实收/实付':'sum'})

Business1=pd.pivot_table(Business,values='订单号',columns='航班类型',index='乘机人',\
aggfunc='count',margins=True)
#print(Business1)

#删除国际机票
International=airport[airport.航班类型.str.contains('国际')]
print(International.shape[0])
airport.drop(International.index,inplace=True)



print('---需统计机票数量---')
print(airport.shape)

# airport['出票日期'].astype()
print(airport['出票日期'].dtype)
#print(airport['出票日期'])
#dates = airport['出票日期'].to_datetime()
#print(dates, dates.dtype, type(dates))
#print(airport['出票日期'][0:4])#=pd.to_datetime(airport['出票日期'],format='%Y-%m-%d %H:%M')

#将出票日期列转化为日期格式
airport['出票日期1']=pd.to_datetime(airport['出票日期'],format='%Y/%m/%d %H:%M:%S')
print(airport['出票日期1'].head(2))
print(airport['出票日期'].head(2))

#通过切片获取月份
# print(airport['出票日期'].str[5:7].head(3))
# airport['出票月份']=airport['出票日期'].str[5:7]
# airport['出票月份']=airport['出票日期1'].month

airport['出票月份'] = pd.DatetimeIndex(airport['出票日期1']).month
print(airport['出票月份'].head(3))
#print('--出票月份--')
#print(airport['出票月份'].dtype)

#数据分析1：按月份统计机票金额，柱形图绘制
airport_totalPrice = airport.groupby('出票月份')['实收/实付'].sum()
print(airport_totalPrice)
#简单柱状图
# airport_totalPrice.plot(kind='bar', rot=10, title='携程各月机票价格', color=['b'],alpha=0.55)
# plt.show()
#house_bizcircle.plot(marker='o')
plt.ylabel('实收/实付价格')
plt.xlabel("月份")
#plt.xlim((-1, 10))
plt.ylim((0, 5000000))
airport_totalPrice1=airport_totalPrice.reset_index()
print(airport_totalPrice1)
plt.title('每月机票总价', fontsize=20)
x=np.arange(3)
y=np.array(list(airport_totalPrice1['实收/实付'].round(decimals=2)))
for a, b in zip(x, y):
    plt.text(a, b+123500, format(b,','), ha='center', va='top', fontsize=9)
print(x)
print(y)
plt.xticks(x,list(airport_totalPrice1['出票月份']))
plt.yticks(np.arange(1000000,6000000,1000000),["100w","200W",'300W','400W','500W'])
print(list(airport_totalPrice1['出票月份']))
plt.bar(x,  # 横轴数据
       y,  # 纵轴数据
        0.37,
       color='blue', label='总价', alpha=0.55,)
#plt.plot(x,y,marker='o')
plt.legend()
plt.savefig('/home/tarena/1909/month05/ctrip/‘每月机票价格.png',dpi=500)
#plt.show()

#数据分析2:折扣TOP
airport['折扣']=airport['折扣'].astype("float64")
print(airport['折扣'].dtype)

discount_top=airport.groupby('乘机人昵称').agg({'订单状态': 'count', '折扣':'mean','实收/实付':'sum'})

discount_top1=discount_top.sort_values(by='折扣',ascending=False)
discount_top1['折扣']=discount_top1['折扣'].round(2)
discount_top1['实收/实付']=discount_top1['实收/实付'].apply(lambda x:format(x,','))
discount_top1=pd.merge(discount_top1,name,on='乘机人昵称',how='left')
discount_top1=discount_top1.rename(columns={'订单状态':'订单量'})
#print(discount_top1[discount_top1['订单状态']>15].head(20))
no1=discount_top1[discount_top1['订单量']>15].head(10)
print(no1)

#数据分析3:根据SBU数据透视
# SBU=pd.pivot_table(airport,values='',columns='',index='',\
# aggfunc='',margins=True)
airport['提早4天']=airport['提早4天'].astype("float64")
airport['是否最低价']=airport['是否最低价'].astype("float64")
airport['是否全价']=airport['是否全价'].astype("float64")
SBU_analysis=airport.groupby('SBU').agg({'订单号': 'count',
                                         '实收/实付':'sum',
                                         '提前天数':'mean',
                                         '提早4天':'sum',
                                         '是否最低价':'sum',
                                         '是否全价':'sum',
                                         '折扣':'mean'})
SBU_analysis=SBU_analysis.sort_values(by='订单号',ascending=False)
SBU_analysis['占比']=SBU_analysis['订单号']/SBU_analysis['订单号'].sum()
SBU_analysis.loc['合计'] = SBU_analysis[['订单号','实收/实付','占比','提早4天',
                                       '是否最低价','是否全价']].apply(lambda x: x.sum())
SBU_analysis['提前4天占比']=SBU_analysis['提早4天']/SBU_analysis['订单号']
SBU_analysis['低价占比']=SBU_analysis['是否最低价']/SBU_analysis['订单号']
SBU_analysis['全价占比']=SBU_analysis['是否全价']/SBU_analysis['订单号']

SBU_analysis['占比'] = SBU_analysis['占比'].apply(lambda x: format(x, '.0%'))
SBU_analysis['提前4天占比'] = SBU_analysis['提前4天占比'].apply(lambda x: format(x, '.0%'))
SBU_analysis['低价占比'] = SBU_analysis['低价占比'].apply(lambda x: format(x, '.2%'))
SBU_analysis['全价占比'] = SBU_analysis['全价占比'].apply(lambda x: format(x, '.1%'))
SBU_analysis['实收/实付']=SBU_analysis['实收/实付'].apply(lambda  x:format(x,','))
SBU_analysis['提前天数']=SBU_analysis['提前天数'].round(2)
SBU_analysis['折扣']=SBU_analysis['折扣'].round(2)

#df['Prices'] = df['Prices'] / df['Prices'].sum()
SBU_analysis=SBU_analysis.rename(columns={'订单号': '机票量','折扣':'平均折扣'})
order = ['机票量', '实收/实付', '占比', '提前天数','提早4天', '提前4天占比','是否最低价',
         '低价占比', '是否全价','全价占比', '平均折扣',]
SBU_analysis= SBU_analysis[order]
SBU_analysis.loc['备注','机票量']= '去除退票及国际机票'
print(SBU_analysis)
#SBU_analysis.to_excel(excel_writer=r"SBU.xlsx")


#数据分析4:退票改签情况
print('-退票改签情况-')
#airport1.info()
tcr_analysis=airport1.groupby('SBU').agg({'订单号': 'count',
                                         '是否退票':'sum',
                                         '退票费':'sum',
                                         '是否改':'sum',
                                         '改签费':'sum'})
tcr_analysis=tcr_analysis.sort_values(by='订单号',ascending=False)
tcr_analysis.loc['合计'] = tcr_analysis.apply(lambda x: x.sum())
tcr_analysis['退票比例']=tcr_analysis['是否退票']/tcr_analysis['订单号']
tcr_analysis['退票比例'] = tcr_analysis['退票比例'].apply(lambda x: format(x, '.2%'))
tcr_analysis['改签比例']=tcr_analysis['是否改']/tcr_analysis['订单号']
tcr_analysis['改签比例'] = tcr_analysis['改签比例'].apply(lambda x: format(x, '.2%'))
tcr_analysis=tcr_analysis.rename(columns={'订单号': '机票量','是否退票':'退票笔数',
                                          '是否改':'改签笔数'})
order1 = ['机票量', '退票笔数', '退票比例','退票费', '改签笔数','改签比例','改签费']
tcr_analysis= tcr_analysis[order1]
print(tcr_analysis)


#数据分析5:退票改签TOP
tcr_Top=airport1.groupby('乘机人昵称').agg({'订单号': 'count','是否退票':'sum'})
tcr_Top=tcr_Top.sort_values(by='是否退票',ascending=False).head(10)
tcr_Top=tcr_Top.rename(columns={'订单号':'机票数','是否退票':'退票数'})
tcr_Top2=airport1.groupby('乘机人昵称').agg({'订单号': 'count','是否改':'sum'})
tcr_Top2=tcr_Top2.sort_values(by='是否改',ascending=False).head(10)
tcr_Top2=tcr_Top2.rename(columns={'订单号':'机票数','是否改':'改签数'})
tcr_Top=pd.merge(tcr_Top,name,on='乘机人昵称',how='left')
tcr_Top2=pd.merge(tcr_Top2,name,on='乘机人昵称',how='left')
print(tcr_Top)
print(tcr_Top2)

#数据分析6：按地区
area=airport.groupby('平台').agg({'订单号': 'count',
                                         '提前天数':'mean',
                                         '是否最低价':'sum',
                                         '是否全价':'sum',
                                         '折扣':'mean'})

area['占比']=area['订单号']/area['订单号'].sum()
area=area.sort_values(by='订单号',ascending=False)
area.loc['合计'] = area[['订单号','占比','是否最低价']].apply(lambda x: x.sum())
area['低价占比']=area['是否最低价']/area['订单号']
area['占比'] = area['占比'].apply(lambda x: format(x, '.0%'))
area['低价占比'] = area['低价占比'].apply(lambda x: format(x, '.1%'))
area['提前天数']=area['提前天数'].round(2)
area['折扣']=area['折扣'].round(2)
area=area.rename(columns={'订单号': '机票量','折扣':'平均折扣'})
order3 = ['机票量', '占比', '提前天数','低价占比', '平均折扣']
area= area[order3]
print(area)

#数据分析7：订票方式
booking=airport.groupby('预订方式').agg({'订单号':'count'}).sort_values(by='订单号',ascending=False)
booking['占比']=(booking['订单号']/booking['订单号'].sum()*100)


booking=booking.rename(columns={'订单号':'订单量'})
#booking.info()
print(booking)
booking1=booking.reset_index()
print(booking1)
plt.figure('订票方式', facecolor='lightgray')
plt.title('订票方式占比', fontsize=20)
# 整理数据
values=list(booking1['占比'])
print(values)
spaces = [0.01, 0.01, 0.05]
labels=list(booking1['预订方式'])
print(labels)
#labels = ['Java', 'C', 'Python', 'C++', 'VB', 'Other']
colors = ['dodgerblue', 'blue', 'orangered']
# 等轴比例
plt.axis('equal')
plt.pie(
    values,  # 值列表
    spaces,  # 扇形之间的间距列表
    labels,  # 标签列表
    colors,  # 颜色列表
    '%.1f%%',  # 标签所占比例格式
    #shadow=True,  # 是否显示阴影
    startangle=90,  # 逆时针绘制饼状图时的起始角度
    radius=0.6  # 半径
)
plt.legend()

booking.loc['合计'] = booking.apply(lambda x: x.sum())
booking['占比'] = booking['占比'].apply(lambda x: format(x/100, '.0%'))
print(booking)
#plt.show()
plt.savefig('/home/tarena/1909/month05/ctrip/‘机票预订方式.png',dpi=600)

#数据分析7：平均折扣分布
area_level = [0, 0.099, 0.199, 0.299, 0.399, 0.499, 0.599, 0.699, 0.799, 0.899, 0.999,1]
label_level = ['0-0.09', '0.1-0.19', '0.2-0.29', '0.3-0.39',
               '0.4-0.49', '0.5-0.59', '0.6-0.69','0.7-0.79', '0.8-0.89', '0.9-0.99',1]
discount_cut = pd.cut(airport['折扣'], bins=area_level, labels=label_level)

airport['折扣分布']=discount_cut
discount_cut1=airport.groupby('折扣分布').agg({'订单号':'count'})
discount_cut1=discount_cut1.rename(columns={'订单号':'订单量'})

discount_cut1.plot(kind='bar', rot=30, grid=False, title='机票折扣分布', fontsize='small', label='单量')
#plt.show()
plt.savefig('/home/tarena/1909/month05/ctrip/‘机票折扣分布.png',dpi=300)
#discount_cut1['订单量']=discount_cut1['订单量'].astype('float64')
discount_cut1=discount_cut1.reset_index()
discount_cut1['占比']=discount_cut1['订单量']/discount_cut1['订单量'].sum()
discount_cut1.loc['合计'] = discount_cut1[['订单量','占比']].apply(lambda x: x.sum())
discount_cut1['占比'] =discount_cut1['占比'].apply(lambda x: format(x, '.1%'))
print(discount_cut1)

#数据分析8：提前订票情况
advance_analysis=airport.groupby('提前天数1（整数）').agg({'订单号': 'count',
                                                   '折扣':'mean','全价':'sum','是否全价':'sum',})
#advance_analysis=advance_analysis.reset_index()
#SBU_analysis=SBU_analysis.sort_values(by='订单号',ascending=False)

#advance_analysis.loc['合计'] = advance_analysis[['订单号']].apply(lambda x: x.sum())
advance_analysis['占比']=advance_analysis['订单号']/advance_analysis['订单号'].sum()
advance_analysis['占比'] =advance_analysis['占比'].apply(lambda x: format(x, '.1%'))
advance_analysis['全价占比']=advance_analysis['是否全价']/advance_analysis['订单号']
advance_analysis['全价占比'] = advance_analysis['全价占比'].apply(lambda x: format(x, '.2%'))
advance_analysis['潜在节省']=(advance_analysis['折扣']-advance_analysis.loc[['4天以上'],['折扣']].values[0][0])*advance_analysis['全价'].values
#advance_analysis=advance_analysis.reset_index()
advance_analysis.loc['合计'] = advance_analysis[['订单号','潜在节省']].apply(lambda x: x.sum())
advance_analysis['潜在节省']=advance_analysis['潜在节省'].round(1)
advance_analysis['潜在节省']=advance_analysis['潜在节省'].apply(lambda  x:format(x,','))
advance_analysis['折扣']=advance_analysis['折扣'].round(2)
advance_analysis=advance_analysis.rename(columns={'订单号':'订单量','是否全价':'全价单量'})
print(advance_analysis)
#print(advance_analysis.loc[['4天以上'],['折扣']].values[0][0])
#advance_analysis.to_excel(excel_writer=r'提前订票.xlsx')
#print(advance_analysis['全价'].values)

#数据分析9：9折以上情况
#discount_09=airport.groupby('SBU').agg({'订单号': 'count','大于等于9折':'sum'})
airport['大于等于9折']=airport['大于等于9折'].astype('float64')
discount_09=airport.groupby('SBU').agg({'订单号': 'count',
                                         '大于等于9折':'sum'})
discount_09=discount_09.sort_values(by='大于等于9折',ascending=False)
discount_09.loc['合计'] = discount_09.apply(lambda x: x.sum())
discount_09['占比']=discount_09['大于等于9折']/discount_09['订单号']
discount_09['占比'] =discount_09['占比'].apply(lambda x: format(x, '.1%'))
print(discount_09)

discount_09TOP=airport.groupby('乘机人昵称').agg({'订单号': 'count',
                                         '大于等于9折':'sum'}).sort_values(by='大于等于9折',ascending=False).head(10)
discount_09TOP=pd.merge(discount_09TOP,name,on='乘机人昵称',how='left')
discount_09TOP=discount_09TOP.rename(columns={'订单号':'订单量','大于等于9折':'9折以上单量'})
print(discount_09TOP)

#导出到excel
writer=pd.ExcelWriter('/home/tarena/1909/month05/ctrip/机票统计.xlsx',engine="xlsxwriter")
SBU_analysis.to_excel(writer,sheet_name="SBU")
no1.to_excel(writer,sheet_name="折扣TOP10")
Business1.to_excel(writer,sheet_name="商务舱情况")
tcr_analysis.to_excel(writer,sheet_name="SBU退改签情况")
tcr_Top.to_excel(writer,sheet_name="退票TOP10")
tcr_Top2.to_excel(writer,sheet_name="改签TOP10")
area.to_excel(writer,sheet_name="按平台")
discount_cut1.to_excel(writer,sheet_name="订票折扣情况")
advance_analysis.to_excel(writer,sheet_name="提前订票情况")
discount_09.to_excel(writer,sheet_name="9折SBU情况")
discount_09TOP.to_excel(writer,sheet_name="9折TOP")
writer.save()


