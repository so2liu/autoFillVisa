import excel_read
from datetime import datetime
from fdfgen import forge_fdf


table=excel_read.gettable_excel("_访问学者_签证信息登记表_20170929154104"+".xls",1-1)
for i in table:
    table[i]=excel_read.ch2de(str(table[i]))

print(table['姓（拼音）']+" "+table['名（拼音）']+' '+table['提交时间'])
print(len(table))
date_de=str(excel_read.datecn2datade(table['出生日期']))
adress= table['家庭住址街名（拼音）']+table['家庭住址门牌号（拼音）']
city= table['出生城市（中文）']

if table['性别']=='nan':
    geschlecht='toggle_2'
else:geschlecht='toggle_4'


fields = [('fill_7', table['姓（拼音）']), ('fill_9',table['名（拼音）']),
          ('fill_10',str(date_de)), ('fill_11',city),
          ('fill_12','VR China'), ('fill_13','Chinesisch'),
          ('fill_16',table['护照号码']),
          ('fill_17',str(excel_read.datecn2datade(table['护照出具日期']))), ('fill_18',str(excel_read.datecn2datade(table['护照失效日期']))),
          ('fill_19', table['签发机关（英文）']), ('fill_23',adress),
          ('fill_24', table['邮编']+city), ('fill_25',table['Email']),
          ('fill_27', table['手机号码']),
          ('fill_37',table['父亲姓（拼音）']), ('fill_38', table['父亲名（拼音）']), ('fill_39',excel_read.ch2de(table['父亲出生日期（日月年）+出生地'])), ('fill_40', 'Chinesisch'), ('fill_41',city),
          ('fill_42',table['母亲姓（拼音）']), ('fill_43', table['母亲名（拼音）']), ('fill_44',excel_read.ch2de(table['母亲出生日期（日月年）+出生地'])), ('fill_45', 'Chinesisch'), ('fill_46',city),
          ('fill_20_3', table['语言班/大学/预科的街道（德文）']+table['语言班/大学/预科的门牌号（德文）']), ('fill_21_3',table['语言班/大学/预科的邮编']+table['语言班/大学/预科的城市（德文）']),
          ('fill_27_3', table['语言班/大学/预科的名字（德文）']),('fill_28_2', table['语言班/大学/预科的街道（德文）']+table['语言班/大学/预科的门牌号（德文）']),('fill_21_3',table['语言班/大学/预科的邮编']+table['语言班/大学/预科的城市（德文）']),
          ('fill_30_2', table['语言班/大学/预科的电话（德文）']), ('fill_31_2',table['语言班/大学/预科的Email（德文）']),
          ('fill_3_4', table['专业（英语或德语）']), ('fill_4_4',table['预期在德国入境时间（日月年）']),('fill_5_4',table['预期在德国入境时间（日月年）'][:-1]+str(int(table['预期在德国入境时间（日月年）'][-1])+2)),
          ('fill_6_4','Meine Eltern'), ('fill_15_4','Nein'),('fill_23_3',city+', VR China'),
          ('fill_16_4','Nein'),
          ('toggle_8','On'),('toggle_3','On'),('toggle_7','On'), #是否保留国内住址
           ('toggle_10_2','On'),
          ('undefined_5','On'),('undefined_7','On'),
          ('toggle_1_3','On'), ('toggle_4_2','On'),
          (geschlecht,'On')]

fdf = forge_fdf("",fields,[],[],[])
fdf_file = open("data.fdf","wb")
fdf_file.write(fdf)
fdf_file.close()

print(table['预期在德国入境时间（日月年）'][:-1]+str(int(table['预期在德国入境时间（日月年）'][-1])+2))




