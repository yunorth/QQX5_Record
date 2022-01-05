# 开发周期：2021/03至2021/12
# Python 3.7.1
# 这是我留给这个圈子最后的东西，告辞
# by 俞北 2022年1月5日18:09:37

import re
from tkinter import messagebox, filedialog
import os
import tkinter as tk
import sys
from pyecharts.charts import Line,Page
from pyecharts.charts import Bar,Pie
from pyecharts.charts import Grid
from pyecharts import options as opts
import csv
import re
import time
import numpy as np
import xlwings as xw
import json
import songJson
import binascii
from pyecharts.components import Table
from pyecharts.options import ComponentTitleOpts


#全局配置
fileName = ""    #录像文件路径

imageSubtitle = "北斗星通用测试版"    #生成html副标题

wb = xw.Book()      #excel全局生成base case


#按键统计类
class keyInfo:
    recordDownTime = 0  #记录时间（按下）
    keyDownType = 0     #按键类型（按下）
    realDownTime = 0    #实际时间（按下）
    timeDownDelay = 0   #延时
    stateDown = "/"     #状态

    recordUpTime = 0    #记录时间（弹起）
    keyUpType = 0       #按键类型（弹起）
    realUpTime = 0      #实际时间（弹起）
    timeUpDelay = 0     #延时
    stateUp = "/"       #状态

    netTime = '/'       # 网络时间

    timeDiffer = 0      #步长

    isMatch = "/"       # 是否匹配

    isSynchronize = "/" # 是否双排

    downMinusDown = "/" #同步差

    downMinusUp = "/"   #异步差

    #设置按下延时 = 按下记录时间 - 按下实际时间
    def setTimeDownDelay(self):
        #因为肯定有按下，一切以按下为标准，所以按下不需要做过多的判断
        self.timeDownDelay = int(self.recordDownTime) - int(self.realDownTime)

    #设置按下状态
    def setStateDown(self):
        if self.timeUpDelay == '/': #这行不要
            self.stateUp = '/'      #这行不要
        if int(self.timeDownDelay) >= 20:   #如果按下延时>=20
            self.stateDown = "靠后"   #设置状态为靠后
        else:
            self.stateDown = "正常"


    #设置弹起延时 = 弹起记录时间 - 弹起实际时间
    def setTimeUpDelay(self):
        if(self.recordUpTime!='/'): #如果按键弹起了，则进行弹起延时设置
            self.timeUpDelay = int(self.recordUpTime) - int(self.realUpTime)
        else:
            self.timeUpDelay = '/'  #如果按键没弹起，不进行设置

    #设置弹起状态
    def setStateUp(self):
        if self.timeUpDelay == '/': #如果按键没弹起，不进行设置
            self.stateUp = '/'
        if self.timeUpDelay != '/' and int(self.timeUpDelay) >= 20: #如果按键弹起了且延时>=20，则设置状态为靠后
            self.stateUp = "靠后"
        if self.timeUpDelay != '/' and int(self.timeUpDelay) < 20:  #如果按键弹起了且延时<20，则设置状态为正常
            self.stateUp = "正常"

    #设置步长 = 弹起实际时间 - 按下时间时间
    def setTimeDiffer(self):
        if self.recordUpTime == '/':   #如果按键没弹起，不进行设置
            self.timeDiffer = '/'
        else:
            self.timeDiffer = int(self.realUpTime)- int(self.realDownTime)  #如果按键弹起了，则进行步长设置

    #设置双排
    def setSynchronize(self):
        if self.timeDiffer == '/':  #如果按键没弹起，不进行设置
            self.isSynchronize = '/'
        if self.timeDiffer == 0 and self.recordDownTime != '0': #如果按键弹起了，则进行双排设置
            self.isSynchronize = "Yes"
        else:
            self.isSynchronize = "No"

    #调试用toString
    def toString(self):
        print(str(self.recordDownTime)+" "+str(self.keyDownType)+" "+str(self.realDownTime)+" "+str(self.timeDownDelay)+" "+str(self.stateDown)+" "+str(self.recordUpTime)+" "+str(self.keyUpType)+" "+str(self.realUpTime)+" "+str(self.timeUpDelay)+" "+str(self.stateUp)+" "+str(self.netTime)+" "+str(self.timeDiffer)+" "+str(self.isMatch)+" "+str(self.isSynchronize))


# 判定统计类
class judgeNode:
    No = 0              #判定序号
    recordTime = '/'    #记录时间
    keyType = '/'       #按键类型
    keyTime = '/'       #实际时间
    netEvent = '/'      #网络时间
    beat = '/'          #小节数
    gameMode = '/'      #模式标识：CT/JZ
    timeDelay = '/'     #延时
    judgeData = '/'     #判定值
    judgeType = '/'     #判定类型

    #调试用toString
    def toString(self):
        print("序号："+str(self.No)+" 记录时间："+str(self.recordTime) +" 按键类型："+str(self.keyType)
              +" 按键时间："+str(self.keyTime) +" 网络时间："+str(self.netEvent)+" 小节："+str(self.beat)
              +" 模式："+str(self.gameMode)+" 延时："+str(self.timeDelay)+" 判定值："+str(self.judgeData)
              +" 判定类型："+str(self.judgeType))

    #设置延时 = 网络时间 - 实际时间
    def setTimeDelay(self):
        if self.keyTime != '/': #如果不是miss
            self.timeDelay = int(self.netEvent) - int(self.keyTime)

    #设置判定值
    def setJudgeData(self):
        if self.keyTime != '/':         #如果有按下，意味着非miss
            listDetail = []             #初始化判定范围list
            songInfo = getSongInfo()    #获取歌曲信息
            BPM = songInfo[2]           #从返回值中获得BPM
            BPM = float(BPM)            #转化为float类型，否则后续无法计算
            jddata = 0.0                #习惯了先初始化一个变量
            if self.gameMode == "CT":   #如果模式是传统
                listDetail = np.array(calculateBPM(BPM, 0))  # 计算BPM
            if self.gameMode == "JZ":   #如果模式是节奏
                listDetail = np.array(calculateBPM(BPM, 1))  # 计算BPM
            Pat = listDetail[0]  # 获取一拍数值
            jddata = int(self.keyTime) - np.round(int(self.keyTime) / Pat, 0) * Pat  # 计算判定偏差：你可以用你自己的公式
            self.judgeData = np.round(jddata, 4)  # 四舍五入保留4位小数

    #设置判定类型
    def setJudgeType(self):
            if self.keyTime != '/': #如果不是miss
                songInfo = getSongInfo()    #获取歌曲信息
                BPM = songInfo[2]           #从返回值中获得BPM
                BPM = float(BPM)            #转化为float类型，否则后续无法计算
                listDetail = np.array(calculateBPM(BPM, 0))     #用传统进行范围计算（因为传统节奏范围都一样）
                Perfect = listDetail[1]  # 初始化
                Great = listDetail[2]  # 初始化
                Good = listDetail[3]  # 初始化
                Bad = listDetail[4]  # 初始化
                jd = np.abs(self.judgeData) #把判定值取绝对值
                res = '/'   #miss
                #根据范围进行判断是哪个判定类型
                if jd <= Perfect:
                    res = "Perfect"
                if jd > Perfect and jd <= Great:
                    res = "Great"
                if jd > Great and jd <= Good:
                    res = "Good"
                if jd > Good and jd <= Bad:
                    res = "Bad"
                if jd > Bad:
                    res = "MISS"
                self.judgeType = res    #设置判定类型


# --------------------------获取录像者信息----------------------------
#   描述： 用于获取录像者信息：
#   参数：
#   返回值：listGameInfo = [录像时间,录像者名字,对局模式,录像者站位,对局人数]
def getGameInfo():
    playerCount = 0  # 记录玩家人数
    listGameInfo = []   # 录像时间  录像者名字  对局模式  录像者站位  对局人数
    with open(fileName, 'r', encoding='gb18030') as myfile: #读取录像文件
        for line in myfile.readlines():
            line = line.strip('\n')  #用\n分开每一行
            if "Record time=\"" in line:    # 录像时间
                RecordTime = re.findall(r"Record time=\"(.+?)\"", line) #正则提取，后续不再赘述
                listGameInfo.append(RecordTime[0])  #添加到listGameInfo中，后续不再赘述
            if "author=\"" in line:         # 录像者
                author = re.findall(r"author=\"(.+?)\"", line)
                listGameInfo.append(author[0])
            if "<GameMode>" in line:        #模式代码
                GameMode = re.findall(r"<GameMode>(.+?)</GameMode>", line)
                GameMode = getGameModeDetail(GameMode[0])   #获取模式代码对应的模式字符串
                listGameInfo.append(GameMode)
            if "<PlayerIndex>" in line:     #玩家站位
                PlayerIndex = re.findall(r"<PlayerIndex>(.+?)</PlayerIndex>", line)
                listGameInfo.append(PlayerIndex[0])
            if "OrgHeartValue" in line:     #忘记这个是什么了，反正从很多个地方可以获取玩家数量，都一样
                playerCount = playerCount + 1  # 筛有几个玩家
        listGameInfo.append(playerCount)    #数完以后把玩家数量加入listGameInfo
        return listGameInfo  #[录像时间,录像者名字,对局模式,录像者站位,对局人数]


# --------------------------获取所有玩家信息----------------------------
#   描述： 用于获取所有玩家信息
#   参数：
#   返回值：listAllPlayerInfo = [[玩家姓名,舞团,性别],.......,[玩家姓名,舞团,性别]]
def getAllPlayerInfo():
    listAllPlayerInfo  = []    #全部玩家list
    with open(fileName,'r',encoding='gb18030') as myfile:
        for line in myfile:
            line = line.strip('\n')         #以回车键为分隔符
            if "<Item Name=" in line:
                listSinglePlayerInfo = []   #初始化单个玩家list:listSinglePlayerInfo
                Name = re.findall(r"<Item Name=\"(.+?)\"",line) #获取玩家姓名
                Band = re.findall(r"Band=\"(.+?)\"",line)   #获取玩家舞团
                if len(Band) == 0 : #如果没有舞团则写无
                    Band = ["无"]
                listSinglePlayerInfo.append(Name[0])    #将玩家姓名加入到单个玩家list中
                listSinglePlayerInfo.append(Band[0])    #将玩家舞团加入到单个玩家list中
            if "<Model>" in line:
                Model = re.findall(r"<Model>(.+?)</Model>",line)    #获取人物性别模型
                if Model == ['male_dress']:
                    Model = ['男']
                if Model == ['female_dress']:
                    Model = ['女']
                listSinglePlayerInfo.append(Model[0])   #将玩家性别加入到单个玩家list中
                listAllPlayerInfo.append(listSinglePlayerInfo)  #将单个玩家list加入到全部玩家list:listAllPlayerInfo中
        return listAllPlayerInfo


# --------------------------获取歌曲信息----------------------------
#   描述： 用于获取所有玩家信息
#   参数：
#   返回值：[level, songName, BPM]
def getSongInfo():
    level = 0 #初始化歌曲编号
    with open(fileName, 'r', encoding='gb18030') as myfile:
        level = 0  # 初始化歌曲level
        for line in myfile.readlines():
            line = line.strip('\n')  # 以回车键为分隔符
            if "<Level>" in line:
                level = re.findall(r"<Level>(.+?)</Level>", line) #找到歌曲编号
                level = level[0]
    songName = ""   #初始化歌曲名
    if level != '0':    #如果不是自制歌
        # 下面这些注释掉的是用csv读取的
        # with open('songInfo.csv', 'r', encoding='utf-8') as f:
        #     reader = csv.reader(f)
        #     for row in reader:
        #         if row[0] == level:
        #             songName = row[1]
        #             BPM = row[3]
        #             return [level, songName, BPM]
        #     messagebox.showinfo("提示", "未找到对应的歌曲信息，请在songInfo.csv中按格式加入歌曲信息，点击确定程序将自动结束。")
        #     sys.exit(0)

        # 下面是用json的方式实现的
        json_data = songJson.songJson; #import songJson ： songJson.py里面有一个叫songJson的json，需要定期手动更新
        i = 0 #遍历json
        while i < len(json_data):
            if json_data[i]['level'] == int(level): #如果找到了我要的歌
                # 返回[level, songName, BPM]
                return [str(json_data[i]['level']), str(json_data[i]['songName']), str(json_data[i]['BPM'])]
            i = i + 1
        messagebox.showinfo("提示", "未找到对应的歌曲信息，请联系作者QQ2533285193")   #找完都还没找到，则报提示未找到
        sys.exit(0) #退出
    if level == '0':#如果是自制歌
        with open(fileName, 'r', encoding='gb18030') as myfile:
            for line in myfile.readlines():
                line = line.strip('\n')
                if "<AuotGenLevelBMP>" in line: #读BPM
                    BPM = re.findall(r"<AuotGenLevelBMP>(.+?)</AuotGenLevelBMP>", line)
                    BPM = BPM[0]
                if "<AuotGenLevelOwnerNick>" in line:   #读上传人
                    updateName = re.findall(r"<AuotGenLevelOwnerNick>(.+?)</AuotGenLevelOwnerNick>", line)
                if "<Data>" in line:    #读自制歌名
                    songName = hex_to_str(re.findall(r"<Data>(.+?)</Data>", line)[0])   #自制歌名在录像中是十六进制，所以要转换成字符串
                    songName = songName+"（上传人：" + updateName[0] + "）"   #拼接歌曲名+上传人
                    return [level, songName, BPM]

# --------------------------十六进制转字符串----------------------------
def hex_to_str(s):	#s="68656c6c6f"
	s=binascii.unhexlify(s) #unhexlify()传入的参数也可以是b'xxxx'(xxxx要符合16进制特征)
	return s.decode('utf-8') #s的类型是bytes类型，用encode()方法转化为str类型


# --------------------------获取判定范围信息----------------------------
#   描述： 用于获取判定范围信息
#   参数：精确BPM, 模式：0代表传统，1代表节奏
#   返回值：[拍值, Perfect, Great, Good, Bad]
def calculateBPM(correctBpm, mode):
    listDetail = []
    Pat = 0
    Perfect = 0
    Great = 0
    Good = 0
    Bad = 0
    #可以换成你自己的公式
    if mode == 0:  #CT
        Pat = 60000 / correctBpm
        Perfect = 2.5 * Pat / 90
        Great = 12 * Pat / 90
        Good = 26 * Pat / 90
        Bad = Pat / 2
    if mode == 1:   #JZ
        Pat = (60000 / correctBpm) / 4
        Perfect = (2.5 * Pat / 90) * 4
        Great = (12 * Pat / 90) * 4
        Good = (26 * Pat / 90) * 4
        Bad = Pat * 4 / 2
    listDetail.append(Pat)
    listDetail.append(Perfect)
    listDetail.append(Great)
    listDetail.append(Good)
    listDetail.append(Bad)
    return np.array(listDetail)


# --------------------------设置判定列表----------------------------
#   描述： 用于获取判定范围信息
#   参数：区分过的NetEvent列表，按键统计列表
#   返回值：judgeList
def setJudgeList(listNetEvent,listKeyInfo):
    i = 0  #循环NetEvent[网络时间，小节，模式]
    k = 1  #记录判定数量
    judgeList = []
    #初始化judgeList
    while i < len(listNetEvent):
        jdNode = judgeNode()    #新建一个判定类对象
        jdNode.No = k           #设置判定序号
        jdNode.netEvent = listNetEvent[i][0]    #获取网络时间
        jdNode.beat = listNetEvent[i][1]        #获取小节数
        jdNode.gameMode = listNetEvent[i][2]    #获取模式标识
        i = i + 1
        k = k + 1   #判定序号++
        judgeList.append(jdNode)    #把单个节点加入到列表中


    i = 0  #循环judgeList
    j = 0  #循环按键统计
    #在judgeList里面循环
    while i < len(judgeList) and j < len(listKeyInfo):
        if listKeyInfo[j].netTime != '/':   #如果在按键统计中发现这个键是个判定键
            if int(listKeyInfo[j].netTime) == int(judgeList[i].netEvent):   #如果这个判定键的网络时间和我预先设置好的网络时间匹配
                judgeList[i].recordTime = listKeyInfo[j].recordDownTime #设置按下记录时间
                judgeList[i].keyType = listKeyInfo[j].keyDownType   #设置按下类型
                judgeList[i].keyTime = listKeyInfo[j].realDownTime  #设置按下实际时间
                judgeList[i].netEvent = listKeyInfo[j].netTime      #设置按下网络时间
                judgeList[i].beat =  listNetEvent[i][1] #设置小节数
                judgeList[i].gameMode = listNetEvent[i][2]  #设置模式标识

                judgeList[i].setTimeDelay() #设置延时
                judgeList[i].setJudgeData() #设置判定值
                judgeList[i].setJudgeType() #设置判定类型

                i = i + 1
                j = j + 1
            else:
                i = i + 1       #不相等的话+1继续找相等的
        else:
            j = j + 1   #如果在按键统计中不是判定键，则++继续找

    # for s in judgeList:
    #     s.toString()
    return judgeList

# --------------------------写判定统计EXCEL表----------------------------
#   描述： 用于写判定统计EXCEL表
#   参数：计算完的判定统计list
#   返回值：
def writeEXCEL3(judgeList):
    portion = os.path.splitext(fileName)  # portion[0] 文件名

    sht2 = wb.sheets.add() # 新增一个表格
    sht2.name = "判定统计"
    head = ["序号","记录时间","按键类型","按键时间","网络时间","小节","模式","延时","判定偏差","判定类型"]

    #给listData赋值
    listData = []
    i = 0
    while i < len(judgeList):
        listData.append([judgeList[i].No,judgeList[i].recordTime,judgeList[i].keyType,judgeList[i].keyTime
                        ,judgeList[i].netEvent,judgeList[i].beat,judgeList[i].gameMode,judgeList[i].timeDelay
                         ,judgeList[i].judgeData,judgeList[i].judgeType])
        i = i + 1
    sht2.range('A1').value = head   #写头
    sht2.range('A2').value = listData   #写数值
    # excel单元格宽度自适应
    sht2.range("A1").columns.autofit()
    sht2.range("B1").columns.autofit()
    sht2.range("C1").columns.autofit()
    sht2.range("D1").columns.autofit()
    sht2.range("E1").columns.autofit()
    sht2.range("F1").columns.autofit()
    sht2.range("G1").columns.autofit()
    sht2.range("H1").columns.autofit()
    sht2.range("I1").columns.autofit()
    sht2.range("J1").columns.autofit()

    sht2.range('A1:J1').api.Font.Bold = True  # 把标题加粗
    sht2.range('A1:J1').api.Font.Color = 0xFF7F50  # 把标题设置为红色

    listJudge = sht2.range("J1:J" + str(len(judgeList) + 1)).value  #选取判定类型列
    #下面这些就是给不同的判定区分不同的颜色
    i = 0
    while i < len(listJudge):
        if listJudge[i] == "Perfect":
            sht2.range('J'+str(i+1)).api.Font.Bold = True  # 加粗
            # 淡紫色0xFFB6C1
            # 黄色0x00BFFF
            sht2.range('J' + str(i + 1)).api.Font.Color = 0x00BFFF  # 设置为黄色
        if listJudge[i] == "Great":
            #sht2.range('J'+str(i+1)).api.Font.Bold = True  # 加粗
            sht2.range('J' + str(i + 1)).api.Font.Color = 0x0000f0  # 设置为红色
        if listJudge[i] == "Good":
            #sht2.range('J'+str(i+1)).api.Font.Bold = True  # 加粗
            sht2.range('J' + str(i + 1)).api.Font.Color = 0x32CD32  # 设置为绿色
        if listJudge[i] == "Bad":
            #sht2.range('J'+str(i+1)).api.Font.Bold = True  # 加粗
            sht2.range('J' + str(i + 1)).api.Font.Color = 0xDC143C  # 设置为蓝色
        i = i + 1

    range1 = sht2.range("A1:J" + str(len(listJudge)))   #选取整个表
    #给表加边框
    range1.api.Borders(11).LineStyle = 1
    range1.api.Borders(12).LineStyle = 1

    wb.save(portion[0] + '数据统计.xls')    #保存


# 匹配按键代码与按键类型
def getKey(keyCode):
    dictKeyCode = {'30009': '↖',
                   '30004': '←',
                   '30007': '↙',
                   '30002': '↑',
                   '30003': '↓',
                   '30008': '↘',
                   '30006': '→',
                   '30010': '↗',
                   '30005': '空格'
                   }
    for key, value in dictKeyCode.items():
        if keyCode == key:
            return value


# ['12662', '空格', '100', '12652']
# ['12722', '空格', '101', '12722']

# --------------------------按键修复----------------------------
#   描述： 进行单一按键的按键修复，主要是根据按下进行弹起的修复
#   参数：单种按键按下list,单种按键弹起list
#   返回值：[listKeyDown, newlistKeyUp]:[修复完的单种按键按下，修复完的单种按键弹起]
def fixKeyData(listKeyDown, listKeyUp):
    lenKeyDown = len(listKeyDown)   #获取按下列表长度
    lenKeyUp = len(listKeyUp)       #获取弹起列表长度
    i = 0   #循环按下
    j = 0   #循环弹起
    newlistKeyUp = []   #新的弹起列表，用于存放修复好的弹起
    # listKeyDown = np.array(listKeyDown)
    # listKeyUp = np.array(listKeyUp)
    # print(lenKeyDown)
    # print(lenKeyUp)
    while i < lenKeyDown:
        if j == lenKeyUp:   #防止前面的按键全部匹配，而最后一个按键没弹起导致的弹起下标越界
            break
        #print("i:"+str(i)+"  j:"+str(j))
        timeDiffer = int(listKeyUp[j][3]) - int(listKeyDown[i][3])  #计算步长
        #print(timeDiffer)
        if timeDiffer < 0:  #如果有多余的弹起
            j = j + 1   #直接跳过
            continue
        if timeDiffer > 1000:   #如果按键未弹起
            newlistKeyUp.append(['/', '/', '101', '/']) #则标记未弹起
            i = i + 1
            continue
        #如果正常
        newlistKeyUp.append([listKeyUp[j][0], listKeyUp[j][1], listKeyUp[j][2], listKeyUp[j][3]])
        i = i + 1
        j = j + 1

    while len(newlistKeyUp) != lenKeyDown:  #修复前面的匹配完了，最后一个没弹起
        newlistKeyUp.append(['/', '/', '101', '/'])

    return [listKeyDown, newlistKeyUp]

#写按键统计
def writeEXCEL1(listDataObject):
    portion = os.path.splitext(fileName)  # portion[0] 文件名
    sht1 = wb.sheets[0]  # 新增一个表格
    sht1.name = "按键统计"

    head = ["序号","记录时间(按下)","按键类型(按下)","按下时间","延时(按下)","状态(按下)","记录时间(弹起)","按键类型(弹起)","弹起时间","延时(弹起)","状态(弹起)","网络时间","步长","匹配","双排","同步差","异步差"]

    i = 0
    listData = []
    #赋值
    while i < len(listDataObject):
        listData.append([i+1,listDataObject[i].recordDownTime, listDataObject[i].keyDownType,listDataObject[i].realDownTime,listDataObject[i].timeDownDelay,listDataObject[i].stateDown
                         ,listDataObject[i].recordUpTime, listDataObject[i].keyUpType,listDataObject[i].realUpTime,listDataObject[i].timeUpDelay,listDataObject[i].stateUp
                         ,listDataObject[i].netTime,listDataObject[i].timeDiffer,listDataObject[i].isMatch,listDataObject[i].isSynchronize
                         ,listDataObject[i].downMinusDown,listDataObject[i].downMinusUp])
        i = i + 1
    sht1.range('A1').value = head
    sht1.range('A2').value = listData
    sht1.range("A1").columns.autofit()
    sht1.range("B1").columns.autofit()
    sht1.range("C1").columns.autofit()
    sht1.range("D1").columns.autofit()
    sht1.range("E1").columns.autofit()
    sht1.range("F1").columns.autofit()
    sht1.range("G1").columns.autofit()
    sht1.range("H1").columns.autofit()
    sht1.range("I1").columns.autofit()
    sht1.range("J1").columns.autofit()
    sht1.range("K1").columns.autofit()
    sht1.range("L1").columns.autofit()
    sht1.range("M1").columns.autofit()
    sht1.range("N1").columns.autofit()
    sht1.range("O1").columns.autofit()
    sht1.range("P1").columns.autofit()
    sht1.range("Q1").columns.autofit()

    sht1.range("A1:Q1").api.Font.Bold = True  # 加粗
    sht1.range("A1:Q1").api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)

    listStatusDown = sht1.range("F1:F"+str(len(listDataObject)+1)).value
    i = 0
    while i < len(listStatusDown):
        if listStatusDown[i] == "靠后":
            #sht1.range('F'+str(i+1)).api.Font.Bold = True  # 加粗
            # 淡紫色0xFFB6C1  黄色0x00BFFF  蓝色0xDC143C
            #sht1.range('F'+str(i+1)).api.Font.Color = 0xDC143C  # 设置为红色RGB(255,0,0)
            sht1.range('F'+str(i+1)).color = 0x00BFFF  # 设置为红色RGB(255,0,0)
        i = i + 1

    listStatusUp = sht1.range("K1:K" + str(len(listDataObject) + 1)).value
    i = 0
    while i < len(listStatusUp):
        if listStatusUp[i] == "靠后":
            #sht1.range('K' + str(i + 1)).api.Font.Bold = True  # 加粗
            # 淡紫色0xFFB6C1
            #sht1.range('K' + str(i + 1)).api.Font.Color = 0xDC143C  # 设置为红色RGB(255,0,0)
            sht1.range('K' + str(i + 1)).color = 0x00BFFF  # 设置为红色RGB(255,0,0)
        i = i + 1

    listisSynchronize = sht1.range("O1:O" + str(len(listDataObject) + 1)).value
    i = 0
    while i < len(listisSynchronize):
        if listisSynchronize[i] == "Yes":
            # sht1.range('K' + str(i + 1)).api.Font.Bold = True  # 加粗
            # 淡紫色0xFFB6C1
            #sht1.range('O' + str(i + 1)).api.Font.Color = 0xDC143C  # 设置为红色RGB(255,0,0)
            sht1.range('O' + str(i + 1)).color = 0x0000ff  # 设置为红色RGB(255,0,0)
        i = i + 1

    range1 = sht1.range("A1:Q"+ str(len(listDataObject) + 1))
    range1.api.Borders(11).LineStyle = 1
    range1.api.Borders(12).LineStyle = 1
    wb.save(portion[0] + '数据统计.xls')



def addKey(listKeyDownEach,listKeyUpEach):
    i = 0
    ki = keyInfo()
    ki.recordDownTime = listKeyDownEach[0]
    ki.keyDownType = listKeyDownEach[1]
    ki.realDownTime = listKeyDownEach[3]
    ki.setTimeDownDelay()
    ki.setStateDown()

    ki.recordUpTime = listKeyUpEach[0]
    ki.keyUpType = listKeyUpEach[1]
    ki.realUpTime = listKeyUpEach[3]
    ki.setTimeUpDelay()
    ki.setStateUp()

    ki.setTimeDiffer()

    ki.setSynchronize()
    return ki

#一轮冒泡：通过按键按下记录时间进行第一次冒泡排序
def bubbleSort(listKeyInfo):
    for i in range(1, len(listKeyInfo)):
        for j in range(0, len(listKeyInfo) - i):
            if int(listKeyInfo[j].recordDownTime) > int(listKeyInfo[j + 1].recordDownTime):
                listKeyInfo[j], listKeyInfo[j + 1] = listKeyInfo[j + 1], listKeyInfo[j]
    return listKeyInfo

#二轮冒泡：如果按键按下记录时间相同，再通过实际时间进行排序
def reSortByRealDownTime(listKeyInfo):
    for i in range(1, len(listKeyInfo)):
        for j in range(0, len(listKeyInfo) - i):
            #如果同个记录时间下 实际时间排序错误 则重新排序
            if int(listKeyInfo[j].realDownTime) > int(listKeyInfo[j + 1].realDownTime) and int(listKeyInfo[j].recordDownTime) == int(listKeyInfo[j + 1].recordDownTime):
                listKeyInfo[j], listKeyInfo[j + 1] = listKeyInfo[j + 1], listKeyInfo[j]
    return listKeyInfo

#计算同步差 异步差
def calKeyMinus(listKeyInfo):
    for i in range(1,len(listKeyInfo)):
        listKeyInfo[i].downMinusDown = int(listKeyInfo[i].realDownTime) - int(listKeyInfo[i-1].realDownTime)
        if listKeyInfo[i-1].realUpTime != '/':
            listKeyInfo[i].downMinusUp = int(listKeyInfo[i].realDownTime) - int(listKeyInfo[i-1].realUpTime)
    return listKeyInfo



# 设置按键统计网络时间
# 延时统一设置为50
def setNetEvent(listKeyInfo):
    i = 0
    j = 0
    gm = getGameMode()
    #传统类
    if gm == '6' or gm == '12' or gm == '17' or gm == '18' or gm == '7':
        netEvent = getNetEventTime1()
        while i < len(listKeyInfo) and j < len(netEvent):
            if netEvent[j][0] - int(listKeyInfo[i].realDownTime) >= 0:
                if netEvent[j][0] - int(listKeyInfo[i].realDownTime) <= 50 and listKeyInfo[i].keyDownType == "空格":
                    listKeyInfo[i].netTime = str(netEvent[j][0])
                    i = i + 1
                    j = j + 1
                else:
                    i = i + 1
            else:
                j = j + 1
    #节奏
    if gm == '5' or gm == '11' :
        netEvent = getNetEventTime1()
        while i < len(listKeyInfo) and j < len(netEvent):
            if netEvent[j][0] - int(listKeyInfo[i].realDownTime) >= 0:
                if netEvent[j][0] - int(listKeyInfo[i].realDownTime) <= 50:
                    listKeyInfo[i].netTime = str(netEvent[j][0])
                    i = i + 1
                    j = j + 1
                else:
                    i = i + 1
            else:
                j = j + 1
    #炫舞模式
    if gm == '1' or gm == '8' :
        netEvent = divdeModeX5(getNetEventTime2())
        while i < len(listKeyInfo) and j < len(netEvent):
            #print(str(i)+ " " + str(j))
            #print(str(netEvent[j][0] - int(listKeyInfo[i].realDownTime)) + " " + str(netEvent[j][2]))

            if netEvent[j][0] - int(listKeyInfo[i].realDownTime) >= 0:
                if netEvent[j][0] - int(listKeyInfo[i].realDownTime) <= 50 and netEvent[j][2] == 'CT' and listKeyInfo[i].keyDownType == "空格":
                    #print("传统有效")
                    listKeyInfo[i].netTime = str(netEvent[j][0])
                    i = i + 1
                    j = j + 1
                    continue
                if netEvent[j][0] - int(listKeyInfo[i].realDownTime) <= 50 and netEvent[j][2] == 'JZ':
                    #print("节奏有效")
                    listKeyInfo[i].netTime = str(netEvent[j][0])
                    i = i + 1
                    j = j + 1
                    continue
                else:
                    i = i + 1
            else:
                j = j + 1

    #团队模式
    if gm == '2' or gm == '9' :
        netEvent = divdeModeTeam(getNetEventTime2())
        while i < len(listKeyInfo) and j < len(netEvent):
            # print(str(i)+ " " + str(j))
            # print(str(netEvent[j][0] - int(listKeyInfo[i].realDownTime)) + " " + str(netEvent[j][2]))
            if netEvent[j][0] - int(listKeyInfo[i].realDownTime) >= 0:
                if netEvent[j][0] - int(listKeyInfo[i].realDownTime) <= 50 and netEvent[j][2] == 'CT' and listKeyInfo[i].keyDownType == "空格":
                    # print("传统有效")
                    listKeyInfo[i].netTime = str(netEvent[j][0])
                    i = i + 1
                    j = j + 1
                    continue
                if netEvent[j][0] - int(listKeyInfo[i].realDownTime) <= 50 and netEvent[j][2] == 'JZ':
                    # print("节奏有效")
                    listKeyInfo[i].netTime = str(netEvent[j][0])
                    i = i + 1
                    j = j + 1
                    continue
                else:
                    i = i + 1
            else:
                j = j + 1

    return listKeyInfo



def getPlayerIndex():
    with open(fileName,'r',encoding='gb18030') as myfile:
        for line in myfile.readlines():
            line = line.strip('\n')
            if "<PlayerIndex>" in line:
                PlayerIndex = re.findall(r"<PlayerIndex>(.+?)</PlayerIndex>",line)[0]
                return PlayerIndex


# --------------------------获取传统/节奏网络时间----------------------------
#   描述： 通过modetag来设置模式标识
#   参数：
#   返回值：listNetEventData：录像者的网络时间：[网络时间，小节数，模式标识]
def getNetEventTime1():
    #print("获取网络判定")
    # ---------------------------获取网络判定时间全部代码---------------------------
    listNetEvent = []  # 网络判定提取
    with open(fileName, 'r', encoding='gb18030') as myfile:
        for line in myfile.readlines():
            line = line.strip('\n')  # 以回车键为分隔符
            if "<NetEvent" in line:  # 如果有NetEvent的代码
                listNetEvent.append(line)  # 则将其加入listNetEvent中
    # ------------------------提取录像者有效NetEvent时间--------------------------------------
    listNetEventData = []  # 录像者NetEvent时间提取
    playerIndex = getPlayerIndex()   #获取玩家站位
    playerIndex = int(playerIndex)   #类型相同才能比较
    gm = getGameMode()
    if gm == '6' or gm == '12' or gm == '17' or gm == '18' or gm == '7':
        modetag = "CT"  #如果模式是传统类，将标识设置为CT
    if gm == '5' or gm == '11' :
        modetag = "JZ"  #如果模式是节奏，将标识设置为JZ
    for line in listNetEvent:
        listData =[]    #要提取的这一行的数据
        pid = re.findall(r"playeridx=\"(.+?)\"",line)   #提取这行的PID
        pid = list(map(int, pid))   #将string变为int
        dt  = re.findall(r"datatype=\"(.+?)\"",line)    #提取这行的datatype
        dt = list(map(int, dt))   #将string变为int
        if pid[0] == playerIndex:    #如果站位和玩家站位相同
            if dt[0] == 1:  #如果是有效数据1
                NetEventTime = re.findall(r"NetEvent t=\"(.+?)\"", line)    #获取网络时间
                data_0 = re.findall(r" data_0=\"(.+?)\"", line) #获取小节数
                NetEventTime = list(map(int, NetEventTime))     # 将string变为int
                data_0 = list(map(int, data_0))                 # 将string变为int
                listData.append(NetEventTime[0]) # 提取网络时间
                listData.append(data_0[0] + 1)   # 提取小节数
                listData.append(modetag)  #设置模式标识
                listNetEventData.append(listData)  #加入到listNetEventTime中
    return listNetEventData


# --------------------------获取炫舞/团队网络时间----------------------------
#   描述： 通过modetag来设置模式标识
#   参数：
#   返回值：listNetEventData：录像者的网络时间：[网络时间，小节数，模式标识（默认为0，后续再改）]
def getNetEventTime2():
    # print("获取网络判定")
    # ---------------------------获取网络判定时间全部代码---------------------------
    listNetEvent = []  # 网络判定提取
    with open(fileName, 'r', encoding='gb18030') as myfile:
        for line in myfile.readlines():
            line = line.strip('\n')  # 以回车键为分隔符
            if "<NetEvent" in line:  # 如果有NetEvent的代码
                listNetEvent.append(line)  # 则将其加入listNetEvent中
    # ------------------------提取录像者有效NetEvent时间--------------------------------------
    listNetEventData = []  # 录像者NetEvent时间提取
    playerIndex = getPlayerIndex()   #获取玩家站位
    playerIndex = int(playerIndex)  #类型相同才能比较
    for line in listNetEvent:
        listData =[] #要提取的这一行的数据
        pid = re.findall(r"playeridx=\"(.+?)\"", line)  # 提取这行的PID
        pid = list(map(int, pid))  # 将string变为int
        dt = re.findall(r"datatype=\"(.+?)\"", line)  # 提取这行的datatype
        dt = list(map(int, dt))  # 将string变为int
        if pid[0] == playerIndex:  # 如果站位和玩家站位相同
            if dt[0] == 1:  # 如果是有效数据1
                NetEventTime = re.findall(r"NetEvent t=\"(.+?)\"", line)    #获取网络时间
                data_0 = re.findall(r" data_0=\"(.+?)\"", line) #获取小节数
                NetEventTime = list(map(int, NetEventTime))  # 将string变为int
                data_0 = list(map(int, data_0))  # 将string变为int
                listData.append(NetEventTime[0]) #提取网络时间
                listData.append(data_0[0]+1)     #提取小节数
                listData.append(0)  #初始化模式为0传统  （1为节奏），后续再改
                listNetEventData.append(listData)  # 将这一行则加入到listNetEventTime中
    return listNetEventData

# --------------------------区分炫舞模式----------------------------
#   描述： 初始化好的网络时间列表[网络时间，小节数，模式标识（默认为0，后续再改）]
#   参数：
#   返回值：listNetEventData：区分好模式的网络时间列表：[网络时间，小节数，模式标识]
def divdeModeX5(listNetEventData):
    dataLen = len(listNetEventData) #获取网络时间列表长度
    count = 0   #段落计数
    i = 0   #循环网络时间列表
    while i < dataLen:
        j = i + 1   #j用来标识当前小节的下一小节数
        if j == dataLen:    #如果j已经到达列表末尾，此时i正在j-1的位置，则退出循环
            break
        else:
            #进入第一大段
            if count == 0:
                #进入传统第一段
                listNetEventData[i][2] = "CT"   #设置为传统
                if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:  # + 4 进入下一大段
                    count = 1   #第一大段结束
                    i = i + 1
                    continue

            #进入第二大段
            if count == 1:
                #进入节奏第一段
                listNetEventData[i][2] = "JZ"   #设置为节奏
                if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:  # + 4进入下一大段
                    count = 2   #第二大段结束
                    i = i + 1
                    continue

            #进入第三大段
            if count == 2:
                #进入传统第二段
                listNetEventData[i][2] = "CT"   #设置为传统
                if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:  # + 4进入下一大段
                    count = 3   #第三大段结束
                    i = i + 1
                    continue
                if j == dataLen-1:
                    listNetEventData[j][2] = "CT"   #在只有 传统/节奏/传统 的形式中，最后一个判定要设置为传统

            #进入第四大段
            if count == 3:
                #进入节奏第二段
                listNetEventData[i][2] = "JZ"    #设置为节奏
                if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:  # + 4进入下一大段
                    count = 4
                    i = i + 1
                    continue
                if j == dataLen-1:
                    listNetEventData[j][2] = "JZ"   #在只有 传统/节奏/传统/节奏 的形式中，最后一个判定要设置为节奏

            #进入第五大段
            if count == 4:
                #进入传统第三段
                listNetEventData[i][2] = "CT"   #设置为传统
                listNetEventData[j][2] = "CT"   #在只有 传统/节奏/传统/节奏/传统 的形式中，最后一个判定要设置为传统
        i = i + 1

    return listNetEventData

# 获取玩家数量
def getPlayerCount():
    playerCount = 0
    with open(fileName, 'r', encoding='gb18030') as myfile:
        for line in myfile.readlines():
            line = line.strip('\n')  #
            if "OrgHeartValue" in line:
                playerCount = playerCount + 1  # 筛有几个玩家\
    return playerCount

# --------------------------区分团队模式----------------------------
#   描述： 初始化好的网络时间列表[网络时间，小节数，模式标识（默认为0，后续再改）]
#   参数：
#   返回值：listNetEventData：区分好模式的网络时间列表：[网络时间，小节数，模式标识]
def divdeModeTeam(listNetEventData):
    dataLen = len(listNetEventData)     #获取网络时间列表长度

    playerIndex = getPlayerIndex()  # 获取玩家站位
    playerIndex = int(playerIndex)  # 类型相同才能比较
    playerCount = getPlayerCount()  #获取玩家数量
    calMode = 1 # 初始化计算模式 6人开局为1  4人开局为2  2人开局为3
    if playerCount == 6:
        calMode = 1
    if playerCount == 4:
        calMode = 2
    if playerCount == 2:
        calMode = 3

    pid = 0 #初始化玩家站位
    #6开
    if calMode == 1:

        pid = playerIndex + 1   #此处的玩家站位使用+1来标识，后续不再赘述
        #6开1号位
        if pid == 1 or pid == 4:  # 传统 节奏 传统 传统 传统
            JZcount = 0 #节奏段计数
            CTcount = 0 #传统段计数
            i = 0   #循环网络时间
            while i < dataLen:
                j = i + 1   #j等于i的下一个，用来做对比的
                #print("i="+str(i)+",j="+str(j))
                if j == dataLen:    #如果j已经到末尾了，此时i在j-1的位置，则退出
                    break
                else:
                    #第一大段
                    if CTcount == 0:
                        #进入传统第一段
                        listNetEventData[i][2] = "CT"
                        # 这个写的我自己有点看不懂。。。别纠结，反正it works
                        # 团队第一小段传统结束后，+2就进入PK环节
                        # 一般来讲团队一开始小节数都是+1，但是不排除也有+2的，所以and的第二个是用来确认是否进入了节奏，j = i + 1，i + 2 = j + 1， i + 3 = j + 2
                        if (listNetEventData[i][1] + 2) == listNetEventData[j][1]  and listNetEventData[i+2][1] == listNetEventData[i+3][1]:
                            CTcount = CTcount + 1
                            i = i + 1
                            continue
                    #第二大段
                    if CTcount == 1 and JZcount == 0:
                        #进入节奏第一段
                        listNetEventData[i][2] = "JZ"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:  # + 4进入下一大段
                            JZcount = JZcount + 1
                            i = i + 1
                            continue
                    #6开1号位在经历过节奏段落后，后段落全是CT
                    if JZcount == 1:
                        listNetEventData[i][2] = "CT"
                        listNetEventData[j][2] = "CT"
                i = i + 1
        #6开2号位
        if pid == 2 or pid == 5:
            JZcount = 0
            CTcount = 0
            i = 0
            while i < dataLen:
                j = i + 1
                if j == dataLen:
                    break
                else:
                    if CTcount == 0:
                        # 进入传统第一大段，注意是第一大段（开头小段+第一段传统）
                        listNetEventData[i][2] = "CT"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:# + 4进入下一大段
                            CTcount = CTcount + 1
                            i = i + 1
                            continue
                    if CTcount == 1 and JZcount == 0:
                        # 进入节奏第一段
                        listNetEventData[i][2] = "JZ"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:# + 4进入下一大段
                            JZcount = JZcount + 1
                            i = i + 1
                            continue
                    #6开2号位在经历过节奏段落后，后段落全是CT
                    if JZcount == 1:
                        listNetEventData[i][2] = "CT"
                        listNetEventData[j][2] = "CT"
                i = i + 1
        #6开3号位
        if pid == 3 or pid == 6:
            JZcount = 0
            CTcount = 0
            i = 0
            while i < dataLen:
                j = i + 1
                if j == dataLen:
                    break
                else:
                    if CTcount == 0:
                        # 进入传统第一大段，注意是第一大段（开头小段+第一段传统）
                        listNetEventData[i][2] = "CT"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:  # + 4进入下一大段
                            CTcount = CTcount + 1
                            i = i + 1
                            continue
                    if CTcount == 1:
                        # 进入传统第二段
                        listNetEventData[i][2] = "CT"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:  # + 4进入下一大段
                            CTcount = CTcount + 1
                            i = i + 1
                            continue

                    if CTcount == 2 and JZcount == 0:
                        # 进入节奏第一段
                        listNetEventData[i][2] = "JZ"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:  # + 4进入下一大段
                            JZcount = JZcount + 1
                            i = i + 1
                            continue
                    #6开3号位在经历过节奏段落后，后段落全是CT
                    if JZcount == 1:
                        listNetEventData[i][2] = "CT"
                        listNetEventData[j][2] = "CT"
                i = i + 1
    if calMode == 2:
        pid = playerIndex + 1
        gameInfo = getGameInfo()    #[录像时间,录像者名字,对局模式,录像者站位,对局人数]
        name = gameInfo[1]  #获取玩家姓名
        playerInfo = getAllPlayerInfo()   #[[玩家姓名,舞团,性别],.......,[玩家姓名,舞团,性别]]
        lenPlayerInfo = len(playerInfo) #获取玩家个数
        count = 0   #循环玩家列表
        while count < lenPlayerInfo:
            if playerInfo[count][0] == name:    #因为名字不能重复，所以名字可以做唯一标识
                newPID = count  #获取站位
                break
            else:
                count = count + 1
        if count== 0 or count == 2 :    #站在1号位
            JZcount = 0
            CTcount = 0
            i = 0
            while i < dataLen:
                j = i + 1
                if j == dataLen:
                    break
                else:
                    #传统 节奏 传统 节奏 传统
                    if CTcount == 0:
                        #传统第一大段
                        listNetEventData[i][2] = "CT"
                        if (listNetEventData[i][1] + 2) == listNetEventData[j][1] and listNetEventData[i+2][1] == listNetEventData[i+3][1]:
                            CTcount = CTcount + 1
                            i = i + 1
                            continue
                    if CTcount == 1 and JZcount == 0:
                        #进入节奏第一段
                        listNetEventData[i][2] = "JZ"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:
                            JZcount = JZcount + 1
                            i = i + 1
                    if CTcount == 1 and JZcount == 1:
                        #传统第二段
                        listNetEventData[i][2] = "CT"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:
                            CTcount = CTcount + 1
                            i = i + 1
                            continue
                    if CTcount == 2 and JZcount == 1:
                        #进入节奏第二段
                        listNetEventData[i][2] = "JZ"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:
                            JZcount = JZcount + 1
                            i = i + 1
                    if JZcount == 2:
                        #最后一段CT
                        listNetEventData[i][2] = "CT"
                        listNetEventData[j][2] = "CT"
                i = i + 1

        if count == 1 or count == 3 :   #站在2号位
            JZcount = 0
            CTcount = 0
            i = 0
            while i < dataLen:
                j = i + 1
                if j == dataLen:
                    break
                else:
                    # 传统 传统 节奏 传统 传统
                    if CTcount == 0:
                        #传统一大段
                        listNetEventData[i][2] = "CT"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:
                            CTcount = CTcount + 1
                            i = i + 1
                            continue
                    if CTcount == 1 and JZcount == 0:
                        #进入节奏第一段
                        listNetEventData[i][2] = "JZ"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:
                            JZcount = JZcount + 1
                            i = i + 1
                    if CTcount == 1 and JZcount == 1:
                        #最后一段CT
                        listNetEventData[i][2] = "CT"
                        listNetEventData[j][2] = "CT"
                i = i + 1
        # messagebox.showinfo("俞北说", "四人模式有待完善")
        # sys.exit(0)


    if calMode == 3 :
        pid = playerIndex + 1
        if pid == 1 or pid == 2 or pid == 3:
            JZcount = 0
            CTcount = 0
            i = 0
            while i < dataLen:
                j = i + 1
                if j == dataLen:
                    break
                else:
                    # 传统 节奏 节奏 节奏 传统
                    if CTcount == 0:
                        #进入传统第一段
                        listNetEventData[i][2] = "CT"
                        if (listNetEventData[i][1] + 2) == listNetEventData[j][1]  and listNetEventData[i+2][1] == listNetEventData[i+3][1]:
                            CTcount = CTcount + 1
                            i = i + 1
                            continue
                    if CTcount == 1 and JZcount == 0:
                        #进入节奏第一段
                        listNetEventData[i][2] = "JZ"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:
                            JZcount = JZcount + 1
                            i = i + 1
                            continue
                    if CTcount == 1 and JZcount == 1:
                        #进入节奏第二段
                        listNetEventData[i][2] = "JZ"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:
                            JZcount = JZcount + 1
                            i = i + 1
                            continue
                    if CTcount == 1 and JZcount == 2:
                        #进入节奏第三段
                        listNetEventData[i][2] = "JZ"
                        if (listNetEventData[i][1] + 4) == listNetEventData[j][1]:
                            JZcount = JZcount + 1
                            i = i + 1
                            continue
                    #后续全是CT
                    if JZcount == 3:
                        listNetEventData[i][2] = "CT"
                        listNetEventData[j][2] = "CT"
                i = i + 1
    return listNetEventData





# --------------------------读取按键----------------------------
#   描述： 读取按键并进行修复，最终写入excel
#   参数：
#   返回值：listKeyInfo：按键统计
def readKey():
    listKeyDown = []    #单种按下列表
    listKeyUp = []      #单种弹起列表
    listKey = ['30009', '30004', '30007', '30002', '30003', '30008', '30006', '30010', '30005'] #按键代码
    listKeyInfo = []    #按键对象列表
    for key in listKey: #通过按键类型代码来做循环匹配
        recordCount = -1        # <Record time="2021年 4月 2日" author="圈外物种"> 这行不算
        flag = False            # 记录块标识，flag一开始都是flase，表示没进记录块
        recordTimeValue = ""    # 记录时间初始化
        with open(fileName, 'r', encoding='gb18030') as myfile:
            for line in myfile.readlines():
                line = line.strip('\n')
                # print(line)
                if "\">" in line:                   #记录时间标签
                    recordCount = recordCount + 1   #记下记录的数量，初始值为-1，也可以从input开始读取，都一样
                    if recordCount > 0:
                        recordTimeValue = re.findall(r"<Item t=\"(.+?)\">", line)[0]    #读取记录时间
                        flag = True     #True表示接下来我要开始读取实际时间
                if flag:                #如果进入了一个记录块
                    if "Key=\"" + key + "\"" in line:
                        # key = keyInfo()
                        keyID = getKey(re.findall(r"Key=\"(.+?)\"", line)[0])   #通过按键代码获取按键类型
                        eventID = re.findall(r"EventID=\"(.+?)\"", line)[0]     #读取按键是100还是101
                        RelT = re.findall(r"RelT=\"(.+?)\"", line)[0]           #读取实际时间
                        if eventID == "100":
                            #k = k + 1
                            listKeyDown.append([recordTimeValue, keyID, eventID, RelT]) #加入到按下列表
                        if eventID == "101":
                            #w = w + 1
                            listKeyUp.append([recordTimeValue, keyID, eventID, RelT])   #加入到弹起列表
                if "</Item>" in line:
                    flag = False        #块读取结束，flag重新变为flase，直到遇到下一个记录块
                if "NetEvent" in line:  #后面的不看了提高效率
                    break

        listKeyDown = fixKeyData(listKeyDown, listKeyUp)[0] #修复按下
        listKeyUp = fixKeyData(listKeyDown, listKeyUp)[1]   #修复弹起

        i = 0   #循环按下（因为经过修复后按下和弹起列表长度一定相同）
        while i < len(listKeyDown):
            knode = keyInfo()           #新建按键节点
            knode = addKey(listKeyDown[i],listKeyUp[i]) #赋予节点按键数据
            listKeyInfo.append(knode)   #把单个按键加入到总列表中
            i = i + 1
        listKeyDown.clear() #清空列表进行下一种按键读取
        listKeyUp.clear()   #清空列表进行下一种按键读取

    #读取完所有类型的按键后
    listKeyInfo = bubbleSort(listKeyInfo)   #第一轮根据记录时间进行排序
    listKeyInfo = reSortByRealDownTime(listKeyInfo) #第二轮保持记录时间不变，通过实际时间再排一次序
    listKeyInfo = calKeyMinus(listKeyInfo)  #计算同异步差
    listKeyInfo = setNetEvent(listKeyInfo)  #设置网络时间
    writeEXCEL1(listKeyInfo)    #写入excel

    return listKeyInfo


# 按键数据统计类
class dataAnalyz:
    totalDown = 0   #总按下次数
    totalUp = 0     #总弹起次数
    realNotUp = 0   #实际未弹起次数
    ZTimes = 0      #Z数量
    XTimes = 0      #X数量
    plusTimes = 0   #+数量
    minusTimes = 0  #-数量
    preDownTimes = 0    #靠前按下次数
    postDownTimes = 0   #靠后按下次数
    preUpTimes = 0      #靠前弹起次数
    postUpTimes = 0     #靠后弹起次数
    synchronizeTimes = 0    #双排次数
    doubleBlanksTimes = 0   #双空格次数

    shangTimes = 0  # ↑ 次数
    xiaTimes = 0    # ↓ 次数
    zuoTimes = 0    # ← 次数
    youTimes = 0    # → 次数
    zuoshangTimes = 0   # ↖ 次数
    zuoxiaTimes = 0     # ↙ 次数
    youshangTimes = 0   # ↗ 次数
    youxiaTimes = 0     # ↘ 次数
    blank = 0           # 空格 次数


# 匹配对局模式代码与模式类型
def getGameModeDetail(GameMode):
    dictKeyCode = {'1': '炫舞模式(4K)',
                   '2': '团队模式(4K)',
                   '5': '节奏模式(4K)',
                   '6': '传统模式(4K)',
                   '8': '炫舞模式(8K)',
                   '9': '团队模式(8K)',
                   '11': '节奏模式(8K)',
                   '12': '传统模式(8K)',
                   '17': '疯狂模式(4K)',
                   '18': '疯狂模式(8K)'
                   }
    for key, value in dictKeyCode.items():
        if GameMode == key:
            return value

# 单独获取模式代码，因为如果每次都从list里面取的话效率是很低的，因为要多读很多东西
def getGameMode():
    with open(fileName,'r',encoding='gb18030') as myfile:
        for line in myfile.readlines():
            line = line.strip('\n')
            if "<GameMode>" in line:
                GameMode = re.findall(r"<GameMode>(.+?)</GameMode>", line)[0]
                return GameMode

#
# --------------------------计算按键数据统计----------------------------
#   描述： 设置“按键统计”EXCEL表格
#   参数：listKeyInfo：按键统计
#   返回值： da,listRes   #按键统计类对象 和 步长统计数据
def setSheet0(listKeyInfo):
    da = dataAnalyz()   #新建按键数据统计类对象
    with open(fileName, 'r', encoding='gb18030') as myfile:
        for line in myfile.readlines():
            line = line.strip('\n')
            if "EventID" in line:
                keyID = re.findall(r"Key=\"(.+?)\"", line)[0]
                eventID = re.findall(r"EventID=\"(.+?)\"", line)[0]
                RelT = re.findall(r"RelT=\"(.+?)\"", line)[0]
                #统计每种按键按下次数
                if eventID == "100":
                    da.totalDown = da.totalDown + 1
                    if keyID == "30002":
                        da.shangTimes = da.shangTimes + 1
                    if keyID == "30003":
                        da.xiaTimes = da.xiaTimes + 1
                    if keyID == "30004":
                        da.zuoTimes = da.zuoTimes + 1
                    if keyID == "30006":
                        da.youTimes = da.youTimes + 1
                    if keyID == "30009":
                        da.zuoshangTimes = da.zuoshangTimes + 1
                    if keyID == "30007":
                        da.zuoxiaTimes = da.zuoxiaTimes + 1
                    if keyID == "30010":
                        da.youshangTimes = da.youshangTimes + 1
                    if keyID == "30008":
                        da.youxiaTimes = da.youxiaTimes + 1
                    if keyID == "30005":
                        da.blank = da.blank + 1

                #统计按键弹起次数
                if eventID == "101":
                    da.totalUp = da.totalUp + 1

                #统计Z次数
                if keyID == "32010":
                    da.ZTimes = da.ZTimes + 1

                #统计X次数
                if keyID == "32011":
                    da.XTimes = da.XTimes + 1

                #统计+次数
                if keyID == "32016":
                    da.plusTimes = da.plusTimes + 1

                #统计-次数
                if keyID == "32015":
                    da.minusTimes = da.minusTimes + 1

    for i in range(0,len(listKeyInfo)):
        #记录基础信息
        #统计未弹起次数
        if listKeyInfo[i].recordUpTime == "/":
            da.realNotUp = da. realNotUp + 1
        # 统计按下靠后次数
        if listKeyInfo[i].stateDown == "靠后":
            da.postDownTimes = da.postDownTimes + 1
        # 统计弹起靠后次数
        if listKeyInfo[i].stateUp == "靠后":
            da.postUpTimes = da.postUpTimes + 1
        # 统计双排次数
        if listKeyInfo[i].isSynchronize == "Yes":
            da.synchronizeTimes = da.synchronizeTimes + 1
        # 统计传统类双空格次数（只在传统类模式生效）
        if listKeyInfo[i].downMinusDown != '/':
            if listKeyInfo[i].keyDownType == "空格" and int(listKeyInfo[i].downMinusDown) <= 40 and listKeyInfo[i - 1].netTime != '/' :
                da.doubleBlanksTimes = da.doubleBlanksTimes + 1

    #初始化按键步长统计
    lessOrEqual10 = 0
    lessOrEqual20 = 0
    lessOrEqual30 = 0
    lessOrEqual40 = 0
    lessOrEqual50 = 0
    lessOrEqual60 = 0
    lessOrEqual70 = 0
    lessOrEqual80 = 0
    lessOrEqual90 = 0
    lessOrEqual100 = 0
    lessOrEqual110 = 0
    lessOrEqual120 = 0
    lessOrEqual130 = 0
    lessOrEqual140 = 0
    lessOrEqual150 = 0
    lessOrEqual160 = 0
    lessOrEqual170 = 0
    lessOrEqual180 = 0
    lessOrEqual190 = 0
    lessOrEqual200 = 0
    lessOrEqual300 = 0
    lessOrEqual400 = 0
    lessOrEqual500 = 0
    greaterThan500 = 0

    listRes = []    #步长统计list
    for i in range(0,len(listKeyInfo)):
        if listKeyInfo[i].timeDiffer == '/':    #如果按键未弹起 跳过进行下一个
            continue
        if listKeyInfo[i].timeDiffer != '/':    #如果按键正常，则根据条件逐一进行统计
            if int(listKeyInfo[i].timeDiffer) <= 10:
                lessOrEqual10 = lessOrEqual10 + 1
            if int(listKeyInfo[i].timeDiffer) > 10 and int(listKeyInfo[i].timeDiffer) <= 20:
                lessOrEqual20 = lessOrEqual20 + 1
            if int(listKeyInfo[i].timeDiffer) > 20 and int(listKeyInfo[i].timeDiffer) <= 30:
                lessOrEqual30 = lessOrEqual30 + 1
            if int(listKeyInfo[i].timeDiffer) > 30 and int(listKeyInfo[i].timeDiffer) <= 40:
                lessOrEqual40 = lessOrEqual40 + 1
            if int(listKeyInfo[i].timeDiffer) > 40 and int(listKeyInfo[i].timeDiffer) <= 50:
                lessOrEqual50 = lessOrEqual50 + 1
            if int(listKeyInfo[i].timeDiffer) > 50 and int(listKeyInfo[i].timeDiffer) <= 60:
                lessOrEqual60 = lessOrEqual60 + 1
            if int(listKeyInfo[i].timeDiffer) > 60 and int(listKeyInfo[i].timeDiffer) <= 70:
                lessOrEqual70 = lessOrEqual70 + 1
            if int(listKeyInfo[i].timeDiffer) > 70 and int(listKeyInfo[i].timeDiffer) <= 80:
                lessOrEqual80 = lessOrEqual80 + 1
            if int(listKeyInfo[i].timeDiffer) > 80 and int(listKeyInfo[i].timeDiffer) <= 90:
                lessOrEqual90 = lessOrEqual90 + 1
            if int(listKeyInfo[i].timeDiffer) > 90 and int(listKeyInfo[i].timeDiffer) <= 100:
                lessOrEqual100 = lessOrEqual100 + 1


            if int(listKeyInfo[i].timeDiffer) > 100 and int(listKeyInfo[i].timeDiffer) <= 110:
                lessOrEqual110 = lessOrEqual110 + 1
            if int(listKeyInfo[i].timeDiffer) > 110 and int(listKeyInfo[i].timeDiffer) <= 120:
                lessOrEqual120 = lessOrEqual120 + 1
            if int(listKeyInfo[i].timeDiffer) > 120 and int(listKeyInfo[i].timeDiffer) <= 130:
                lessOrEqual130 = lessOrEqual130 + 1
            if int(listKeyInfo[i].timeDiffer) > 130 and int(listKeyInfo[i].timeDiffer) <= 140:
                lessOrEqual140 = lessOrEqual140 + 1
            if int(listKeyInfo[i].timeDiffer) > 140 and int(listKeyInfo[i].timeDiffer) <= 150:
                lessOrEqual150 = lessOrEqual150 + 1
            if int(listKeyInfo[i].timeDiffer) > 150 and int(listKeyInfo[i].timeDiffer) <= 160:
                lessOrEqual160 = lessOrEqual160 + 1
            if int(listKeyInfo[i].timeDiffer) > 160 and int(listKeyInfo[i].timeDiffer) <= 170:
                lessOrEqual170 = lessOrEqual170 + 1
            if int(listKeyInfo[i].timeDiffer) > 170 and int(listKeyInfo[i].timeDiffer) <= 180:
                lessOrEqual180 = lessOrEqual180 + 1
            if int(listKeyInfo[i].timeDiffer) > 180 and int(listKeyInfo[i].timeDiffer) <= 190:
                lessOrEqual190 = lessOrEqual190 + 1
            if int(listKeyInfo[i].timeDiffer) > 190 and int(listKeyInfo[i].timeDiffer) <= 200:
                lessOrEqual200 = lessOrEqual200 + 1


            if int(listKeyInfo[i].timeDiffer) > 200 and int(listKeyInfo[i].timeDiffer) <= 300:
                lessOrEqual300 = lessOrEqual300 + 1
            if int(listKeyInfo[i].timeDiffer) > 300 and int(listKeyInfo[i].timeDiffer) <= 400:
                lessOrEqual400 = lessOrEqual400 + 1
            if int(listKeyInfo[i].timeDiffer) > 400 and int(listKeyInfo[i].timeDiffer) <= 500:
                lessOrEqual500 = lessOrEqual500 + 1
            if int(listKeyInfo[i].timeDiffer) > 500:
                greaterThan500 = greaterThan500 + 1

    #将统计结果依次加入到listRes中
    listRes.append(lessOrEqual10)
    listRes.append(lessOrEqual20)
    listRes.append(lessOrEqual30)
    listRes.append(lessOrEqual40)
    listRes.append(lessOrEqual50)
    listRes.append(lessOrEqual60)
    listRes.append(lessOrEqual70)
    listRes.append(lessOrEqual80)
    listRes.append(lessOrEqual90)
    listRes.append(lessOrEqual100)
    listRes.append(lessOrEqual110)
    listRes.append(lessOrEqual120)
    listRes.append(lessOrEqual130)
    listRes.append(lessOrEqual140)
    listRes.append(lessOrEqual150)
    listRes.append(lessOrEqual160)
    listRes.append(lessOrEqual170)
    listRes.append(lessOrEqual180)
    listRes.append(lessOrEqual190)
    listRes.append(lessOrEqual200)
    listRes.append(lessOrEqual300)
    listRes.append(lessOrEqual400)
    listRes.append(lessOrEqual500)
    listRes.append(greaterThan500)

    return da,listRes   #返回 按键统计类对象 和 步长统计数据



# --------------------------写数据统计EXCEL----------------------------
#   描述： 设置“按键统计”EXCEL表格
#   参数：da：设置好的按键统计类对象，listRes：步长统计数据
#   返回值：
def writeEXCEL2(da,listRes):
    sht0 = wb.sheets.add() # 新增一个表格
    sht0.name = "数据统计"  #设置表格名
    if int(da.totalDown) == 0:  #如果是个内存挂直接不看了
        messagebox.showinfo("提示", "本录像按下次数为0")
        sys.exit(0)
    #如果不是很垃圾的内存挂则继续
    #接下来都是写入EXCEL
    head1 = ["类型","数值","","类型","次数","频率"]
    sht0.range('A1').value = head1
    sht0.range('A2').value = ["总按下次数",da.totalDown]
    sht0.range('A3').value = ["总弹起次数",da.totalUp]
    sht0.range('A4').value = ["实际未弹起次数",da.realNotUp]
    sht0.range('A5').value = ["“ Z ” 次数",da.ZTimes]
    sht0.range('A6').value = ["“ X ” 次数",da.XTimes]
    sht0.range('A7').value = ["“ + ” 次数",da.plusTimes]
    sht0.range('A8').value = ["“ - ” 次数",da.minusTimes]
    sht0.range('A9').value = ["靠前次数(按下)",da.preDownTimes]
    sht0.range('A10').value = ["靠后次数(按下)",da.postDownTimes]
    sht0.range('A11').value = ["靠前次数(弹起)",da.preUpTimes]
    sht0.range('A12').value = ["靠后次数(弹起)",da.postUpTimes]
    sht0.range('A13').value = ["双排次数",da.synchronizeTimes]
    sht0.range('A14').value = ["双空格次数(传统类)",da.doubleBlanksTimes]


    sht0.range('D2').value = ["↑",da.shangTimes,str(da.shangTimes / da.totalDown * 100)+"%"]
    sht0.range('D3').value = ["↓",da.xiaTimes,str(da.xiaTimes / da.totalDown * 100)+"%"]
    sht0.range('D4').value = ["←",da.zuoTimes,str(da.zuoTimes / da.totalDown * 100)+"%"]
    sht0.range('D5').value = ["→",da.youTimes,str(da.youTimes / da.totalDown * 100)+"%"]
    sht0.range('D6').value = ["↖",da.zuoshangTimes,str(da.zuoshangTimes / da.totalDown * 100)+"%"]
    sht0.range('D7').value = ["↙",da.zuoxiaTimes,str(da.zuoxiaTimes / da.totalDown * 100)+"%"]
    sht0.range('D8').value = ["↗",da.youshangTimes,str(da.youshangTimes / da.totalDown * 100)+"%"]
    sht0.range('D9').value = ["↘",da.youxiaTimes,str(da.youxiaTimes / da.totalDown * 100)+"%"]
    sht0.range('D10').value = ["空格",da.blank,str(da.blank / da.totalDown * 100)+"%"]


    sht0.range('H1').value = ["步长","次数"]
    sht0.range('H2').value = ["<=10",listRes[0]]
    sht0.range('H3').value = ["<=20",listRes[1]]
    sht0.range('H4').value = ["<=30",listRes[2]]
    sht0.range('H5').value = ["<=40",listRes[3]]
    sht0.range('H6').value = ["<=50",listRes[4]]
    sht0.range('H7').value = ["<=60",listRes[5]]
    sht0.range('H8').value = ["<=70",listRes[6]]
    sht0.range('H9').value = ["<=80",listRes[7]]
    sht0.range('H10').value = ["<=90",listRes[8]]
    sht0.range('H11').value = ["<=100",listRes[9]]
    sht0.range('H12').value = ["<=110",listRes[10]]
    sht0.range('H13').value = ["<=120",listRes[11]]
    sht0.range('H14').value = ["<=130",listRes[12]]
    sht0.range('H15').value = ["<=140",listRes[13]]
    sht0.range('H16').value = ["<=150",listRes[14]]
    sht0.range('H17').value = ["<=160",listRes[15]]
    sht0.range('H18').value = ["<=170",listRes[16]]
    sht0.range('H19').value = ["<=180",listRes[17]]
    sht0.range('H20').value = ["<=190",listRes[18]]
    sht0.range('H21').value = ["<=200",listRes[19]]
    sht0.range('H22').value = ["<=300",listRes[20]]
    sht0.range('H23').value = ["<=400",listRes[21]]
    sht0.range('H24').value = ["<=500",listRes[22]]
    sht0.range('H25').value = [">500",listRes[23]]

    sht0.range('A1:I1').api.Font.Bold = True  # 加粗
    sht0.range('A1:I1').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('A2:A14').api.Font.Bold = True  # 加粗
    sht0.range('A2:A14').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('D2:D10').api.Font.Bold = True  # 加粗
    sht0.range('D2:D10').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('H2:H25').api.Font.Bold = True  # 加粗
    sht0.range('H2:H25').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)

    sht0.range("A1:F14").columns.autofit()

    wb.save(portion[0]+'数据统计.xls')

#------------------------------同时获取有效无效步长--------------------------------#
#   描述：实现步长有效和无效同时计算
#   参数：listKeyInfo：按键统计
#   返回值：listTimeDiffer：长度为空格次数，宽度为2的numpy数组
def getboth(listKeyInfo):
    lenj = len(listKeyInfo) #获取按键总长度

    j = 0   #循环按键统计
    blankCount = 0  #空格次数统计
    while j < lenj:
        if listKeyInfo[j].keyDownType == '空格':
            blankCount = blankCount + 1
        j = j + 1

    # 新建一个长度为空格次数，宽度为2的numpy数组
    # 第一列是有效步长，第二列是无效步长
    # 当第一列是有效步长时，第二列用0替代
    # 当第二列是无效步长时，第一列用0替代
    listTimeDiffer = np.zeros(shape=(blankCount, 2))


    i = 0 #循环按键统计
    j = 0 #创建有效无效列表
    while i < lenj:
        #有效
        if str(listKeyInfo[i].netTime) != '/':  #如果有网络时间
            listTimeDiffer[j][1] = 0    #设置同排无效步长为0
            if listKeyInfo[i].timeDiffer != '/':    #如果按键正常弹起
                listTimeDiffer[j][0] = int(listKeyInfo[i].timeDiffer)   #设置有效步长
                j = j + 1
            else:
                listTimeDiffer[j][0] = 0    #如果按键未弹起，设置为0（这是为了方便画图）
                j = j + 1

        #无效
        if str(listKeyInfo[i].keyDownType) == '空格' and str(listKeyInfo[i].netTime) == '/':  #如果按键是个空格并且没有网络时间
            #那么他是个无效空格
            listTimeDiffer[j][0] = 0    #设置同排有效步长为0
            if listKeyInfo[i].timeDiffer != '/':    #如果按键正常弹起
                listTimeDiffer[j][1] = int(listKeyInfo[i].timeDiffer)   #设置无效步长
                j = j + 1
            else:
                listTimeDiffer[j][1] = 0    #如果按键未弹起，设置为0（这是为了方便画图）
                j = j + 1
        i = i + 1

    return listTimeDiffer

#------------------------------画图--------------------------------#
#   描述：画各种图
#   参数：listKeyInfo：按键统计，judgeList：判定统计
#   返回值：
def draw(listKeyInfo,judgeList):
    portion = os.path.splitext(fileName)  # portion[0] 文件名
    #------------------------------画步长图--------------------------------
    # 初始化纵坐标
    y = {}  # 有效
    j = 0
    for i in range(0,len(listKeyInfo)):
        if listKeyInfo[i].netTime != '/':
            if listKeyInfo[i].timeDiffer != '/':
                y[j] = int(listKeyInfo[i].timeDiffer)
                j = j + 1
            else:
                y[j] = 0
                j = j + 1
    y_list = list(y.values())  # 转为list类型

    # 初始化横坐标
    x = {}
    i = 1
    while True:
        x[i] = i
        i = i + 1
        if i > len(y_list):
            break
    x_list = list(x.values())  # 转为list类型

    gm = getGameMode()
    #传统的图步长加了数值
    if gm == '6' or gm == '12' or gm == '17' or gm == '18' or gm == '7':
        bar1 = (
            Bar()
                .add_xaxis(x_list)
                .add_yaxis('有效步长', y_list, itemstyle_opts=opts.ItemStyleOpts(color='#DDA0DD'))  ##DDA0DD
                .set_series_opts(label_opts=opts.LabelOpts(is_show=True, font_size=12),
                                 markline_opts=opts.MarkLineOpts(data=[opts.MarkLineItem(type_="average", name="平均值"),
                                                                       opts.MarkLineItem(type_="max", name="最大值"),
                                                                       opts.MarkLineItem(type_="min", name="最小值")
                                                                       ],
                                                                 linestyle_opts=opts.LineStyleOpts(type_='dashed',
                                                                                                   opacity=0.5,
                                                                                                   color='black')
                                                                 )
                                 )
                .set_global_opts(title_opts=opts.TitleOpts(title='玩家步长图', subtitle=imageSubtitle,pos_top="48%"),
                                 legend_opts=opts.LegendOpts(pos_top="48%"),
                                 datazoom_opts=opts.DataZoomOpts(type_="inside"),
                                 yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True))
                                 )
                .set_series_opts(label_opts=opts.LabelOpts(is_show=True, font_size=9))
        )
    else:
        bar1 = (
            Bar()
                .add_xaxis(x_list)
                .add_yaxis('有效步长', y_list, itemstyle_opts=opts.ItemStyleOpts(color='#DDA0DD'))  ##DDA0DD
                .set_series_opts(label_opts=opts.LabelOpts(is_show=True, font_size=12),
                                 markline_opts=opts.MarkLineOpts(data=[opts.MarkLineItem(type_="average", name="平均值"),
                                                                       opts.MarkLineItem(type_="max", name="最大值"),
                                                                       opts.MarkLineItem(type_="min", name="最小值")
                                                                       ],
                                                                 linestyle_opts=opts.LineStyleOpts(type_='dashed',
                                                                                                   opacity=0.5,
                                                                                                   color='black')
                                                                 )
                                 )
                .set_global_opts(title_opts=opts.TitleOpts(title='玩家步长图', subtitle=imageSubtitle, pos_top="48%"),
                                 legend_opts=opts.LegendOpts(pos_top="48%"),
                                 datazoom_opts=opts.DataZoomOpts(type_="inside"),
                                 yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True))
                                 )
                .set_series_opts(label_opts=opts.LabelOpts(is_show=False, font_size=9))
        )


    songInfo = getSongInfo()  # [level,songName,BPM]
    BPM = float(songInfo[2])  # 需要float类型
    listDetail = np.array(calculateBPM(BPM, 0))  # 计算BPM
    Perfect = listDetail[1]  # [OnePat Perfect Great Good Bad]
    listJDdata = []
    for data in judgeList:
        if data.judgeType != '/':
            listJDdata.append(data.judgeData)

    # -------------------------------------画波动图----------------------------------#
    # 初始化横坐标
    x = {}
    i = 0
    j = 0
    while True:
        x[i] = i
        i = i + 1
        if i == len(listJDdata):
            break
    x_list = list(x.values())  # 转为list类型

    # 初始化纵坐标
    y = {}
    i = 0
    j = 0
    while i < len(listJDdata):
        y[i] = listJDdata[i]
        i = i + 1
    y_list = list(y.values())  # 转为list类型

    # P点范围
    p1 = {}
    i = 0
    while True:
        p1[i] = Perfect
        i = i + 1
        if i > len(listJDdata):
            break
    p1_list = list(p1.values())  # 转为list类型

    p2 = {}
    i = 0
    while True:
        p2[i] = np.negative(Perfect)
        i = i + 1
        if i > len(listJDdata):
            break
    p2_list = list(p2.values())  # 转为list类型

    # 颜色备选
    #   波动线     波动背景     步长
    #  #0000EE   #DDA0DD    #DDA0DD   蓝粉
    #  #000000   #FF0000    #FF0000   黑红：俞北
    #  #FF00FF   #FF8C00    #FF8C00   紫橙
    #  #0000CD   #7FFFAA    #7FFFAA   蓝绿
    #  #FF1493   #87CEFA    #87CEFA   天蓝粉
    #  #DB9B66   #FFD8B9    #FFD8B9   橙粉：忍冬专属色
    c = (
        Line()
            .add_xaxis(x_list)
            .add_yaxis("波动", y_list, label_opts=opts.LabelOpts(is_show=False),
                       itemstyle_opts=opts.ItemStyleOpts(color='#0000EE'))
            .add_yaxis("P点范围1", p1_list, areastyle_opts=opts.AreaStyleOpts(opacity=0.5),
                       label_opts=opts.LabelOpts(is_show=False), is_symbol_show=False,
                       itemstyle_opts=opts.ItemStyleOpts(color='#DDA0DD'))
            .add_yaxis("P点范围2", p2_list, areastyle_opts=opts.AreaStyleOpts(opacity=0.5),
                       label_opts=opts.LabelOpts(is_show=False), is_symbol_show=False,
                       itemstyle_opts=opts.ItemStyleOpts(color='#DDA0DD'))
            .set_series_opts(label_opts=opts.LabelOpts(is_show=True, font_size=6))
            # .set_series_opts(label_opts=opts.LabelOpts(is_show=False, font_size=14),
            #                  markline_opts=opts.MarkLineOpts(data=[opts.MarkLineItem(type_="average", name="平均值"),
            #                                                        opts.MarkLineItem(type_="max", name="最大值"),
            #                                                        opts.MarkLineItem(type_="min", name="最小值")
            #                                                        ],
            #                                                  linestyle_opts=opts.LineStyleOpts(type_='dashed',opacity=0.5,color='black')
            #                                                  )
            #                  )
            .set_global_opts(title_opts=opts.TitleOpts(title="玩家波动图", subtitle=imageSubtitle),
                             datazoom_opts=opts.DataZoomOpts(type_="inside"),
                             yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True)
                                                      #,
    #                                                   min_='-15',   #设置y轴最小值
    #                                                   max_='15'     #设置y轴最大值  用于有些录像太飘
                                                      ),
                             toolbox_opts=opts.ToolboxOpts(feature={"saveAsImage": {},
                                                                    "restore": {},
                                                                    "magicType": {"show": True,
                                                                                  "type": ["line", "bar"]},
                                                                 "dataView": {}})
                             )
    )
    # c = (
    #     Line()
    #         .add_xaxis(x_list)
    #         .add_yaxis("波动", y_list, label_opts=opts.LabelOpts(is_show=False),
    #                    itemstyle_opts=opts.ItemStyleOpts(color='#0000EE'))
    #         .add_yaxis("P点范围1", p1_list, areastyle_opts=opts.AreaStyleOpts(opacity=0.5),
    #                    label_opts=opts.LabelOpts(is_show=False), is_symbol_show=False,
    #                    itemstyle_opts=opts.ItemStyleOpts(color='#DDA0DD'))
    #         .add_yaxis("P点范围2", p2_list, areastyle_opts=opts.AreaStyleOpts(opacity=0.5),
    #                    label_opts=opts.LabelOpts(is_show=False), is_symbol_show=False,
    #                    itemstyle_opts=opts.ItemStyleOpts(color='#DDA0DD'))
    #         .set_series_opts(label_opts=opts.LabelOpts(is_show=True, font_size=6))
    #         # .set_series_opts(label_opts=opts.LabelOpts(is_show=False, font_size=14),
    #         #                  markline_opts=opts.MarkLineOpts(data=[opts.MarkLineItem(type_="average", name="平均值"),
    #         #                                                        opts.MarkLineItem(type_="max", name="最大值"),
    #         #                                                        opts.MarkLineItem(type_="min", name="最小值")
    #         #                                                        ],
    #         #                                                  linestyle_opts=opts.LineStyleOpts(type_='dashed',opacity=0.5,color='black')
    #         #                                                  )
    #         #                  )
    #         .set_global_opts(title_opts=opts.TitleOpts(title="玩家波动图", subtitle=imageSubtitle),
    #                          datazoom_opts=opts.DataZoomOpts(type_="inside"),
    #                          yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True),
    #                                                   min_='-15',   #设置y轴最小值
    #                                                   max_='15'     #设置y轴最大值  用于有些录像太飘
    #                                                   ),
    #                          toolbox_opts=opts.ToolboxOpts(feature={"saveAsImage": {},
    #                                                                 "restore": {},
    #                                                                 "magicType": {"show": True,
    #                                                                               "type": ["line", "bar"]},
    #                                                              "dataView": {}})
    #                          )
    # )
    #c.render(portion[0] + '波动图.html')
    grid = (
        Grid()
            # 通过设置图形相对位置，来调整是整个并行图是竖直放置，还是水平放置
            .add(c, grid_opts=opts.GridOpts(pos_bottom="60%"))
            .add(bar1, grid_opts=opts.GridOpts(pos_top="60%"))
            .render(portion[0] + "全图.html")
    )


    gm = getGameMode()#获取模式代码
    # 如果是传统类，则多画有效无效图、按键步长图
    if gm == '6' or gm == '12' or gm == '17' or gm == '18' or gm == '7':
        # ---------------------------------画有效无效步长图-------------------------------#
        listTimeDiffer = getboth(listKeyInfo)
        # 初始化横坐标
        x = {}
        i = 1
        while True:
            x[i] = i
            i = i + 1
            if i > len(listTimeDiffer):
                break
        x_list = list(x.values())  # 转为list类型

        # 初始化有效步长
        y_valid = {}
        i = 0
        while i < len(listTimeDiffer):
            y_valid[i] = listTimeDiffer[i][0]
            i = i + 1
        y_valid_list = list(y_valid.values())  # 转为list类型

        y_invalid = {}
        i = 0
        while i < len(listTimeDiffer):
            y_invalid[i] = listTimeDiffer[i][1]
            i = i + 1
        y_invalid_list = list(y_invalid.values())  # 转为list类型
        bar2 = (
            Bar()
                .add_xaxis(x_list)
                .add_yaxis('有效步长', y_valid_list, itemstyle_opts=opts.ItemStyleOpts(color='red'))
                .add_yaxis('无效步长', y_invalid_list, itemstyle_opts=opts.ItemStyleOpts(color='blue'))
                .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
                .set_global_opts(title_opts=opts.TitleOpts(title="玩家有效无效步长图", subtitle=imageSubtitle),
                                 datazoom_opts=opts.DataZoomOpts(type_="inside"),
                                 yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True))
                                 )
        )
        bar2.render(portion[0] + '有效无效步长.html')


        # ---------------------------------画按键步长图-------------------------------#
        # 计算按键list
        listKeyDirectionDiffer = []
        i = 0
        while i < len(listKeyInfo):
            if listKeyInfo[i].keyDownType != '空格':
                if listKeyInfo[i].timeDiffer != '/':
                    listKeyDirectionDiffer.append(int(listKeyInfo[i].timeDiffer))
                else:
                    listKeyDirectionDiffer.append(0)
            i = i + 1

        # 初始化横坐标
        x = {}
        i = 1
        while i < len(listKeyDirectionDiffer):
            x[i] = i
            i = i + 1
        x_list = list(x.values())  # 转为list类型

        # 初始化按键步长
        y_valid = {}
        i = 0
        while i < len(listKeyDirectionDiffer):
            y_valid[i] = listKeyDirectionDiffer[i]
            i = i + 1
        y_valid_list = list(y_valid.values())  # 转为list类型

        bar3 = (
            Bar()
                .add_xaxis(x_list)
                .add_yaxis('方向键步长', y_valid_list, itemstyle_opts=opts.ItemStyleOpts(color='green'))
                .set_global_opts(title_opts=opts.TitleOpts(title="玩家方向键步长", subtitle=imageSubtitle),
                                 datazoom_opts=opts.DataZoomOpts(type_="inside"),
                                 yaxis_opts=opts.AxisOpts(splitline_opts=opts.SplitLineOpts(is_show=True))
                                 )
                .set_series_opts(label_opts=opts.LabelOpts(is_show=False, font_size=12),
                                 markline_opts=opts.MarkLineOpts(data=[opts.MarkLineItem(type_="average", name="平均值"),
                                                                       opts.MarkLineItem(type_="max", name="最大值"),
                                                                       opts.MarkLineItem(type_="min", name="最小值")
                                                                       ],
                                                                 linestyle_opts=opts.LineStyleOpts(type_='dashed',
                                                                                                   opacity=0.5,
                                                                                                   color='black')
                                                                 )
                                 )
        )


        bar3.render(portion[0] + '方向按键.html')

        # page1 = Page(layout=Page.DraggablePageLayout)
        # # 将多个不同的数据视图，封装到一个page对象中
        # page1.add(
        #     bar3
        # )
        # # 将page1的内容输出到一个html文件中
        # page1.render(portion[0] + 'test.html')

#--------------------------将string转为int并变为array----------------------------
def strToInt(string):
    res = np.array(list(map(int, string)))
    return res


# ---------------------------------获取对局结果------------------------------------
#   描述： 用于获取对局结果
#   参数：
#   返回值：listAllPlayerInfo：[[总分,最大连P,Perfect,Great,Good,bad,Miss],....,[总分,最大连P,Perfect,Great,Good,bad,Miss]]
def getResults():
    # print("获取对局结果")
    # #定义对局结果数组
    listGameInfo = getGameInfo()
    playerCount = listGameInfo[4]
    playerRes = np.zeros(shape=(playerCount, 8))  # 总分 最大连P 最大连C Perfect Great Good bad Miss

    # 遍历Score
    with open(fileName, 'r', encoding='gb18030') as myfile:
        i = 0
        for line in myfile:
            line = line.strip('\n')  # 以回车键为分隔符
            if "<Score>" in line:
                Score = strToInt(re.findall(r"<Score>(.+?)</Score>", line))
                playerRes[i][0] = Score
                i = i + 1
    # 遍历MaxPer
    with open(fileName, 'r', encoding='gb18030') as myfile:
        i = 0
        for line in myfile:
            line = line.strip('\n')  # 以回车键为分隔符
            if " <MaxPer>" in line:
                MaxPer = strToInt(re.findall(r"<MaxPer>(.+?)</MaxPer>", line))
                playerRes[i][1] = MaxPer
                i = i + 1

    # 遍历MaxCombo
    with open(fileName, 'r', encoding='gb18030') as myfile:
        i = 0
        for line in myfile:
            line = line.strip('\n')  # 以回车键为分隔符
            if " <MaxCombo>" in line:
                MaxPer = strToInt(re.findall(r"<MaxCombo>(.+?)</MaxCombo>", line))
                playerRes[i][2] = MaxPer
                i = i + 1

    # 遍历Perfect
    with open(fileName, 'r', encoding='gb18030') as myfile:
        i = 0
        for line in myfile:
            line = line.strip('\n')  # 以回车键为分隔符
            if " <Perfect>" in line:
                Perfect = strToInt(re.findall(r"<Perfect>(.+?)</Perfect>", line))
                playerRes[i][3] = Perfect
                i = i + 1
    # 遍历Great
    with open(fileName, 'r', encoding='gb18030') as myfile:
        i = 0
        for line in myfile:
            line = line.strip('\n')  # 以回车键为分隔符
            if " <Cool>" in line:
                Great = strToInt(re.findall(r"<Cool>(.+?)</Cool>", line))
                playerRes[i][4] = Great
                i = i + 1
    # 遍历Good
    with open(fileName, 'r', encoding='gb18030') as myfile:
        i = 0
        for line in myfile:
            line = line.strip('\n')  # 以回车键为分隔符
            if " <Good>" in line:
                Good = strToInt(re.findall(r"<Good>(.+?)</Good>", line))
                playerRes[i][5] = Good
                i = i + 1
    # 遍历Bad
    with open(fileName, 'r', encoding='gb18030') as myfile:
        i = 0
        for line in myfile:
            line = line.strip('\n')  # 以回车键为分隔符
            if " <Bad>" in line:
                Bad = strToInt(re.findall(r"<Bad>(.+?)</Bad>", line))
                playerRes[i][6] = Bad
                i = i + 1
    # 遍历Miss
    with open(fileName, 'r', encoding='gb18030') as myfile:
        i = 0
        for line in myfile:
            line = line.strip('\n')  # 以回车键为分隔符
            if " <Miss>" in line:
                Miss = strToInt(re.findall(r"<Miss>(.+?)</Miss>", line))
                playerRes[i][7] = Miss
                i = i + 1
    # 总分 最大连P Perfect Great Good bad Miss
    return playerRes

# def writeEXCEL4(lenNetEvent,listGameInfo,listAllPlayerInfo,playerRes,judgeList):

#对局结果
def writeEXCEL4(lenNetEvent,listGameInfo,listAllPlayerInfo,playerRes,judgeList):

    portion = os.path.splitext(fileName)  # portion[0] 文件名

    #head1 录像玩家信息
    head1 = ["录像时间","录像者","对局模式","玩家站位"]
    RecordTime = listGameInfo[0]
    author     = listGameInfo[1]
    GameMode    = listGameInfo[2]
    PlayerIndex = listGameInfo[3]
    playerCount = listGameInfo[4]
    content1 = [RecordTime,author,GameMode,PlayerIndex]

    #AP数量
    AP = lenNetEvent
    #head2 歌曲信息
    head2 = ["歌曲编号","歌曲名字","精确BPM","AP","1拍数值","Perfect范围","Great范围","Good范围","Bad范围"]
    songInfo = getSongInfo()
    BPM = float(songInfo[2])
    listDetail = calculateBPM(BPM,0)
    content2 = [songInfo[0],songInfo[1],songInfo[2],AP,listDetail[0],listDetail[1],listDetail[2],listDetail[3],listDetail[4]]

    #head3 对局结果
    head3 = ["玩家序号","玩家昵称","舞团","性别","总分","最大连P","最大连C","Perfect","Great","Good","Bad","Miss","P率"]

    i = 0
    num = 0
    content3 = [] #给对局结果赋值
    while True:
        Name = listAllPlayerInfo[i][0]
        Band = listAllPlayerInfo[i][1]
        Model = listAllPlayerInfo[i][2]
        if i == int(PlayerIndex):
            Ppercent = str(np.round(int(playerRes[i][3]) / AP *100,2))+"%"  #只写录像者的P率
            content3.append([num,Name,Band,Model,playerRes[i][0],playerRes[i][1],playerRes[i][2],playerRes[i][3],playerRes[i][4],playerRes[i][5],playerRes[i][6],playerRes[i][7],Ppercent])
        else:
            content3.append([num,Name,Band,Model,playerRes[i][0],playerRes[i][1],playerRes[i][2],playerRes[i][3],playerRes[i][4],playerRes[i][5],playerRes[i][6],playerRes[i][7],'/'])
        num = num + 1
        i = i + 1
        if i == playerCount:
            break

    lenth = len(judgeList)
    #代码分析判定统计
    RealPer = 0
    RealGreat = 0
    RealGood = 0
    RealBad = 0
    RealMiss = 0
    i = 0
    while i < lenth:
        if judgeList[i].judgeType != "/":
            if judgeList[i].judgeType == "Perfect":
                RealPer = RealPer + 1
            if judgeList[i].judgeType == "Great":
                RealGreat = RealGreat + 1
            if judgeList[i].judgeType == "Good":
                RealGood = RealGood + 1
            if judgeList[i].judgeType == "Bad":
                RealBad = RealBad + 1
        i = i + 1
    RealMiss = int(AP) - RealPer - RealGreat - RealGood - RealBad
    head4 = ["", "Perfect", "Great", "Good", "Bad", "Miss"]

    PlayerIndex = int(PlayerIndex)
    content4 = ["系统记录", playerRes[PlayerIndex][3], playerRes[PlayerIndex][4], playerRes[PlayerIndex][5],
                         playerRes[PlayerIndex][6], playerRes[PlayerIndex][7]]

    content5 = ["代码分析", RealPer, RealGreat, RealGood, RealBad, RealMiss]

    PerAcc = 0
    GreatAcc = 0
    GoodAcc = 0
    BadAcc = 0
    MissAcc = 0
    # 误差 = 代码分析 - 系统记录
    PerAcc = np.negative(int(playerRes[PlayerIndex][3]) - RealPer)
    GreatAcc = np.negative(int(playerRes[PlayerIndex][4]) - RealGreat)
    GoodAcc = np.negative(int(playerRes[PlayerIndex][5]) - RealGood)
    BadAcc = np.negative(int(playerRes[PlayerIndex][6]) - RealBad)
    MissAcc = np.negative(int(playerRes[PlayerIndex][7]) - RealMiss)
    content6 = ["误差", PerAcc, GreatAcc, GoodAcc, BadAcc, MissAcc]
    sht0 = wb.sheets.add() # 新增一个表格
    sht0.name = "对局结果"

    # 写入数据
    sht0.range('A1').value = head1
    sht0.range('A2').value = content1
    sht0.range('A1').api.Font.Bold = True  # 加粗
    sht0.range('A1').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('B1').api.Font.Bold = True  # 加粗
    sht0.range('B1').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('C1').api.Font.Bold = True  # 加粗
    sht0.range('C1').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('D1').api.Font.Bold = True  # 加粗
    sht0.range('D1').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)

    sht0.range('A4').value = head2
    sht0.range('A5').value = content2
    sht0.range('A4').api.Font.Bold = True  # 加粗
    sht0.range('A4').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('B4').api.Font.Bold = True  # 加粗
    sht0.range('B4').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('C4').api.Font.Bold = True  # 加粗
    sht0.range('C4').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('D4').api.Font.Bold = True  # 加粗
    sht0.range('D4').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('E4').api.Font.Bold = True  # 加粗
    sht0.range('E4').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('F4').api.Font.Bold = True  # 加粗
    sht0.range('F4').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('G4').api.Font.Bold = True  # 加粗
    sht0.range('G4').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('H4').api.Font.Bold = True  # 加粗
    sht0.range('H4').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('I4').api.Font.Bold = True  # 加粗
    sht0.range('I4').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)



    sht0.range('A7').value = head3
    sht0.range('A8').value = content3
    sht0.range('A7').api.Font.Bold = True  # 加粗
    sht0.range('A7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('B7').api.Font.Bold = True  # 加粗
    sht0.range('B7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('C7').api.Font.Bold = True  # 加粗
    sht0.range('C7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('D7').api.Font.Bold = True  # 加粗
    sht0.range('D7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('E7').api.Font.Bold = True  # 加粗
    sht0.range('E7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('F7').api.Font.Bold = True  # 加粗
    sht0.range('F7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('G7').api.Font.Bold = True  # 加粗
    sht0.range('G7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('H7').api.Font.Bold = True  # 加粗
    sht0.range('H7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('I7').api.Font.Bold = True  # 加粗
    sht0.range('I7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('J7').api.Font.Bold = True  # 加粗
    sht0.range('J7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('K7').api.Font.Bold = True  # 加粗
    sht0.range('K7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('L7').api.Font.Bold = True  # 加粗
    sht0.range('L7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('M7').api.Font.Bold = True  # 加粗
    sht0.range('M7').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)

    sht0.range('A15').value = head4
    sht0.range('A16').value = content4
    sht0.range('A17').value = content5
    sht0.range('A18').value = content6
    sht0.range('B15').api.Font.Bold = True  # 加粗
    sht0.range('B15').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('C15').api.Font.Bold = True  # 加粗
    sht0.range('C15').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('D15').api.Font.Bold = True  # 加粗
    sht0.range('D15').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('E15').api.Font.Bold = True  # 加粗
    sht0.range('E15').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('F15').api.Font.Bold = True  # 加粗
    sht0.range('F15').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('A16').api.Font.Bold = True  # 加粗
    sht0.range('A16').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('A17').api.Font.Bold = True  # 加粗
    sht0.range('A17').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
    sht0.range('A18').api.Font.Bold = True  # 加粗
    sht0.range('A18').api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)



    sht0.range("A1:A25").columns.autofit()
    sht0.range("B1:B25").columns.autofit()
    sht0.range("C1:C25").columns.autofit()
    sht0.range("C1:C25").columns.autofit()
    sht0.range("D1:D25").columns.autofit()
    sht0.range("E1:E25").columns.autofit()
    sht0.range("F1:F25").columns.autofit()
    sht0.range("G1:G25").columns.autofit()
    sht0.range("H1:AH25").columns.autofit()
    sht0.range("I1:I25").columns.autofit()
    sht0.range("J1:J25").columns.autofit()
    sht0.range("K1:K25").columns.autofit()
    wb.save(portion[0]+'数据统计.xls')

#按键分割
def writeEXCEL5(listKeyInfo):
    gm = getGameMode()
    # 节奏
    if gm == '5' or gm == '11':
        return
    else:
        portion = os.path.splitext(fileName)  # portion[0] 文件名
        sht0 = wb.sheets.add()  # 新增一个表格
        sht0.name = "按键分割"
        flag = False
        listPartition  = []
        i = 1 #记录行数
        for eachKey in listKeyInfo:
            listPartition.append(eachKey.keyDownType)
            if eachKey.netTime != '/':
                sht0.range('A'+str(i)).value = listPartition
                if len(listPartition) > 26 and len(listPartition) <= 52:
                    cell = "A" + str(chr((len(listPartition)-26+64)))+str(i)
                if len(listPartition) > 52 and len(listPartition) <= 78:
                    cell = "B" + str(chr((len(listPartition)-52+64)))+str(i)
                if len(listPartition) > 78 and len(listPartition) <= 104:
                    cell = "C" + str(chr((len(listPartition)-78+64)))+str(i)
                if len(listPartition) > 104 and len(listPartition) <= 130:
                    cell = "D" + str(chr((len(listPartition)-104+64)))+str(i)
                if len(listPartition) > 130 :
                    cell = "E" + str(chr((len(listPartition)-130+64)))+str(i)

                if len(listPartition) <=26:
                    cell = chr(len(listPartition)+64)+str(i)
                #print(cell)
                sht0.range(cell).api.Font.Bold = True  # 加粗
                sht0.range(cell).api.Font.Color = 0x0000ff  # 设置为红色RGB(255,0,0)
                i = i + 1
                listPartition.clear()
        wb.save(portion[0]+'数据统计.xls')
#
# def draw2():
#     portion = os.path.splitext(fileName)  # portion[0] 文件名
#     table = Table()
#
#     headers = [None]
#     rows = [
#         ["Brisbane", 5905, 1857594, 1146.4],
#         ["Adelaide", 1295, 1158259, 600.5],
#         ["Darwin", 112, 120900, 1714.7],
#         ["Hobart", 1357, 205556, 619.5],
#         ["Sydney", 2058, 4336374, 1214.8],
#         ["Melbourne", 1566, 3806092, 646.9],
#         ["Perth", 5386, 1554769, 869.4],
#     ]
#     table.add(headers,rows)
#     table.set_global_opts(
#         title_opts=ComponentTitleOpts(title="玩家基础信息")
#     )
#     table.render(portion[0]+"玩家基础信息.html")

# 画同异步差图
def setHashMap(listKeyInfo):
    #------------------------------------异步差----------------------------------------------
    downMinusUp_hash = dict()
    i = 0
    while i < len(listKeyInfo):
        if listKeyInfo[i].downMinusUp != '/':
            if listKeyInfo[i].keyDownType != '空格':
                key = listKeyInfo[i].downMinusUp
                if downMinusUp_hash.get(key,False) == False:
                    downMinusUp_hash.setdefault(key, 1)
                else:
                    newvalue = downMinusUp_hash.get(key, False) + 1
                    downMinusUp_hash.update({key:newvalue})
        i = i + 1
    #按照值从小到大排序
    listdownMinusUp_hash = sorted(downMinusUp_hash.items(), key = lambda kv:(kv[1], kv[0]))
    portion = os.path.splitext(fileName)  # portion[0] 文件名
    i = 0
    attr = []
    v = []
    while i < len(listdownMinusUp_hash):
        attr.append(listdownMinusUp_hash[i][0])
        v.append(listdownMinusUp_hash[i][1])
        i = i +1

    pie1 = (
        Pie()
            .add(
                    "",
                    data_pair = [list(z) for z in zip(attr, v)],
                    center=[650, 300]
            )
            .set_global_opts(
                legend_opts=opts.LegendOpts(is_show=False),
                title_opts=opts.TitleOpts(title="方向键异步差",pos_left='50%',subtitle=imageSubtitle)
            )
    )
    #pie1.render(path=portion[0] + "异步差饼图.html")



    # ------------------------------------同步差----------------------------------------------
    downMinusDown_hash = dict()
    i = 0
    while i < len(listKeyInfo):
        if listKeyInfo[i].downMinusDown != '/':
            if listKeyInfo[i].keyDownType != '空格':
                key = listKeyInfo[i].downMinusDown
                if downMinusDown_hash.get(key, False) == False:
                    downMinusDown_hash.setdefault(key, 1)
                else:
                    newvalue = downMinusDown_hash.get(key, False) + 1
                    downMinusDown_hash.update({key: newvalue})
        i = i + 1
    # 按照值从小到大排序
    listdownMinusDown_hash = sorted(downMinusDown_hash.items(), key=lambda kv: (kv[1], kv[0]))
    portion = os.path.splitext(fileName)  # portion[0] 文件名
    i = 0
    attr = []
    v = []
    while i < len(listdownMinusDown_hash):
        attr.append(listdownMinusDown_hash[i][0])
        v.append(listdownMinusDown_hash[i][1])
        i = i + 1

    pie2 = (
        Pie()
            .add(
                "",
                data_pair=[list(z) for z in zip(attr, v)],
                center=[220, 300]
            )
            .set_global_opts(
                legend_opts=opts.LegendOpts(is_show=False,pos_left="0%",pos_top="55",orient='vertical'),
                title_opts=opts.TitleOpts(title="方向键同步差",subtitle=imageSubtitle)
            )

    )
    #pie2.render(path=portion[0] + "同步差饼图.html")

    grid = (
        Grid()
            # 通过设置图形相对位置，来调整是整个并行图是竖直放置，还是水平放置
            .add(pie1, grid_opts=opts.GridOpts(pos_left="96%"))
            .add(pie2, grid_opts=opts.GridOpts(pos_right="5%"))
            .render(portion[0] + "同异步差.html")
    )

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    fileName = filedialog.askopenfilename()  # 获得选择好的文件
    portion = os.path.splitext(fileName)    #分割文件名和后缀
    if fileName == '':  #如果未选择
        messagebox.showinfo("提示", "未选择录像文件。")
        sys.exit(0)
    if portion[1] != '.rcd':    #如果后缀不是.rcd
        messagebox.showinfo("提示", "你选择的不是炫舞录像文件！")
        sys.exit(0)
    else:
        time_start = time.time()    #开始计时
        listKeyInfo = readKey()     #先读取按键

        da,listRes = setSheet0(listKeyInfo) #设置按键数据统计
        writeEXCEL2(da,listRes)     #把按键数据统计写入EXCEL

        writeEXCEL5(listKeyInfo)    #写按键分割进EXCEL

        gm = getGameMode()  #获取模式代码
        # 传统类
        if gm == '6' or gm == '12' or gm == '17' or gm == '18' or gm == '7':
            listNetEvent = getNetEventTime1()   #获取网络时间
            setHashMap(listKeyInfo) #画同异步差图
        # 节奏
        if gm == '5' or gm == '11':
            listNetEvent = getNetEventTime1()   #获取网络时间

        # 炫舞模式
        if gm == '1' or gm == '8':
            listNetEvent = divdeModeX5(getNetEventTime2())  #获取网络时间并进行模式区分
        # 团队模式
        if gm == '2' or gm == '9' :
            listNetEvent = divdeModeTeam(getNetEventTime2())#   获取网络时间并进行模式区分

        judgeList = setJudgeList(listNetEvent, listKeyInfo)     #获取判定对象列表

        writeEXCEL3(judgeList)  #写判定统计

        draw(listKeyInfo,judgeList) #画所有需要的图

        listGameInfo = getGameInfo()    #获取对局信息
        listAllPlayerInfo = getAllPlayerInfo()  #获取玩家信息
        playerRes = getResults()    #获取玩家对局结果
        writeEXCEL4(len(listNetEvent), listGameInfo, listAllPlayerInfo, playerRes, judgeList)   #写对局基础信息

        time_end = time.time()  #计时结束
        messagebox.showinfo("提示", "分析结束，耗时：" + str(np.round(time_end - time_start, 3)) + "s")