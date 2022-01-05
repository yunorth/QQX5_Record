# QQ炫舞录像鉴定工具开源(python)

## 作者的话
游戏脚本太多了，通过此工具可查看玩家录像是否为脚本录像。

*如果有大佬看见请放过我这只卑微的蚂蚁，在学了在学了别骂了。*

师傅领进门，修行靠个人。

最新版工具：https://www.weiyun.com/disk/folder/49ddfe960b9833e04a1459a20cfb5aaf
后续停止更新

这是不懂事的小孩留给这个圈子最后的东西。

****
## 工具介绍
本工具适用于QQ炫舞端游游戏录像辅助全自动鉴定。

本工具使用Python语言编写，运行速度快，成品图清晰动态且细节。

以下传统、疯狂、自由统称为传统类。

作者QQ：2533285193

**功能如下：**

 1. 工具支持传统类、节奏、炫舞、团队模式的鉴定全图生成
 2. 工具支持传统类无效步长图、按键步长图、按键同异步差图生成
 3. 工具支持除VOS外大部分模式按键统计
 4. 工具自动标注步长最大值、最小值以及平均值
 5. 工具支持以上模式玩家基础信息查看
 6. 工具支持判定误差信息总结

****
## 适用人群

 - 学过录像代码基础的人
 - 学过怎么作图的人
 - 以及懂录像原理的人
 -  本工具目的在于提高看录像的速度，不是让你什么都不学直接看

****
## 使用方法
***使用前提：***

使用工具的前提是电脑上有安装[Office EXCEL](https://www.microsoft.com/en-ww/microsoft-365/microsoft-office?rtc=1) 或者 [WPS excel](https://www.wps.cn/)，这两个必须有一个。

浏览器的话，最好使用[谷歌浏览器](http://chrome.illzjp.cn/browser.html)看图，没有谷歌浏览器其他浏览器也可以，但是效果可能没有谷歌浏览器好。

本工具适用于Windows 10/7/XP/03

***使用步骤***

双击打开exe直接选取录像

![在这里插入图片描述](https://img-blog.csdnimg.cn/748787e53dd441938423ee4cb2940f28.png?x-oss-process=image/watermark,type_ZHJvaWRzYW5zZmFsbGJhY2s,shadow_50,text_Q1NETiBA5Yqx5b-X6KaB5LiKOTg1,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)


选取后会在与该录像相同的路径下生成多个数据文件

分别是：

 1. 录像鉴定全图.html
 3. 录像数据统计.xls
 4. 无效步长图.html（只有传统模式有）
 5. 按键步长图.html（只有传统模式有）
 6. 按键同异步差图.html（只有传统模式有）
![在这里插入图片描述](https://img-blog.csdnimg.cn/c75884bccbdb4bd7be44d3c84f573240.png#pic_center)


.html后缀的文件用浏览器打开，.xls后缀的文件用EXCEL打开

浏览器打开网页后可查看波动图，通过鼠标滚轮可缩放波动图细节，步长图自动标注

![在这里插入图片描述](https://img-blog.csdnimg.cn/887fad0ce0374f57b9c39889af688c30.png?x-oss-process=image/watermark,type_ZHJvaWRzYW5zZmFsbGJhY2s,shadow_50,text_Q1NETiBA5Yqx5b-X6KaB5LiKOTg1,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)
![在这里插入图片描述](https://img-blog.csdnimg.cn/c65b5f37665a46dc94a22ac10cf11185.png?x-oss-process=image/watermark,type_ZHJvaWRzYW5zZmFsbGJhY2s,shadow_50,text_Q1NETiBA5Yqx5b-X6KaB5LiKOTg1,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)
![在这里插入图片描述](https://img-blog.csdnimg.cn/c0ded4e3f1c54289955287d99f5b2920.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBA5L-e5YyX,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)
![在这里插入图片描述](https://img-blog.csdnimg.cn/481f292a20224d64b980f8aa31eb9887.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBA5L-e5YyX,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)

数据文件用EXCEL打开后可查看对局结果、判定统计、按键分割、数据统计、按键统计

![在这里插入图片描述](https://img-blog.csdnimg.cn/90659ad01b474214853f0c8ad1281b16.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBA5L-e5YyX,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)
![在这里插入图片描述](https://img-blog.csdnimg.cn/e35497cfc0ab47dcb56e565db49e0e57.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBA5L-e5YyX,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)
![在这里插入图片描述](https://img-blog.csdnimg.cn/71efa6fc0c74474995a814312c174eec.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBA5L-e5YyX,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)
![在这里插入图片描述](https://img-blog.csdnimg.cn/b99fe098e12548d9aff5b183ac9c55e4.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBA5L-e5YyX,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)
![在这里插入图片描述](https://img-blog.csdnimg.cn/556460e0f77948858eb14ee3730b208e.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBA5L-e5YyX,size_20,color_FFFFFF,t_70,g_se,x_16#pic_center)




