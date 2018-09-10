# count_clock_in
复旦大学暑期打卡统计器：基于腾讯问卷-微信/QQ登录签到

## 设计

### 主要特性

1. 根据qq号或者微信内部id，使腾讯问卷的被访者和班级的同学一一对应
2. 根据问卷的选项，自动判断同学是否回家以及是否签到
3. 每次签到自动增加日期，自动写Excel文件
4. 背景高亮两种特殊情况

### 缺点与改进

1. 每天任需要清空数据（虽然是下载数据时顺手操作的），可以加入对签到时间的判断来解决。
2. 对于特殊字符的不支持，第一天需要手动标记一些同学

## NOTES

### openpyxl

+ 单元格选择: https://www.jianshu.com/p/642456aa93e2
+ 单元格样式: https://blog.csdn.net/aishenghuomeidaoli/article/details/52165305

### mac快捷工具

+ 命令行打开Excel文件 `open -a Microsoft\ Excel.app 2018暑期住宿.xlsx`
+ alias化简命令, so easy!