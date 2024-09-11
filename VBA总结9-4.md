# VBA
## VBE

### 如何打开VBE环境?

#### 按＜Alt+F11＞组合键
#### 依次执行【开发者】→【Visual Basic编辑器】
#### 右键单击工作表标签， 执行【查看代码】菜单命令

### 有哪些窗口以及它们的用途?

#### 工程窗口: 管理模块
#### 属性窗口: 设置或更改属性
#### 代码窗口: 写代码的
#### 立即窗口: 调试代码的
#### 本地窗口: 调试代码，监视对象和变量的
#### 菜单栏
#### 工具栏
##### 引用: 如果需要操作其他APP, 如MODS, 需要在引用里勾选对应的库.

### 代码窗口相关的操作:

#### 鼠标左键双击对应的模块/工作簿/工作表 进入代码窗口
#### 光标需要在Sub 和 End Sub之间
#### 单步执行: F8
#### 直接执行: F5
#### 断点: F9
#### 注释: 英文的 " ' "


### 注意: **缩进**, 虽然缩进不影响代码执行， 但是影响代码的阅读.



## 对象
### 对象的定义: 
#### 用代码操作或控制的东西即为对象, 如工作簿、 工作表、 单元格、 图片、 图表、 透视表等.


### 对象的属性: 
#### 定义: 属性是对象包含的内容或特点
#### 相对性: 对象和属性是相对的，如Sheet1工作表的A1单元格,A1单元格就是Sheet1工作表的属性；A1单元格的内容，内容就是A1单元格的属性.
#### 书写: 对象和属性之间用点连接，对象在前，属性在后，如A1单元格的内容，用汉字表达为：A1.内容 Range("A1").Value



### 对象的方法:
#### 定义: 方法是指在对象上执行的某个动作，如选中A1单元格，​“选中”是在A1单元格这个对象上执行的操作，就是A1单元格的方法。
#### 书写: 对象和方法之间也用点"."连接, 对象在前，方法在后，如选中A1单元格写成代码为：Range("A1").Select

### 所有对象的祖宗(究极对象): Application
#### 对象属性和方法的区分: 
##### 属性: 手指图标
##### 方法: 文件夹图标


## 常量和变量


### 常量
#### 定义: 在程序运行过程中其值不能改变的量
#### 书写: 用关键字Const定义，如Const pi = 3.1415926
#### 注意: 常量名一般使用大写字母，如Const PI = 3.1415926

### 变量
#### 定义: 在程序运行过程中其值可以改变的量
#### 书写: 用关键字Dim定义，如Dim i As Integer
#### 注意: 变量名一般使用小写字母，如Dim i As Integer

## 数据类型
### 整型: Integer
### 长整型 Long
### 单精度浮点型: Single
### 双精度浮点型: Double
### 字符串型: String
### 货币型: Currency
### 逻辑型: Boolean
### 日期型: Date
### 对象型: Object
### 变体型: Variant

### 注意: **小容器装大数据** 容易造成数据丢失，如把一个字符串赋值给一个整型变量，整型变量会丢失字符串中的非数字字符，如把字符串"123abc"赋值给整型变量，整型变量只会得到数字123，字符串中的"abc"会被丢弃.



##  运算符 
### 算术运算符
#### 加法运算符: +
#### 减法运算符: -
#### 乘法运算符: *
#### 除法运算符: /
#### 整除运算符: \
#### 求余运算符: Mod
#### 幂运算符: ^


### 比较运算符
#### 等于: =
#### 不等于: <>
#### 大于: >
#### 小于: <
#### 大于等于: >=
#### 小于等于: <=

### 逻辑运算符
#### And: 逻辑与
#### Or: 逻辑或
#### Not: 逻辑非




### 连接运算符
#### &: 连接字符串和数字
#### +: 连接字符串



### 运算符的优先级
#### 1. 括号
#### 2. 幂运算符
#### 3. 乘法、除法、整除、求余
#### 4. 加法、减法
#### 5. 比较运算符
#### 6. 逻辑运算符


## 流程控制语句
### 条件语句
#### If...Then...ElseIf...Else...End If


### 循环语句
#### For...Next
#### For Each...Next
#### Do While...Loop
#### Do...Loop Until
#### Do Until...Loop


### 跳转语句
#### Exit For
#### Exit Do
#### Exit Sub


## VBA内置函数  


### 数学函数  


#### ABS  
##### 返回数字的绝对值  
##### ABS(number)




#### INT  
##### 返回数字向下取整后的值  
##### INT(number)




#### ROUND  
##### 返回数字四舍五入后的值  
##### ROUND(number, num_digits)




#### ROUNDDOWN  
##### 返回数字向下取整后的值  
##### ROUNDDOWN(number, num_digits)




#### ROUNDUP  
##### 返回数字向上取整后的值  
##### ROUNDUP(number, num_digits)




#### TRUNC  
##### 返回数字截断后的值  
##### TRUNC(number, [num_digits])




#### SQRT  
##### 返回数字的平方根  
##### SQRT(number)




#### POWER  
##### 返回一个数的指定次幂  
##### POWER(number, power)




#### RAND  
##### 返回一个介于0和1之间的随机数  
##### RAND()




#### RANDBETWEEN  
##### 返回一个介于指定两个数之间的随机整数  
##### RANDBETWEEN(bottom, top)




### 字符串函数  


#### LEFT  
##### 返回字符串最左边的字符  
##### LEFT(text, [num_chars])




#### RIGHT  
##### 返回字符串最右边的字符  
##### RIGHT(text, [num_chars])




#### MID  
##### 返回字符串中间的字符  
##### MID(text, start_num, num_chars)




#### LEN  
##### 返回字符串的长度  
##### LEN(text)




#### LOWER  
##### 将字符串转换为小写  
##### LOWER(text)




#### UPPER  
##### 将字符串转换为大写  
##### UPPER(text)




#### REPLACE  
##### 替换字符串中的字符  
##### REPLACE(old_text, start_num, num_chars, new_text)




#### SUBSTITUTE  
##### 替换字符串中的字符  
##### SUBSTITUTE(text, old_text, new_text, [instance_num])




#### CONCATENATE  
##### 连接字符串  
##### CONCATENATE(text1, text2, ...)




#### CONCAT  
##### 连接字符串  
##### CONCAT(text1, text2, ...)




#### TRIM  
##### 删除字符串两端的空格  
##### TRIM(text)




#### REPT  
##### 重复字符串  
##### REPT(text, number_times)




#### FIND  
##### 查找字符串的位置  
##### FIND(find_text, within_text, [start_num])




#### SEARCH  
##### 查找字符串的位置  
##### SEARCH(find_text, within_text, [start_num])




#### EXACT  
##### 比较两个字符串是否完全相同  
##### EXACT(text1, text2)




### 类型判断函数  


#### ISNUMBER  
##### 判断一个值是否为数字  
##### ISNUMBER(value)




#### ISBLANK  
##### 判断一个值是否为空  
##### ISBLANK(value)




#### ISERROR  
##### 判断一个值是否为错误  
##### ISERROR(value)




#### ISLOGICAL  
##### 判断一个值是否为逻辑值  
##### ISLOGICAL(value)




#### ISTEXT  
##### 判断一个值是否为文本  
##### ISTEXT(value)




#### ISNONTEXT  
##### 判断一个值是否不是文本  
##### ISNONTEXT(value)


#### ISNONTEXT  
##### 判断一个值是否为空  
##### ISEMPTY(value)








### 时间日期函数  


#### NOW  
##### 返回当前的日期和时间  
##### NOW()




#### TODAY  
##### 返回当前的日期  
##### TODAY()




#### DATE  
##### 返回当前的日期  
##### DATE()




#### TIME  
##### 返回当前的时间  
##### TIME()




#### YEAR  
##### 返回日期中的年份  
##### YEAR(date)




#### MONTH  
##### 返回日期中的月份  
##### MONTH(date)




#### DAY  
##### 返回日期中的天数  
##### DAY(date)




#### HOUR  
##### 返回时间中的小时  
##### HOUR(time)




#### MINUTE  
##### 返回时间中的分钟  
##### MINUTE(time)




#### SECOND  
##### 返回时间中的秒  
##### SECOND(time)




#### DATEADD  
##### 在指定日期上添加一个时间间隔  
##### DATEADD(interval, number, date)




#### DATEDIFF  
##### 返回两个日期之间的时间间隔  
##### DATEDIFF(interval, date1, date2)




#### DATEPART  
##### 返回指定日期的指定部分  
##### DATEPART(interval, date)




#### FORMAT  
##### 格式化日期或时间  
##### FORMAT(expression, format)




#### WEEKDAY  
##### 返回日期是星期几  
##### WEEKDAY(date, [firstdayofweek])




#### EOMONTH  
##### 返回指定日期所在月份的最后一天  
##### EOMONTH(start_date, months)




#### WORKDAY  
##### 返回指定日期之后的工作日  
##### WORKDAY(start_date, days, [holidays])




#### NETWORKDAYS  
##### 返回两个日期之间的工作日数量  
##### NETWORKDAYS(start_date, end_date, [holidays])