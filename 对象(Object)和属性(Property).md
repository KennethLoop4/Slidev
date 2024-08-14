---
theme: apple-basic
background: https://cover.sli.dev
layout: intro
class: text-center
highlighter: shiki
transition: view-transition
--- 

# 对象 & 对象的属性和方法

Power by Will Ran

<div class="pt-12">
  <span @click="next" class="px-2 p-1 rounded cursor-pointer hover:bg-white hover:bg-opacity-10">
    Learn more <carbon:arrow-right class="inline"/>
  </span>
</div>

---
transition: view-transition
---

# 对象

<div v-click=1>


用代码操作和控制的东西即为对象，如<span v-mark.circle.red="2">工作簿</span>、<span v-mark.circle.red="3">工作表</span>、<span v-mark.circle.red="4">单元格</span>、图片、图表、透视表等。




</div>



<div v-click=1>

```vb {None|None|1-3|5-7|9-11|all} twoslash
'工作簿  
Dim wb As Workbook  
Set wb = Workbooks.Open("C:\path\to\your\file.xlsx")

'工作表
Dim ws As Worksheet  
Set ws = ThisWorkbook.Sheets("Sheet1")

'单元格
Dim rng As Range  
Set rng = ws.Range("A1")
```

</div>

---
transition: view-transition
---


# 对象的属性

<br>


<div v-click=1>

每个对象都有属性，属性是对象包含<mark>内容或特点</mark>。

如<span v-mark.circle.red="2">Sheet1工作表</span>的<span v-mark.circle.red="3">A1单元格</span>，A1单元格就是Sheet1工作表的属性；<span v-mark.circle.red="4">A1单元格</span>的<span v-mark.circle.red="5">内容</span>，内容就是A1单元格的属性。

</div>


<div v-click=6>

在书写时，对象和属性之间用点（<span v-mark.circle.red="7">.</span>）连接，对象在前，属性在后，如A1单元格的内容，用汉字表达为：<span v-mark.orange="7">A1.内容</span>


</div>


<div v-click=8>

```vb
Range("A1").Value
```

</div>








---
transition: view-transition
---


# 对象的方法

<div v-click=1>


每个对象都有方法，方法是指在对象上执行的某个<mark>动作</mark>。如选中A1单元格，​“选中”是在A1单元格这个对象上执行的操作，就是A1单元格的方法。<span v-mark.red="2">对象和方法之间也用点（.）连接</span>，对象在前，方法在后，如选中A1单元格写成代码为：

</div>



<div v-click=3>

```vb
Range("A1").Select
```

</div>

<br>
<br>


<div v-click=4>

###### 属性和方法
```vb
Range("A1").Value
Range("A1").Select
```

</div>




