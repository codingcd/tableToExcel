# tableToExcel
tableToExcel是一个带样式支持多sheet的前端表格导出Exel的插件，我们只需要简单的引入插件，就能用我们熟悉的html标签定义我们导出表格的样式



## 快速入门
**1.下载插件**

**2.在html文件中引入插件**
```js
<script src="./tableToExel.js" type="text/javascript"></script>  // 路径是根据自己的项目引入就行
```

**3.使用插件**
```
// 第一种用法:传入DOM对象
var tab1 = document.getElementById("tab1"); 
var tab2 = document.getElementById("tab2");
var tab3 = document.getElementById("tab3");
tableToExcel([tab1,tab2,tab3], ['测试1','测试2','测试3']);  // 传入两个数组参数，第一个是选择的DOM对象组成的数组，第二个是定义导出后表格对应的sheet名称

// 第二种用法：传入拼接字符串
var tab1 = "<table><thead><tr><th>"+测试1+"</th></tr></thead>"+
  "<tbody><tr><td>"+222+"</td></tr></tbody></table>";
var tab2 = "<table><thead><tr><th>"+测试2+"</th></tr></thead>"+
  "<tbody><tr><td>"+333+"</td></tr></tbody></table>";
var tab3 = `<table>
    <thead><tr><th>测试3</th></tr></thead>
    <tbody><tr><td></td></tr></tbody>
  </table>`;
tableToExcel([tab1,tab2,tab3], ['测试1','测试2','测试3']);

```

**4.功能特性**
- 支持多sheet
- 支持大部分html语法定义的表格样式
- 支持自定义表格名称，前提是你浏览器设置了表格下载弹出询问框
- 暂不支持图片导出

**5.兼容性**

兼容除ie之外的几乎所有现代流行浏览器
