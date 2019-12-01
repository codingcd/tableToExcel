/*
   导出表格到exel
    传递参数，table数据数组，sheet名称数组
    不兼容ie
*/
var tableToExcel = (function() {
    var uri = 'data:application/vnd.ms-excel;base64,'
    home_template = `MIME-Version: 1.0
X-Document-Type: Workbook
Content-Type: multipart/related; boundary="----BOUNDARY_9527----"

------BOUNDARY_9527----
Content-Location: file:///C:/0E8D990C/MimeExcel.xml
Content-Transfer-Encoding: quoted-printable
Content-Type: text/html; charset="us-ascii"

<html xmlns:o=3D"urn:schemas-microsoft-com:office:office"
xmlns:x=3D"urn:schemas-microsoft-com:office:excel"
xmlns=3D"http://www.w3.org/TR/REC-html40">

<head>
<xml>
 <o:DocumentProperties>
  <o:Author>xcb</o:Author>
  <o:LastAuthor>xcb</o:LastAuthor>
  <o:Created>2005-07-10T17:02:17Z</o:Created>
  <o:LastSaved>2005-07-10T17:06:05Z</o:LastSaved>
  <o:Company>u-soft</o:Company>
  <o:Version>11.5606</o:Version>
 </o:DocumentProperties>
</xml>
<xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
   {exlsheet}</x:ExcelWorksheets>
 </x:ExcelWorkbook>
</xml>
</head>
</html>`;
    home_child_template = `<x:ExcelWorksheet>
    <x:Name>{name}</x:Name>
    <x:WorksheetSource HRef=3D"cid:{sheet}"/>
</x:ExcelWorksheet>
`;
    sheet_page_template = `
------BOUNDARY_9527----
Content-ID: {sheetid}
Content-Transfer-Encoding: quoted-printable
Content-Type: text/html; charset="us-ascii"

<html xmlns:o=3D"urn:schemas-microsoft-com:office:office"
xmlns:x=3D"urn:schemas-microsoft-com:office:excel"
xmlns=3D"http://www.w3.org/TR/REC-html40">

<head>
<xml>
 <x:WorksheetOptions>
  <x:ProtectContents>False</x:ProtectContents>
  <x:ProtectObjects>False</x:ProtectObjects>
  <x:ProtectScenarios>False</x:ProtectScenarios>
 </x:WorksheetOptions>
 </xml>
</head>
<body>
{table}
</body>
</html>`;
    base64 = function(s) {
        return window.btoa(unescape(encodeURIComponent(s)))
    },
    // 下面这段函数作用是：将template中的变量替换为页面内容ctx获取到的值
    format = function(s, c) {
        return s.replace(/{(\w+)}/g,
            function(m, p) {
                return c[p];
        })
    };
    fdomToString = function(node) {  
             let tmpNode = document.createElement('div')
             tmpNode.appendChild(node) 
             let str = tmpNode.innerHTML
             tmpNode = node = null; // 解除引用，以便于垃圾回收  
             return str;  
        }
    // 设置表格样式，传入table字符串
    setStyle = function(tab){
        var str = tab.replace(/\=/g,'=3D');
        return str.replace(/px/g,'pt');
    }
    return function(table,name) {
        // 定义整体页面
        var page = ''
        // 先替换home_child_template
        var sheet_name_ctx = {name:'',sheet:''};
        var home_child = ''
        for(var i = 0;i<name.length;i++){
            sheet_name_ctx.name = name[i];
            sheet_name_ctx.sheet = 'sheet' + i;
            home_child += format(home_child_template,sheet_name_ctx);
        }
        // 再替换home_template
        var home_ctx = {exlsheet:home_child}
        var home = format(home_template,home_ctx);
        // 替换sheet页内容
        var sheet_ctx = {sheetid: '', table: ''};
        var sheet = ''
        for(var j = 0;j<table.length;j++){
            // 判断table是否为字符串，若不为字符串则先将其转换成字符串
            if(typeof table[j] != "string"){
                table[j] = fdomToString(table[j]);
            }
            sheet_ctx.sheetid = 'sheet' + j;
            sheet_ctx.table = setStyle(table[j]);
            sheet += format(sheet_page_template, sheet_ctx);
        }
        page = home + sheet;
        // format()函数：通过格式操作使任意类型的数据转换成一个字符串
        // base64()：进行编码
        window.location.href =uri+ base64(page);
    }
})()