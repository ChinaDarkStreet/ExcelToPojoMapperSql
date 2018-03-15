<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>pojo, mapper, createTableSQL语句文件生成器</title>
</head>
<body>
    <h1>文件必须是xlsx格式文件(即必须是2007以上版本的office文档)</h1>
    <hr>
    <h2>上传你的文档之后, 提交, 几秒钟后会自动下载, 请耐心等待 ....</h2>
    <ul style="margin-left: 20px">
        <li>上传文件大小 < 1MB</li>
        <li>文本内部使用 "\n" 换行符, windows记事本 "\r\n" 为换行, 所以会导致排版错误, 编辑器和linux中可完美显示</li>
        <li>最终结果为zip压缩文件, 暂时无用户备份功能</li>
        <li>请不要使用多线程(IDM等)下载器下载, 会导致出错!</li>
        <li>请不要恶意攻击此服务器, 否则会列入黑名单, 无法使用服务</li>
    </ul>
    <br>
    <form action="/excelToPojo" method="post" enctype="multipart/form-data">
        <input type="file" name="file">
        <input type="submit">
    </form>
    <br>
    <img style="width: 48%" src="image/testGeneratePojo.gif">
</body>
</html>