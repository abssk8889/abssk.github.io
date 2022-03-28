<!--#include file="CE.asp"-->
<!--#include file="c.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=request.form("search")%>_阿梁网络实验室</title>
<style type="text/css">
<!--
body,td,th {
	font-family: 宋体;
	font-size: 13px;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
a {
	font-family: 宋体;
	font-size: 14px;
}
.cc{clear:both;height:0;font-size:1px;line-height:0px;}
-->
</style>
</head>
<body>
<div style="margin:auto; height:auto; width:100%;">
<div class="cc"></div>
<div style="margin:auto; width:100%;">
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <form action="/ZiLiang_inc/search.asp" name="search" method="post">
      <tr>
        <td width="126" align="center" valign="middle"><img src="logo.gif" width="119" height="47" /></td>
        <td width="331" height="60" align="center" valign="middle">&nbsp;
          <input name="search" type="text" value="请输入关键词:阿梁blog" size="40" onmouseover="this.select()" style="color:#666666;" />
          &nbsp;</td>
        <td width="794" align="left" valign="middle">&nbsp;
        <input type="submit" name="Submit" value="搜索"></td>
      </tr>
    </form>
  </table>
</div>
<div class="cc"></div>
<div style="background:#B8ECEF; height:25px; line-height:25px; text-indent:1em; margin:auto;"> <a href="http://aliang.zjxyk.com/">阿梁网络实验</a> 您目前所搜索的关链词：<font color=#ffffff><%=request.form("search")%></font> <a href=/>返回首页</a></div>
<div style="width:70%; height:100%; margin-left:10px; margin-top:10px;float:left;">
  <%
dim gjz
gjz=trim(request.form("search"))
set rs=server.createobject("adodb.recordset")
exec="select * from zl_content where zl_news_title like'%"&gjz&"%' or zl_news_key like'%"&gjz&"%' or zl_news_maiosu like'%"&gjz&"%' or zl_news_lei_name like '%"&gjz&"%' and zl_news_lei_id=2"
rs.open exec,conn,1,1

	IF rs.eof or gjz="" then
		response.write"暂无记录：单击<a href="&chr(34)&"javascript:history.back(1)"&chr(34)&">上一步</a>按钮，尝试其关键词搜索。"
	Else
		Do while not rs.eof
		search_y = year(rs("zl_news_time"))
		search_id = rs("zl_news_id")
		search_wjm = search_y+search_id
		search_lm = c(rs("zl_news_lei_name"))
		search_bt = rs("zl_news_title")
		search_nr = rs("zl_news_maiosu")
		t = formatdatetime(rs("zl_news_time"),2)
		response.write("<a href="&chr(34)&"/"&search_lm&"/"&search_wjm&".html"&chr(34)&" target="&chr(34)&"_blank"&chr(34)&">")
		response.Write replace(search_bt,""&gjz&"","<font color=#ff0000>"&gjz&"</font>")
		response.write("</a>")
		response.Write("<br>") 
		response.write search_nr
		response.Write("<br>")
		response.write("<font color='#009933'>http://aliang.zjxyk.com/"&search_lm&"/"&search_wjm&".html ")
		response.write t&" </font>"
		response.write(" <a href="&chr(34)&"/"&search_lm&"/"&search_wjm&".html"&chr(34)&" target="&chr(34)&"_blank"&chr(34)&">")
		response.Write "<font color=#666666>详细页面</font>"
		response.write("</a>")
		response.Write("<br><br>")
		Rs.movenext
		Loop
	End IF
Rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</div>
<div class="cc"></div>
<div style="background:#9AE4E8; margin-top:30px;float:left; width:100%;">
  <table cellpadding="0"  cellspacing="0" width="100%">
    <tr>
      <td height="30" align="center" valign="middle"><div style="display:none;"><script language="javascript" type="text/javascript" src="http://js.users.51.la/3519468.js"></script>
        <noscript>
        <a href="http://www.51.la/?3519468" target="_blank"><img alt="我要啦免费统计" src="http://img.users.51.la/3519468.asp" style="border:none" /></a>
        </noscript></div>
        <font color="#FFFFFF">@2010 阿梁seo,阿梁英文网站建设,阿梁网站建设　版权所有 阿梁网络实验室  QQ:41864438</font></td>
    </tr>
    <tr>
      <td height="10" bgcolor="#ECECEC"></td>
    </tr>
  </table>
</div>
</div>
</body>
</html>
