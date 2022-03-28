<!--#include file="../ziliang_inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<LINK href="backdoor.css" type=text/css rel=stylesheet>
</head>
<body bgcolor="#CCCCCC">
<table width="100%" height="91" border="0" cellpadding="0" cellspacing="0" background="../ziliang_images/back_image/back_images_top.gif" >
  <!--DWLayoutTable-->
<tr>
<td width="255" rowspan="2" align="center" valign="middle"><img src="../ZiLiang_Images/aliang.gif" alt="阿梁博客管理系统"></td>
<td width="996" height="52">&nbsp;<font color="#FF9900"><b><%=session("username")%></b></font>&nbsp;[<a href="?class=Out" target="_parent">安全退出</a>，<a href="ziliang_user.asp" target="mainFrame">用户管理</a>] <font color="#FF9900">今天是:<%=year(now())%>年<%=month(now())%>月<%=day(now())%>号 星期<% if WeekDay (date())-1 =0 then %>天<% else if WeekDay (date())-1 =1 then %>一<% else if WeekDay (date())-1 =2 then %>二<% else if WeekDay (date())-1 =3 then %>三<% else if WeekDay (date())-1 =4 then %>四<% else if WeekDay (date())-1 =5 then %>五<% else if WeekDay (date())-1 =6 then %>六<% end if%><% end if%><% end if%><% end if%><% end if%><% end if%><% end if%></font></td>  
</tr>
<tr>
  <td height="39" align="left" valign="middle" ><table width="700" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr>
    <td height="39" align="center" valign="middle" background="../ziliang_images/back_image/back_top_1.png"><a href="html/ziliang_index.asp" target="mainFrame"><font color="#FFFFFF">生成首页</font></a></td>
    <td width="10" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td width="90" align="center" valign="middle" background="../ziliang_images/back_image/back_top_1.png"><a href="ziliang_class.asp" target="mainFrame"><font color="#FFFFFF">页面生成</font></a></td>
	    <td width="10" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td width="90" align="center" valign="middle" background="../ziliang_images/back_image/back_top_1.png"><a href="http://aliang.zjxyk.com/" target="mainFrame"><font color="#FFFFFF">帮助中心</font></a></td>
    <td width="410"></td>
  </tr>
</table>
</td>
  </tr>
</table>

</body>
</html>
