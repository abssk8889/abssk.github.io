<!--#include file="../ziliang_inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<LINK href="backdoor.css" type=text/css rel=stylesheet>
</head>
<body bgcolor="#CCCCCC">
<table width="100%" height="91" border="0" cellpadding="0" cellspacing="0" background="../ziliang_images/back_image/back_images_top.gif" >
  <!--DWLayoutTable-->
<tr>
<td width="255" rowspan="2" align="center" valign="middle"><img src="../ZiLiang_Images/aliang.gif" alt="�������͹���ϵͳ"></td>
<td width="996" height="52">&nbsp;<font color="#FF9900"><b><%=session("username")%></b></font>&nbsp;[<a href="?class=Out" target="_parent">��ȫ�˳�</a>��<a href="ziliang_user.asp" target="mainFrame">�û�����</a>] <font color="#FF9900">������:<%=year(now())%>��<%=month(now())%>��<%=day(now())%>�� ����<% if WeekDay (date())-1 =0 then %>��<% else if WeekDay (date())-1 =1 then %>һ<% else if WeekDay (date())-1 =2 then %>��<% else if WeekDay (date())-1 =3 then %>��<% else if WeekDay (date())-1 =4 then %>��<% else if WeekDay (date())-1 =5 then %>��<% else if WeekDay (date())-1 =6 then %>��<% end if%><% end if%><% end if%><% end if%><% end if%><% end if%><% end if%></font></td>  
</tr>
<tr>
  <td height="39" align="left" valign="middle" ><table width="700" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr>
    <td height="39" align="center" valign="middle" background="../ziliang_images/back_image/back_top_1.png"><a href="html/ziliang_index.asp" target="mainFrame"><font color="#FFFFFF">������ҳ</font></a></td>
    <td width="10" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td width="90" align="center" valign="middle" background="../ziliang_images/back_image/back_top_1.png"><a href="ziliang_class.asp" target="mainFrame"><font color="#FFFFFF">ҳ������</font></a></td>
	    <td width="10" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td width="90" align="center" valign="middle" background="../ziliang_images/back_image/back_top_1.png"><a href="http://aliang.zjxyk.com/" target="mainFrame"><font color="#FFFFFF">��������</font></a></td>
    <td width="410"></td>
  </tr>
</table>
</td>
  </tr>
</table>

</body>
</html>
