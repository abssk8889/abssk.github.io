<!--#include file="../ziliang_inc/inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>栏目管理</title>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg"><strong>&nbsp;栏目管理</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font></td>
  </tr>
</table>
<table cellpadding="0" cellspacing="0" width="100%" border="1" bordercolordark="#FFFFFF" bordercolorlight="#CCCCCC">
<tr>
<td width="5%" align="center" valign="middle" bgcolor="#D9EFFE"><strong>排序</strong></td>
<td width="30%" height="30" align="left" valign="middle" bgcolor="#D9EFFE" style="padding:2px;"><strong>分类</strong></td>
<td  align="left" valign="middle" bgcolor="#D9EFFE" style="padding:2px;"><strong>操作</strong></td>
<td  align="left" valign="middle" bgcolor="#D9EFFE" style="padding:2px;"><strong>状态</strong></td>
</tr>
<%FenLei()%>
</table>
</body>
</html>
