<!--#include file="../ziliang_inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>�û�����</title>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg"><strong>&nbsp;�û�����</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font></td>
  </tr>

</table>
<% If request.querystring("Edit")="B_E" Then 
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_admin where zl_id ="&request("zl_id")
	Rs.open Sql,conn,1,1
%>
<table width="100%" cellpadding="0" cellspacing="0">
    <tr>
    <td height="25" bgcolor="#F5F5F5" style="text-indent:2em;color:#FF0000;" id="zx" colspan="2">�޸��û�</td>
  </tr>
 <form name="pas" action="?Class=pas_Edit" method="post">
  <tr>
    <td height="30" align="center" valign="middle">�û�Ȩ��</td>
    <td><select name="zl_flag" id="zl_flag">
      <option value="1" <%if rs("zl_flag")=1 then%>selected<%end if%>>����Ա</option>
      <option value="0" <%if rs("zl_flag")=0 then%>selected<%end if%>>��ͨԱ</option>
    </select>
    </td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle">�û���</td>
    <td><input name="user_name" type="text" id="user_name" value="<%=rs("zl_user")%>" size="20"></td>
  </tr>
    <tr>
    <td width="15%" height="30" align="center" valign="middle">����</td>
    <td><input name="user_pas" type="text" id="user_pas"  size="20">
    ����Ϊ�ղ��޸�</td>
  </tr>
    <tr>
    <td width="15%" height="30"></td>
    <td><input type="submit" name="Submit" value="�޸�"><input type="hidden" name="zl_id" value="<%=request("zl_id")%>" /></td>
  </tr>
  </form>
</table>
<%Else%>
<table width="100%" cellpadding="0" cellspacing="0">
    <tr>
    <td height="25" bgcolor="#F5F5F5" style="text-indent:2em;color:#FF0000;" id="zj" colspan="2">����û�</td>
  </tr>
 <form name="pas" action="?Class=pas" method="post">
  <tr>
    <td height="30" align="center" valign="middle">�û�Ȩ��</td>
    <td><select name="zl_flag">
      <option value="1" selected>����Ա</option>
      <option value="0">��ͨԱ</option>
    </select>
    </td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle">�û���</td>
    <td><input name="user_name" type="text" id="user_name" size="20"></td>
  </tr>
    <tr>
    <td width="15%" height="30" align="center" valign="middle">����</td>
    <td><input name="user_pas" type="text" id="user_pas" size="20"></td>
  </tr>
    <tr>
    <td width="15%" height="30"></td>
    <td><input type="submit" name="Submit" value="����"></td>
  </tr>
  </form>
</table>
<%End if%>
<br>
<table width="100%" cellpadding="0" cellspacing="0" >
<tr>
<td height="39" colspan="5" class="tr_bg" style="color:#FF0000;">&nbsp;�û��б�&nbsp;[<a href="ziliang_user.asp#zj">�����û�</a>]</td>
</tr>
<%call user()%>
</table>
</body>
</html>