<!--#include file="../ziliang_inc/inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>��Ŀ����</title>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg"><strong>&nbsp;��Ŀ����</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font></td>
  </tr>
</table>
<%IF request.querystring("add")="add" then%>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<form action="?Class=Ladd" method="post" name="LeiBie">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F3F3F3">&nbsp;<font color="#ff0000">������Ŀ����</font></td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">һ����Ŀ</td>
<td><select name="zl_big">
  <option value="">ѡ��</option>
  <%call Big_input()%>
</select>
  *��ѡ</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">������Ŀ</td>
<td><input name="zl_s_name" type="text" id="zl_s_name" />
  *����</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">������Ŀ����</td>
<td><input name="zl_s_bt" type="text" id="zl_s_bt" size="50" />
�磬flash�̳�|flash���|</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">������Ŀ�ؼ���</td>
<td><input name="zl_s_key" type="text" id="zl_s_key" size="50" />
  ����2~3�����磺asp,java,php</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">������Ŀ����</td>
<td><textarea name="zl_s_des" cols="40" rows="5" id="zl_s_des"></textarea>
  ����100�ַ��ڡ�</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">����</td>
<td><input name="zl_paixu" type="text" id="zl_paixu" value="0" size="5" />
  ����,0,1,2,3,4,5����С��������Ĭ��Ϊ0��ע��ֻ����������</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">�����ļ�</td>
<td><input name="zl_adress" type="text" id="zl_adress" size="40" />
  ����ģ���ļ���ַ����Ҫʱ��ġ�</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">�Ƿ���</td>
<td><input name="flag" type="checkbox" id="flag" value="yes" />  
  �Ƿ���</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">&nbsp;</td>
<td><input type="submit" name="Submit" value="���Ӷ�����Ŀ" /></td>
</tr>
</form>
</table>
<%else%>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<form action="?Class=Led" method="post" name="LeiBie2">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F3F3F3">&nbsp;<font color="#ff0000">������Ŀ�޸�</font></td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">һ����Ŀ</td>
<td><select name="zl_big">
  <option value="">ѡ��</option>
  <%call Big_input()%>
</select>
  *��ѡ</td>
</tr>
<%
Set Rs=server.createobject("ADODB.Recordset")
Sql="select * from zl_small_class where zl_s_id="&cint(request("zl_s_id"))
Rs.open Sql,conn,1,1
%>
<tr>
<td width="15%" height="30" align="center" valign="middle">������Ŀ</td>
<td><input name="zl_s_name" type="text" id="zl_s_name" value="<%=rs("zl_s_name")%>" />
  *����</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">������Ŀ����</td>
<td><input name="zl_s_bt" type="text" id="zl_s_bt" value="<%=rs("zl_s_bt")%>" size="50" />
�磬flash�̳�|flash���|</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">������Ŀ�ؼ���</td>
<td><input name="zl_s_key" type="text" id="zl_s_key" value="<%=rs("zl_s_key")%>" size="50" />
  ����2~3�����磺asp,java,php</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">������Ŀ����</td>
<td><textarea name="zl_s_des" cols="40" rows="5" id="zl_s_des"><%=rs("zl_s_des")%></textarea>
  ����100�ַ��ڡ�</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">����</td>
<td><input name="zl_paixu" type="text" id="zl_paixu" value="<%=rs("zl_paixu")%>" size="5" />
  ����,0,1,2,3,4,5����С��������Ĭ��Ϊ0��ע��ֻ����������</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">�����ļ�</td>
<td><input name="zl_adress" type="text" id="zl_adress" value="<%=rs("zl_adress")%>" size="40" />
����ģ���ļ���ַ����Ҫʱ��ġ�</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">�Ƿ���</td>
<td><input name="flag" type="checkbox" id="flag" value="yes" <% if rs("flag")=1 then response.Write("checked") end if%> />  
  �Ƿ���</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">&nbsp;</td>
<td><input type="submit" name="Submit" value="�޸Ķ�����Ŀ" /><input type="hidden" name="zl_s_id" value="<%=rs("zl_s_id")%>"></td>
</tr>
</form>
</table>
<%end if%>
</body>
</html>
