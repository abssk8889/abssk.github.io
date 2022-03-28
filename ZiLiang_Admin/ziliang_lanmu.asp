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
<%IF request.querystring("add")="add" then%>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<form action="?Class=Ladd" method="post" name="LeiBie">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F3F3F3">&nbsp;<font color="#ff0000">二级栏目增加</font></td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">一级栏目</td>
<td><select name="zl_big">
  <option value="">选择</option>
  <%call Big_input()%>
</select>
  *必选</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">二级栏目</td>
<td><input name="zl_s_name" type="text" id="zl_s_name" />
  *必填</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">二级栏目标题</td>
<td><input name="zl_s_bt" type="text" id="zl_s_bt" size="50" />
如，flash教程|flash描绘|</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">二级栏目关键词</td>
<td><input name="zl_s_key" type="text" id="zl_s_key" size="50" />
  输入2~3个。如：asp,java,php</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">二级栏目描述</td>
<td><textarea name="zl_s_des" cols="40" rows="5" id="zl_s_des"></textarea>
  控制100字符内。</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">排序</td>
<td><input name="zl_paixu" type="text" id="zl_paixu" value="0" size="5" />
  输入,0,1,2,3,4,5按从小到大排序。默认为0。注：只能输入数字</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">链接文件</td>
<td><input name="zl_adress" type="text" id="zl_adress" size="40" />
  链接模版文件地址，必要时需改。</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">是否开启</td>
<td><input name="flag" type="checkbox" id="flag" value="yes" />  
  是否开启</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">&nbsp;</td>
<td><input type="submit" name="Submit" value="增加二级栏目" /></td>
</tr>
</form>
</table>
<%else%>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<form action="?Class=Led" method="post" name="LeiBie2">
<tr>
<td height="30" colspan="2" align="left" valign="middle" bgcolor="#F3F3F3">&nbsp;<font color="#ff0000">二级栏目修改</font></td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">一级栏目</td>
<td><select name="zl_big">
  <option value="">选择</option>
  <%call Big_input()%>
</select>
  *必选</td>
</tr>
<%
Set Rs=server.createobject("ADODB.Recordset")
Sql="select * from zl_small_class where zl_s_id="&cint(request("zl_s_id"))
Rs.open Sql,conn,1,1
%>
<tr>
<td width="15%" height="30" align="center" valign="middle">二级栏目</td>
<td><input name="zl_s_name" type="text" id="zl_s_name" value="<%=rs("zl_s_name")%>" />
  *必填</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">二级栏目标题</td>
<td><input name="zl_s_bt" type="text" id="zl_s_bt" value="<%=rs("zl_s_bt")%>" size="50" />
如，flash教程|flash描绘|</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">二级栏目关键词</td>
<td><input name="zl_s_key" type="text" id="zl_s_key" value="<%=rs("zl_s_key")%>" size="50" />
  输入2~3个。如：asp,java,php</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">二级栏目描述</td>
<td><textarea name="zl_s_des" cols="40" rows="5" id="zl_s_des"><%=rs("zl_s_des")%></textarea>
  控制100字符内。</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">排序</td>
<td><input name="zl_paixu" type="text" id="zl_paixu" value="<%=rs("zl_paixu")%>" size="5" />
  输入,0,1,2,3,4,5按从小到大排序。默认为0。注：只能输入数字</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">链接文件</td>
<td><input name="zl_adress" type="text" id="zl_adress" value="<%=rs("zl_adress")%>" size="40" />
链接模版文件地址，必要时需改。</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">是否开启</td>
<td><input name="flag" type="checkbox" id="flag" value="yes" <% if rs("flag")=1 then response.Write("checked") end if%> />  
  是否开启</td>
</tr>
<tr>
<td width="15%" height="30" align="center" valign="middle">&nbsp;</td>
<td><input type="submit" name="Submit" value="修改二级栏目" /><input type="hidden" name="zl_s_id" value="<%=rs("zl_s_id")%>"></td>
</tr>
</form>
</table>
<%end if%>
</body>
</html>
