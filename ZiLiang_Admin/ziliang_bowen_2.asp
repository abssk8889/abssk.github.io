<!--#include file="../ziliang_inc/inc.asp"-->
<!--#include file="ziliang_aTOh.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>信息添加</title>
</head>
<body>


<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg"><strong>&nbsp;信息添加</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font> </td>
  </tr>
</table>
<%
IF request.QueryString("Edit")="N_E" then
	Set Rs_b=server.createobject("ADODB.Recordset")
	Sql_b="select * from zl_content where zl_news_id="&cint(request("zl_news_id"))
	Rs_b.open Sql_b,conn,1,1
%>
<table width="100%" cellpadding="2" cellspacing="0">
<form name="contenttt" action="?Class=content_edit" method="post">
<tr><td height="25" colspan="2" bgcolor="#F5F5F5" style="text-indent:2em; color:#FF0000">信息修改</td></tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">选 择</td>
    <td><select name="lei">
        <option value="">选择</option>
        <%call small_input2()%>
      </select>
      排序
      <input name="paixu" type="text" id="paixu" value="<%=rs_b("zl_news_paixu")%>" size="5">
      如0,1,2,3,4,5以小到大排列      <font color="#FF0000">*</font></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">标 题</td>
    <td><input name="title" type="text" id="title" value="<%=rs_b("zl_news_title")%>" size="50">
      <font color="#FF0000">*</font></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">关键词</td>
    <td><input name="key" type="text" id="key" value="<%=rs_b("zl_news_key")%>" size="40">
      如果多个关键词以“,”分开，如asp,flash,jsp<font color="#FF0000">*</font></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">描 述</td>
    <td><textarea name="des" cols="50" rows="4" id="des"><%=rs_b("zl_news_maiosu")%></textarea>
      <font color="#FF0000">*</font></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">内 容</td>
    <td><input name="content" type="hidden" id="content" value="<%=Server.HtmlEncode(rs_b("zl_news_nr"))%>">
      <IFRAME ID="content" SRC="/ZiLiang_Edit/ewebeditor.asp?id=content&style=s_blue" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="360"></IFRAME></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">发布者</td>
    <td><input name="user" type="text" id="user" value="ZiLiang" size="15">
      时间
      <input name="time" type="text" id="time" value="<%=now()%>" size="20" readonly="readonly"></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">&nbsp;</td>
    <td><input type="submit" name="Submit" value="修改"><input type="hidden" name="zl_news_id" value="<%=rs_b("zl_news_id")%>"></td>
  </tr>
  </form>
</table>
<%Else%>
<table width="100%" cellpadding="2" cellspacing="0">
<tr><td height="25" colspan="2" bgcolor="#F5F5F5" style="text-indent:2em; color:#FF0000">信息添加</td>
</tr>
<form name="contentt" action="?Class=content" method="post">
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">选 择</td>
    <td><select name="lei">
        <option value="">选择</option>
        <%call small_input2()%>
      </select>
      排序
      <input name="paixu" type="text" id="paixu" value="0" size="5">
      如0,1,2,3,4,5以小到大排列      <font color="#FF0000">*</font></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">标 题</td>
    <td><input name="title" type="text" id="title" size="50">
      <font color="#FF0000">*</font></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">关键词</td>
    <td><input name="key" type="text" id="key" size="40">
      如果多个关键词
      以“,”分开，如asp,flash,jsp<font color="#FF0000">*</font></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">描 述</td>
    <td><textarea name="des" cols="50" rows="4" id="des"></textarea>
      <font color="#FF0000">*</font></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">内 容</td>
    <td><input name="content" type="hidden" id="content">
      <IFRAME ID="content" SRC="/ZiLiang_Edit/ewebeditor.asp?id=content&style=s_blue" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="360"></IFRAME></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">发布者</td>
    <td><input name="user" type="text" id="user" value="ZiLiang" size="15">
      时间
      <input name="time" type="text" id="time" value="<%=now()%>" size="20" readonly="readonly"></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F5F5F5">&nbsp;</td>
    <td><input type="submit" name="Submit" value="增加"></td>
  </tr>
  </form>
</table>
<%End IF%>
</body>
</html>
