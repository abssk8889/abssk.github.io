<!--#include file="../ziliang_inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>友情链接</title>
</head>

<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg"><strong>&nbsp;链接添加</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font></td>
  </tr>
</table>

<%If request.querystring("Edit")="Y_E" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_Links where clkj_yqId="&request("clkj_yqId")
	Rs.open Sql,conn,1,1
%>
<table width="100%" cellpadding="0" cellspacing="0">
   <tr>
    <td height="25" colspan="2" bgcolor="#F5F5F5" style="text-indent:2em;color:#FF0000;">修改</td>
  </tr>
<form name="links" action="?Class=links_Edit" method="post">
  <tr>
    <td width="15%" height="30" align="center" valign="middle">链接名</td>
    <td class="h_td"><input name="clkj_name" type="text" id="clkj_name" value="<%=Rs("clkj_title")%>" size="20" />
      如：Clkj</td>
  </tr>
    <tr>
    <td width="15%" height="30" align="center" valign="middle">链接地址</td>
    <td class="h_td"><input name="clkj_links" type="text" id="clkj_links" value="<%=Rs("clkj_LinkURL")%>" size="30" />
      如：http://www.zjxyk.com/</td>
  </tr>
        <tr>
    <td width="15%" height="30" align="center" valign="middle">描述</td>
    <td class="h_td"><input name="zl_l_miaosu" type="text" id="zl_l_miaosu" value="<%=Rs("zl_l_ms")%>" size="40" /></td>
  </tr>
      <tr>
    <td width="15%" height="30" align="center" valign="middle">排序</td>
    <td class="h_td"><input name="zl_l_paixu" type="text" id="zl_l_paixu" value="<%=Rs("zl_l_paixu")%>" size="5" />
      输入,0,1,2,3,4,5按从小到大排序。默认为0。注：只能输入数字</td>
  </tr>
      <tr>
    <td height="30" width="15%"></td>
    <td><input type="submit" name="Submit" value="修改" /><input type="hidden" value=<%=request.querystring("clkj_yqId")%> name="clkj_yqId"></td>
  </tr>
</form>  
</table>
<%Else%>
<table width="100%" cellpadding="0" cellspacing="0">
   <tr>
    <td height="25" colspan="2" bgcolor="#F5F5F5" style="text-indent:2em;color:#FF0000;">添加</td>
  </tr>
<form name="links" action="?Class=links" method="post">
  <tr>
    <td width="15%" height="30" align="center" valign="middle">链接名</td>
    <td class="h_td"><input name="clkj_name" type="text" id="clkj_name" size="20" />
      如：Clkj</td>
  </tr>
    <tr>
    <td width="15%" height="30" align="center" valign="middle">链接地址</td>
    <td class="h_td"><input name="clkj_links" type="text" id="clkj_links" size="30" />
      如：http://www.zjxyk.com/</td>
  </tr>
            <tr>
    <td width="15%" height="30" align="center" valign="middle">描述</td>
    <td class="h_td"><input name="zl_l_miaosu" type="text" id="zl_l_miaosu" size="40" />
    输入关键词</td>
  </tr>
        <tr>
    <td width="15%" height="30" align="center" valign="middle">排序</td>
    <td class="h_td"><input name="zl_l_paixu" type="text" id="zl_l_paixu" value="0" size="5" />
      输入,0,1,2,3,4,5按从小到大排序。默认为0。注：只能输入数字</td>
  </tr>

      <tr>
    <td height="30" width="15%"></td>
    <td><input type="submit" name="Submit" value="增加" /></td>
  </tr>
</form>  
</table>
<%End If%>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
  <!--DWLayoutTable-->
  <tr>
    <td width="100%" height="30" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <!--DWLayoutTable-->
      <tr>
        <td width="100%" height="39" align="left" valign="middle" class="tr_bg">&nbsp;<strong>链接管理</strong></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="25" valign="top">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <!--DWLayoutTable-->
      <tr>
        <td width="50" height="25" align="center" valign="middle" bgcolor="#F5F5F5"><strong>编号</strong></td>
		<td width="50" height="25" align="center" valign="middle" bgcolor="#F5F5F5">排序</td>
        <td width="180" align="left" valign="middle" bgcolor="#F5F5F5">网名</td>
        <td width="275" align="left" valign="middle" bgcolor="#F5F5F5">网址</td>
        <td width="100" align="center" valign="middle" bgcolor="#F5F5F5">操作</td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20" valign="top"><table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #A5B9C9;">
	
	<%
	sql="select * from zl_Links order by zl_l_paixu ,clkj_yqId asc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
IF Not rs.eof Then
proCount=rs.recordcount
	rs.PageSize=10			  '定义每页显示数目
		if not IsEmpty(Request("ToPage")) then
	       ToPage=CInt(Request("ToPage"))
		   if ToPage>rs.PageCount then
					rs.AbsolutePage=rs.PageCount
					intCurPage=rs.PageCount
		   elseif ToPage<=0 then
					rs.AbsolutePage=1
					intCurPage=1
				else
					rs.AbsolutePage=ToPage
					intCurPage=ToPage
				end if
			else
			        rs.AbsolutePage=1
					intCurPage=1
		  end if
	intCurPage=CInt(intCurPage)
	For i = 1 to rs.PageSize
	if rs.EOF then     
	Exit For 
	end if
%> 

       <tr>
        <td width="50" height="25" align="center" valign="middle" bgcolor="#F5F5F5" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("clkj_yqId")%></td>
		 <td width="50" height="25" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("zl_l_paixu")%></td>
        <td width="180" align="left" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("clkj_title")%></td>
        <td width="275" align="left" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("clkj_LinkURL")%></td>
        <td width="100" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><a href=ziliang_links.asp?clkj_yqId=<%=rs("clkj_yqId")%>&Edit=Y_E>修改</a>|<a href=?clkj_yqId=<%=rs("clkj_yqId")%>&Class=links_Del onClick="return confirm('是否将此删除?');">删除</a></td>
        </tr>
		<%
 rs.MoveNext
 next
 end if
 %><tr><td height=25 colspan="5" align="right" valign="middle" bgcolor="#F5F5F5"><div><span style="font-size: 9pt;"> 共&nbsp;<font color="#ff0000"><%=proCount%></font>&nbsp;条, 当前第：<font color="#ff0000"><%=intCurPage%></font>/<font color="#ff0000"><%=rs.PageCount%></font>页
                <% if intCurPage<>1 then %>
                <a href="ziliang_links.asp?cla=<%=nb%>&amp;ToPage=1" class="redlink">首页</a>&nbsp;|&nbsp;<a href="ziliang_links.asp?cla=<%=nb%>&amp;ToPage=<%=intCurPage-1%>" class="redlink">上一页</a>&nbsp;|
                <%else%>
                <font color="#999999">首页&nbsp;|&nbsp;上一页&nbsp;|
                <% end if%>
                <%if intCurPage<>rs.PageCount then %>
                <a href="ziliang_links.asp?cla=<%=nb%>&amp;ToPage=<%=intCurPage+1%>" class="redlink">下一页</a>&nbsp;| <a href="ziliang_links.asp?cla=<%=nb%>&amp;ToPage=<%=rs.PageCount%>" class="redlink"> 末页</a>&nbsp;|&nbsp;
                <%else%>
                <font color="#999999">下一页&nbsp;|&nbsp;末页&nbsp;|&nbsp; </font>
                <% end if%></div></td></tr>
    </table></td>
  </tr>
</table>
</body>
</html>