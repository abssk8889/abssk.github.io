<!--#include file="../ziliang_inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>��Ϣ���</title>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg"><strong>&nbsp;��Ϣ����</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font> [<a href="ziliang_bowen.asp">������Ϣ</a>]</td>
  </tr>
</table>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30" bgcolor="#EAEAEA">&nbsp;������Ŀ��
    <%call small_lei()%></td>
  </tr>
    <tr>
    <td height="30" bgcolor="#EAEAEA">&nbsp;������Ŀ��
    <%call small_lei2()%></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <!--DWLayoutTable-->
      <tr>
        <td width="50" height="25" align="center" valign="middle" bgcolor="#F5F5F5">���</td>
        <td width="80" align="left" valign="middle" bgcolor="#F5F5F5">���·���</td>
        <td width="260" align="left" valign="middle" bgcolor="#F5F5F5">���±���</td>
		<td width="40" align="center" valign="middle" bgcolor="#F5F5F5">����</td>
		<td width="40" align="center" valign="middle" bgcolor="#F5F5F5">δ��</td>
		<td width="40" align="center" valign="middle" bgcolor="#F5F5F5">���</td>
        <td width="80" align="center" valign="middle" bgcolor="#F5F5F5">¼��Ա</td>
        <td width="40" align="center" valign="middle" bgcolor="#F5F5F5">����</td>
        <td width="140" align="center" valign="middle" bgcolor="#F5F5F5">����</td>
  </tr>
    </table>
	</td>
  </tr>
  <tr>
    <td height="20" valign="top"><table width="100%" cellpadding="2" cellspacing="0" style="border-top:1px solid #A5B9C9;">
	
	<%
	di=cint(request("id"))
	if request("id")="" then
	sql="select * from zl_content order by zl_news_paixu asc,zl_news_id desc"
	else
	sql="select * from zl_content where zl_news_Lei="&di&" order by zl_news_paixu asc,zl_news_id desc"
	end if
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
IF Not rs.eof Then
proCount=rs.recordcount
	rs.PageSize=17			  '����ÿҳ��ʾ��Ŀ
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
        <td width="50" height="25" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("zl_news_id")%></td>
        <td width="80" align="left" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;">
		<%call back(rs("zl_news_Lei"))%></td>

        <td width="260" align="left" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("zl_news_title")%></td>
				 <td width="40" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><a href="ziliang_pr_hf.asp?zl_news_id=<%=rs("zl_news_id")%>&z_p=1" target="_blank"><font color="#ff0000"><%call pinglunz(rs("zl_news_id"))%></font></td>
				 				 <td width="40" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;color:#ff0000;"><a href="ziliang_pr_hf.asp?zl_news_id=<%=rs("zl_news_id")%>" target="_blank"><font color="#0099ff"><%call pinglunn(rs("zl_news_id"))%></font></a></td>
		  <td width="40" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("hits")%></td>
        <td width="80" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("zl_news_user")%></td>
        <td width="40" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%=rs("zl_news_paixu")%></td>
        
		<td width="140" align="center" valign="middle" style="border-bottom:1px solid #A5B9C9;border-right:1px solid #A5B9C9;"><%if request.querystring("i")<>"1" and rs("zl_news_lei_id")=2 then%><%if rs("zl_news_zhiding")=0 then%><a href="?Class=zd&zl_news_id=<%=rs("zl_news_id")%>"><font color="#0066FF">δ�ö�</font></a><%else%><a href="?Class=qd&zl_news_id=<%=rs("zl_news_id")%>"><font color="#FF0000">ȡ���ö�</font></a><%end if%>|<a href=ziliang_bowen.asp?zl_news_id=<%=rs("zl_news_id")%>&Edit=N_E&zl_news_Lei=<%=rs("zl_news_Lei")%>>�޸�</a>|<a href=?zl_news_id=<%=rs("zl_news_id")%>&Class=content_del onClick="return confirm('�Ƿ񽫴�ɾ��?');">ɾ��</a><%else%><a href=ziliang_bowen_2.asp?zl_news_id=<%=rs("zl_news_id")%>&Edit=N_E&zl_news_Lei=<%=rs("zl_news_Lei")%>>�޸�</a>|<a href=?zl_news_id=<%=rs("zl_news_id")%>&Class=content_del onClick="return confirm('�Ƿ񽫴�ɾ��?');">ɾ��</a><%end if%></td>
      </tr>
		<%
 rs.MoveNext
 next
 end if
 %>
 <tr>
 <td height=25 colspan="10" align="right" valign="middle" bgcolor="#F5F5F5"><div><span style="font-size: 9pt;"> ��&nbsp;<font color="#ff0000"><%=proCount%></font>&nbsp;��, ��ǰ�ڣ�<font color="#ff0000"><%=intCurPage%></font>/<font color="#ff0000"><%=rs.PageCount%></font>ҳ
                <% if intCurPage<>1 then %>
                <a href="ziliang_newsm.asp?cla=<%=nb%>&id=<%=request("id")%>&amp;ToPage=1" class="redlink">��ҳ</a>&nbsp;|&nbsp;<a href="ziliang_newsm.asp?cla=<%=nb%>&id=<%=request("id")%>&amp;ToPage=<%=intCurPage-1%>" class="redlink">��һҳ</a>&nbsp;|
                <%else%>
                <font color="#999999">��ҳ&nbsp;|&nbsp;��һҳ&nbsp;|
                <% end if%>
                <%if intCurPage<>rs.PageCount then %>
                <a href="ziliang_newsm.asp?cla=<%=nb%>&id=<%=request("id")%>&amp;ToPage=<%=intCurPage+1%>" class="redlink">��һҳ</a>&nbsp;| <a href="ziliang_newsm.asp?cla=<%=nb%>&id=<%=request("id")%>&amp;ToPage=<%=rs.PageCount%>" class="redlink"> ĩҳ</a>&nbsp;|&nbsp;
                <%else%>
                <font color="#999999">��һҳ&nbsp;|&nbsp;ĩҳ&nbsp;|&nbsp; </font>
                <% end if%></div></td></tr>
    </table>
</body>
</html>
