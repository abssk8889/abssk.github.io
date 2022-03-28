<!--#include file="../ziliang_inc/inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>评论审核</title>
</head>
<body>

<table width="500" border="0" cellpadding="0" cellspacing="0" style="border:1px solid #CCCCCC;" align="center">
  <!--DWLayoutTable-->
  <tr>
    <td width="500" height="30" align="left" valign="middle" bgcolor="#F6F6F6" style="padding:5px;"><strong>评论审核</strong></td>
  </tr>
    <tr>
    <td width="500" height="30" align="left" valign="middle" style="padding:5px;color:#ff0000;" >标题：<%call nrbt(request("zl_news_id") )%></td>
  </tr>
  <%dim n
    n=0
	Set Rshf=server.createobject("ADODB.Recordset")
	if request.querystring("z_p")="1" then
	Sqlhf="select * from zl_message where m_nid="&request("zl_news_id") 
	else
	Sqlhf="select * from zl_message where m_nid="&request("zl_news_id")&" and m_sh=0"
	end if
	Rshf.open Sqlhf,conn,1,1
	do while not rshf.eof
	n=n+1

%>
  <tr>
    <td height="30" align="left" valign="middle" style="padding:5px;">序号：<font color="#FF0000"><%=n%></font> 由 <font color="#FF0000"><%=rshf("m_n")%></font> 评论 时间: <font color="#FF0000"><%=rshf("m_time")%></font> <a href="ziliang_pr_hh.asp?bohf=del&m_id=<%=rshf("m_id")%>&zl_news_id=<%=rshf("m_nid")%>">删除</a></td>
  </tr>
  <tr>
    <td height="30" align="left" valign="middle" style="padding:5px;">内容：<%=rshf("m_nr")%></td>
  </tr>
  <tr>
    <td height="40" valign="top" style="padding:5px;">
      <form id="form1" name="form1" method="post" action="ziliang_pr_hh.asp?bohf=bohf">
        回复：<textarea name="m_hf" cols="40" rows="5"><%=rshf("m_hf")%></textarea>
		<input type="hidden" name="m_id" value="<%=rshf("m_id")%>">
            <input type="submit" name="Submit" value="回复审核" />
      </form>
    </td>
  </tr>
 <%rshf.movenext
 loop
  rshf.close  
   %>
</table>

</body>
</html>
