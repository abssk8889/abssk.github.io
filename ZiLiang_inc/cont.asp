<!--#include file="c.asp"-->
<%  id=cint(request("id"))
    IF request.querystring("ck")="ck" then
	
	sqlStr="update zl_content set hits=hits+1 where zl_news_id="&id
	conn.execute(sqlStr)
	
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from zl_content where zl_news_id="&id
		Rs.open Sql,conn,1,1
		hits=rs("hits")
		rs.close
		rs=nothing
	%>
document.write("<%=hits%>");
	<%
	Else IF request.querystring("pr")="pr" then 
	      
			Set Rs=server.createobject("ADODB.Recordset")
			Sql="select * from zl_message where m_nid="&id
			Rs.open Sql,conn,1,3
			przs=rs.recordcount
			rs.close
	%>
	document.write("<%=przs%>");
	<%
	Else IF request.querystring("prl")="prl" then
	
	      Set Rs=server.createobject("ADODB.Recordset")
			Sql="select * from zl_message where m_nid="&id&" and m_sh=1"
			Rs.open Sql,conn,1,1
			do while not rs.eof
			nrr=rs("m_nr")
			xmm=rs("m_n")
			m_hf=rs("m_hf")
			sjj=rs("m_time")
           prhf=prhf&"<li><font>留言者：<font color='#FF3300'>"&xmm&"</font> 时间：<font color='#FF3300'>"&sjj&"</font></font></li><li><p>内容:"&nrr&"<br>回复："&m_hf&"</p></li><li><hr></li>"
		   rs.movenext
		   loop
		   rs.close
			%>
			
	document.write("<%=prhf%>");
	
	<%		
	End IF
	End IF
	End IF
	%>