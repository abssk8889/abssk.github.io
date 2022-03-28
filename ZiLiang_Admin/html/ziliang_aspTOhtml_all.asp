    <!--#include file="../../ziliang_inc/inc.asp"-->
	<!--#include file="ziliang_aTOh.asp"-->
	
<%
if request.QueryString("s")="zong" then	
	set rs_s=Server.CreateObject("ADODB.Recordset")
	sql_s="select * from zl_content where zl_news_Lei="&request("zl_s_id")
	rs_s.open sql_s,conn,1,1
	do while not rs_s.eof
		idd=rs_s("zl_news_id")
		pp=c(rs_s("zl_news_lei_name"))
		yy=year(rs_s("zl_news_time"))
		'mm=month(rs_s("zl_news_time"))
		Nn=yy+idd'年+id作为文件名		
	Call Clkj_asptohtml(idd,pp,(Nn&".html"))	
	rs_s.movenext
	loop
	response.write"<script language='javascript'>alert('生成完毕');history.go(-1);</script>"
	conn.close
	set conn = nothing
Else

	set rs_s=Server.CreateObject("ADODB.Recordset")
	sql_s="select * from zl_content where zl_news_Lei="&request("zl_s_id")
	rs_s.open sql_s,conn,1,1
	do while not rs_s.eof
		idd=rs_s("zl_news_id")
		pp=c(rs_s("zl_news_lei_name"))
		yy=year(rs_s("zl_news_time"))
		t=rs_s("zl_news_time")
		Nn=yy+idd'年+id作为文件名
				
	if formatdatetime (now(),2)=formatdatetime(t,2) then'生成当天页面	
	Call Clkj_asptohtml(idd,pp,(Nn&".html"))	
	end if
	
	rs_s.movenext
	loop
	response.write"<script language='javascript'>alert('生成完毕');history.go(-1);</script>"
	conn.close
	set conn = nothing
End IF	

	%>
