<!--#include file="CE.asp"-->
<!--#include file="ziliang_conn.asp"-->
<!--#include file="Md5.asp"-->
<!--#include file="ziliang_function.asp"-->
<%
IF session("username")="" or session("userpas")="" Then
Response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');top.location.href='/ZiLiang_Admin/index.asp';</script>"
Response.End
End IF
%> 

