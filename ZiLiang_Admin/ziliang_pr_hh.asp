<!--#include file="../ziliang_inc/inc.asp"-->
<%
IF request.QueryString("bohf")="bohf" then

	Set Rshf=server.createobject("ADODB.Recordset")
	Sqlhf="select * from zl_message where m_id="&request.form("m_id")
	Rshf.open Sqlhf,conn,1,3
	rshf("m_hf")=request.form("m_hf")
    rshf("m_sh")=1
	rshf.update
	response.write"<script language='javascript'>alert('�ɹ��ظ�');history.go(-1);</script>"
    rshf.close
End if
%>
<%
If request.QueryString("bohf")="del" Then'ɾ��
		Sql="delete * from zl_message where m_id="&request("m_id")  
		set Rs=server.CreateObject("adodb.recordset")
		conn.execute Sql
		Response.Redirect "ziliang_pr_hf.asp?Lei=��ʾ��ɾ���ɹ�&z_p=1&zl_news_id="&request("zl_news_id")
End If

%>