<%
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'  ��ʽȨ����:�� ����                                  '
	'  ��ϵ�ˣ��� ����                                     '
	'  ��ϵQQ��41864438                                    '
	'  msn: zjxyk@hotmail.com                              '
	'  ��ַ��aliang.zjxyk.com                             '
	'''''''''''''''''''''''''''''''''''''''' '''''''''''''''
IF Request.Querystring("Class")="pas" Then'����
			Set Rs=server.createobject("ADODB.Recordset")
			Sql="select * from zl_admin where zl_user='"&trim(request.Form("user_name"))&"'"
			Rs.open Sql,conn,1,3
		 IF Not (Rs.Eof and Rs.Bof) or request.Form("user_name")="" Then
			 response.write "<script language='javascript'>alert('�ظ���ӻ��û���Ϊ�գ��뷵�ز鿴');history.go(-1);</script>"
		 Else
			Rs.addnew
			Rs("zl_user") = trim(Request.Form("user_name"))
			Rs("zl_password") = md5(Request("user_pas"))
			Rs("zl_flag") = Request.Form("clkj_flg")
			Rs.update
			Response.Redirect "ziliang_user.asp?Lei=��ʾ�����ӳɹ�"
			Rs.close
		 End IF
Else IF Request.Querystring("Class")="pas_Edit" Then'�޸�
         	Set Rs=server.createobject("ADODB.Recordset")
			Sql="select * from zl_admin where zl_id="&request("zl_id")
			Rs.open Sql,conn,1,3
			Rs("zl_user") = trim(Request.Form("user_name"))
			IF trim(request.Form("user_pas"))="" or len(request.form("user_pas"))=0 Then
			Rs("zl_password")=Rs("zl_password")
			Else
			Rs("zl_password") = Md5(Request.Form("user_pas"))
			End IF
			Rs("zl_flag") = Request.Form("zl_flag")
			Rs.update
			Response.Redirect "ziliang_user.asp?Lei=��ʾ�����ӳɹ�"
			Rs.close
Else If request.QueryString("Class")="pas_Del" Then'ɾ��
		Sql="delete * from zl_admin where zl_id="&request("zl_id")  
		set Rs=server.CreateObject("adodb.recordset")
		conn.execute Sql
		Response.Redirect "ziliang_user.asp?Lei=��ʾ��ɾ���ɹ�"
End If
End IF
End IF

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
IF request.Querystring("Class")="Ladd" then'������Ŀ����
			Set Rs=server.createobject("ADODB.Recordset")
			Sql="select * from zl_small_class where zl_s_name='"&trim(request.Form("zl_s_name"))&"'"
			Rs.open Sql,conn,1,3
		 IF Not (Rs.Eof and Rs.Bof) or request.Form("zl_s_name")="" or request.form("zl_big")="" Then
			 response.write "<script language='javascript'>alert('�ظ���ӻ�Ϊ�գ��뷵�ز鿴');history.go(-1);</script>"
		 Else
			Rs.addnew
			Rs("zl_b_id") = cint(trim(Request.Form("zl_big")))
			Rs("zl_s_name") = trim(request.Form("zl_s_name"))
			Rs("zl_s_key") = trim(request.Form("zl_s_key"))
			Rs("zl_s_des") = trim(request.Form("zl_s_des"))
            Rs("zl_s_s_name") = c(trim(request.form("zl_s_name")))
			RS("zl_paixu") = cint(trim(request.form("zl_paixu")))
			Rs("zl_adress") = trim(request.form("zl_adress"))
			Rs("zl_s_bt") = trim(request.form("zl_s_bt"))
			if request.form("flag")="yes" then
			Rs("flag")=1
			else 
			rs("flag")=0
			end if
			Rs.update
			Response.Redirect "ziliang_lanmu.asp?Lei=��ʾ����ӳɹ�&add=add"
			Rs.close
		  End IF
		  
Else IF request.Querystring("Class")="Led" then '������Ŀ�޸�
			Set Rs=server.createobject("ADODB.Recordset")
			Sql="select * from zl_small_class where zl_s_id="&request("zl_s_id")
			Rs.open Sql,conn,1,3
			Rs("zl_b_id") = cint(trim(Request.Form("zl_big")))
			Rs("zl_s_name") = trim(request.Form("zl_s_name"))
			Rs("zl_s_key") = trim(request.Form("zl_s_key"))
			Rs("zl_s_des") = trim(request.Form("zl_s_des"))
            Rs("zl_s_s_name") = c(trim(request.form("zl_s_name")))
			RS("zl_paixu") = cint(trim(request.form("zl_paixu")))
			Rs("zl_adress") = trim(request.form("zl_adress"))
			Rs("zl_s_bt") = trim(request.form("zl_s_bt"))
			if request.form("flag")="yes" then
			Rs("flag")=1
			else 
			rs("flag")=0
			end if
			
			call edit(request("zl_s_id"),trim(request.Form("zl_s_name")))'�޸����ݱ��е������
			
	
			Rs.update
		    Response.Redirect "ziliang_class.asp"
			Rs.close
Else IF request.querystring("Class")="Ldel" then
		Sql="delete * from zl_small_class where zl_s_id="&request("zl_s_id")  
		set Rs=server.CreateObject("adodb.recordset")
		conn.execute Sql
		Response.Redirect "ziliang_class.asp?Lei=��ʾ��ɾ���ɹ�"
End IF
End IF
End IF

'============================== �������� ===============================================================================
IF Request.Querystring("Class")="links" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_Links where clkj_title='"&trim(request.Form("clkj_name"))&"'"
	Rs.open Sql,conn,1,3
  If Not (Rs.Eof and Rs.Bof) or request.Form("clkj_name")="" Then
response.write "<script language='javascript'>alert('�ظ���ӻ�������Ϊ�գ��뷵�ز鿴');history.go(-1);</script>"
	Else
	Rs.Addnew
	Rs("clkj_title") = trim(Request.Form("clkj_name"))
	Rs("clkj_LinkURL") = Request.Form("clkj_links")
	Rs("zl_l_paixu") = trim(request.form("zl_l_paixu"))
	Rs("zl_l_ms") = trim(request.form("zl_l_miaosu"))
	Rs.update
	Response.Redirect "ziliang_links.asp?Lei=��ʾ�����Ӹĳɹ�"
	Rs.close
  End If
Else IF Request.Querystring("Class")="links_Edit" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_Links where clkj_yqId="&request("clkj_yqId")
	Rs.open Sql,conn,1,3
    Rs("clkj_title") = trim(Request.Form("clkj_name"))
	Rs("clkj_LinkURL") = Request.Form("clkj_links")
	Rs("zl_l_paixu") = trim(request.form("zl_l_paixu"))
	Rs("zl_l_ms") = trim(request.form("zl_l_miaosu"))
	Rs.update
	Response.Redirect "ziliang_links.asp?Lei=��ʾ���޸ĳɹ�"
	Rs.close
Else If request.QueryString("Class")="links_Del" Then'��Ŀɾ��
		Sql="delete * from zl_Links where clkj_yqId="&request("clkj_yqId")  
		set Rs=server.CreateObject("adodb.recordset")
		conn.execute Sql
		Response.Redirect "ziliang_links.asp?Lei=��ʾ��ɾ���ɹ�"
End If			
End If
End If

'''''''''''''''''''''''''''''''''''''''''''''''�û�����'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub user()'�û�����,��̨�г������û�	
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from zl_admin"
		Rs.open Sql,conn,1,1	
	If Rs.Eof and Rs.Bof Then
		response.write "<tr><td style='padding:5px; height:30px;'>�����û�</td></tr>"
	Else
		Do while not Rs.Eof
		n=n+1
		IF Rs("zl_flag")=1 then
		flag="����Ա"
		Else
		flag="��ͨԱ"
		End if
		response.write "<td height='30' style='padding:4px;' style='border-bottom:1px dashed #d4d4d4;color:#FF0000'>"&flag&"&nbsp;"&Rs("zl_user")&"&nbsp;[<a href='ziliang_user.asp?zl_id="&Rs("zl_id")&"&Edit=B_E#xy'>�޸�</a>&nbsp;<a href='ziliang_user.asp?zl_id="&Rs("zl_id")&"&Class=pas_Del' onclick="&chr(34)&"return confirm('�Ƿ񽫴�ɾ��?');"&chr(34)&">ɾ��</a>]&nbsp;</td>"
		if n mod 5=0 then
		j=j+1
		response.write("<tr></tr>")
		if j=100 then exit do
		end if
		Rs.Movenext
		Loop
		Rs.close		
	End If				
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''��̨������ʾ'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FenLei()
 set rs=server.createobject("adodb.recordset")
		exec="select * from zl_big_class order by zl_b_id asc"		
		rs.open exec,conn,1,1
		do while not rs.eof
		big_name=rs("zl_b_name")
		big_id = rs("zl_b_id")
		response.write "<tr><td width="&chr(34)&"5%"&chr(34)&" align="&chr(34)&"center"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" bgcolor="&chr(34)&"#ECF6FC"&chr(34)&"><input type="&chr(34)&"checkbox"&chr(34)&" name="&chr(34)&"checkbox"&chr(34)&" value="&chr(34)&"checkbox"&chr(34)&" /></td><td width="&chr(34)&"30%"&chr(34)&" height="&chr(34)&"30"&chr(34)&" align="&chr(34)&"left"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" bgcolor="&chr(34)&"#ECF6FC"&chr(34)&" style="&chr(34)&"padding:2px;"&chr(34)&"><strong>"&big_name&"</strong></td><td align="&chr(34)&"left"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" bgcolor="&chr(34)&"#ECF6FC"&chr(34)&" style="&chr(34)&"padding:2px;"&chr(34)&"><a href="&chr(34)&"ziliang_lanmu.asp?zl_b_id="&big_id&"&add=add"&chr(34)&">��Ӷ�����Ŀ</a> </td><td align="&chr(34)&"left"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" bgcolor="&chr(34)&"#ECF6FC"&chr(34)&" style="&chr(34)&"padding:2px;"&chr(34)&">&nbsp;</td></tr>"
			set rs_1=server.createobject("adodb.recordset")
			exec_1="select * from zl_small_class where zl_b_id="&big_id&" order by zl_paixu,zl_s_id asc"		
			rs_1.open exec_1,conn,1,1
			do while not rs_1.eof
			small_name = rs_1("zl_s_name")
			small_idddd=rs_1("zl_s_id")
			if rs_1("flag")=1 then
			F="�ѿ���"
			else
			F="�ѹر�"
			end if
			IF big_id=1 then
			sc="<a href="&chr(34)&"html/ziliang_nav.asp?zl_s_id="&small_idddd&""&chr(34)&" onclick="&chr(34)&"return confirm('�Ƿ����ɵ�ҳ?');"&chr(34)&">���ɵ�ҳ</a>"
			else
			sc="<a href="&chr(34)&"/ZiLiang_Admin/html/ziliang_liebiao.asp?zl_s_id="&rs_1("zl_s_id")&""&chr(34)&">������Ŀ</a> <a href="&chr(34)&"html/ziliang_aspTOhtml_all.asp?zl_s_id="&small_idddd&"&s=zong"&chr(34)&" onclick="&chr(34)&"return confirm('�Ƿ�����ȫ����ҳ?');"&chr(34)&">����ȫ����ҳ</a> <a href="&chr(34)&"html/ziliang_aspTOhtml_all.asp?zl_s_id="&small_idddd&"&s=day"&chr(34)&" onclick="&chr(34)&"return confirm('�Ƿ����ɵ���ҳ��?');"&chr(34)&"><font color=#ff0000>���ɵ���ҳ��</font></a>"
			end if
			
			
			response.write "<tr><td width="&chr(34)&"5%"&chr(34)&" align="&chr(34)&"center"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" bgcolor="&chr(34)&"#F5F5F5"&chr(34)&">"&rs_1("zl_paixu")&"</td><td width="&chr(34)&"30%"&chr(34)&" height="&chr(34)&"30"&chr(34)&" align="&chr(34)&"left"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" bgcolor="&chr(34)&"#F5F5F5"&chr(34)&" style="&chr(34)&"padding:2px;"&chr(34)&"> ��| "&small_name&"</td><td align="&chr(34)&"left"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" bgcolor="&chr(34)&"#F5F5F5"&chr(34)&" style="&chr(34)&"padding:2px;"&chr(34)&"><a href="&chr(34)&"ziliang_lanmu.asp?zl_b_id="&big_id&"&zl_s_id="&rs_1("zl_s_id")&"&add=ed"&chr(34)&">�޸���Ŀ</a> <a href='?Class=Ldel&zl_s_id="&rs_1("zl_s_id")&"' onclick="&chr(34)&"return confirm('�Ƿ񽫴�ɾ��?');"&chr(34)&"><font color="&chr(34)&"#ff0000"&chr(34)&">ɾ��</font></a> "&sc&" </td><td align="&chr(34)&"left"&chr(34)&" valign="&chr(34)&"middle"&chr(34)&" bgcolor="&chr(34)&"#F5F5F5"&chr(34)&" style="&chr(34)&"padding:2px;"&chr(34)&">"&F&"</td></tr>"
			rs_1.movenext			
			loop
			rs_1.close	
		rs.movenext
		loop
		rs.close
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Big_input()'����һ����Ŀ����Ŀ������ʽ
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from zl_big_class"
		Rs.open Sql,conn,1,1	
	If Rs.Eof and Rs.Bof Then
		response.write "������Ŀ"
	Else
		Do while not RS.Eof
		response.write "<option value="
		response.write rs("zl_b_id")
		IF rs("zl_b_id")=cint(request("zl_b_id")) then
		response.write" selected=selected"
		end if
		response.write">"
		response.write rs("zl_b_name")
		response.write "</option>"

		Rs.Movenext
		Loop
	End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub small_input()'����һ����Ŀ����Ŀ������ʽ
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from zl_small_class where zl_b_id=2"
		Rs.open Sql,conn,1,1	
	If Rs.Eof and Rs.Bof Then
		response.write "������Ŀ"
	Else
		Do while not RS.Eof
		response.write "<option value="
		response.write rs("zl_s_id")
		IF rs("zl_s_id")=cint(request("zl_news_Lei")) then
		response.write" selected=selected"
		end if
		response.write">"
		response.write rs("zl_s_name")
		response.write "</option>"
		Rs.Movenext
		Loop
	End If
End Sub

Sub small_input2()'����һ����Ŀ����Ŀ������ʽ
		Set Rs=server.createobject("ADODB.Recordset")
		Sql="select * from zl_small_class where zl_b_id=1"
		Rs.open Sql,conn,1,1	
	If Rs.Eof and Rs.Bof Then
		response.write "������Ŀ"
	Else
		Do while not RS.Eof
		response.write "<option value="
		response.write rs("zl_s_id")
		IF rs("zl_s_id")=cint(request("zl_news_Lei")) then
		response.write" selected=selected"
		end if
		response.write">"
		response.write rs("zl_s_name")
		response.write "</option>"
		Rs.Movenext
		Loop
	End If
End Sub


''''''''''''''''''''''''''''��վ��������''''''''''''''''''''''''''''''''''''''''''''

IF request.querystring("Class")="config" then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_config where config_id="&request("config_id")
	Rs.open Sql,conn,1,3
	Rs("zl_web_zhanzhang") = Request.Form("zl_web_zhanzhang")
	Rs("zl_wel_logo") = Request.Form("logo")
	Rs("zl_web_title") = Request.Form("name")	
	Rs("zl_web_miaosu") = Request.Form("maiosu")
	Rs("zl_web_wangzi") = Request.Form("wangzhi")
	Rs("zl_web_banquan") = Request.Form("banquan")
	Rs("zl_wel_liebiao") = Request.Form("liebaio")
	Rs("zl_web_ikey") = Request.Form("ikey")
	Rs("zl_web_ides") = Request.Form("ides")
		IF Request.Form("links") = "yes" Then
		Rs("zl_web_links_off") = 1
		Else
		Rs("zl_web_links_off") = 0
		End IF
	Rs("zl_wel_liebaio_flag") = Request.Form("xianshi")
	Rs("zl_wel_height") = Request.Form("height")
	Rs("zl_wel_width") = Request.Form("width")
		IF Request.Form("pr") = "yes" Then'��Ŀ���ܿ���
		Rs("zl_web_pr") = 1
		Else
		Rs("zl_web_pr") = 0
		end if
		IF Request.Form("cd") = "yes" Then
		Rs("zl_web_cd") = 1
		Else 
		Rs("zl_web_cd") = 0
		end if
		IF Request.Form("wz") = "yes" Then	
		Rs("zl_web_wz") = 1
		else
		Rs("zl_web_wz") = 0
		end if
	Rs.update
	Response.Redirect"ziliang_config.asp?Lei=��ʾ���޸ĳɹ�"
	Rs.close	
End IF

IF request.querystring("Class")="content" then'��Ϣ���
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_content"
	Rs.open Sql,conn,1,3
	If request.Form("lei")=""or request.form("title")="" or request.form("des")="" or request.form("key")="" Then
response.write "<script language='javascript'>alert('��*�ŵ�Ϊ����');history.go(-1);</script>"
	Else
	Rs.Addnew
	Rs("zl_news_Lei") = Request.Form("lei")
			Set Rs_2=server.createobject("ADODB.Recordset")
			Sql_2="select * from zl_small_class where zl_s_id="&Request.Form("lei")
			Rs_2.open Sql_2,conn,1,1
			N=rs_2("zl_s_name")
			m=rs_2("zl_b_id")'����id
			rs_2.close
	Rs("zl_news_lei_name") = N'��������
	Rs("zl_news_lei_id") = cint(m)'����id
	Rs("zl_news_title") = Request.Form("title")
	Rs("zl_news_key") = Request.Form("key")
	Rs("zl_news_maiosu") = Request.Form("des")
	Rs("zl_news_nr") = Request.Form("content")
	Rs("zl_news_user") = Request.Form("user")
	Rs("zl_news_paixu") = Request.Form("paixu")
	Rs.update
	if m=2 then
	Response.Redirect"ziliang_bowen.asp?Lei=��ʾ�����ӳɹ�"
	else
	Response.Redirect"ziliang_bowen_2.asp?Lei=��ʾ�����ӳɹ�"
	end if
	Rs.close
	end if
	
Else IF request.querystring("Class")="content_edit" then'��Ϣ�޸�
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_content where zl_news_id="&request("zl_news_id")
	Rs.open Sql,conn,1,3
	Rs("zl_news_Lei") = Request.Form("lei")
			Set Rs_2=server.createobject("ADODB.Recordset")
			Sql_2="select * from zl_small_class where zl_s_id="&Request.Form("lei")
			Rs_2.open Sql_2,conn,1,1
			N=rs_2("zl_s_name")
			m=rs_2("zl_b_id")'����id
			rs_2.close
	Rs("zl_news_lei_name") = N'��������
	rs_2("zl_b_id")=cint(m)'����id
	Rs("zl_news_title") = Request.Form("title")
	Rs("zl_news_key") = Request.Form("key")
	Rs("zl_news_maiosu") = Request.Form("des")
	Rs("zl_news_nr") = Request.Form("content")
	Rs("zl_news_user") = Request.Form("user")
	Rs("zl_news_paixu") = Request.Form("paixu")
	Rs.update
	call shencheng(cint(request("zl_news_id")))'�����޸�ҳ������
	Response.Redirect"ziliang_newsm.asp?Lei=��ʾ���޸ĳɹ�"
	Rs.close	
	
Else IF request.querystring("Class")="content_del" then
        Sql="delete * from zl_content where zl_news_id="&request("zl_news_id")
		set Rs=server.CreateObject("adodb.recordset")
		conn.execute Sql
		Response.Redirect "ziliang_newsm.asp?Lei=��ʾ��ɾ���ɹ�"
		
Else IF request.querystring("Class")="zd" then'��Ϣ�ö�
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_content where zl_news_id="&request("zl_news_id")
	Rs.open Sql,conn,1,3
	Rs("zl_news_zhiding")=1
	Rs.update
	Rs.close
Else IF request.querystring("Class")="qd" then'ȡ����Ϣ�ö�
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_content where zl_news_id="&request("zl_news_id")
	Rs.open Sql,conn,1,3
	Rs("zl_news_zhiding")=0
	Rs.update
	Rs.close
End IF
End IF
End IF
End IF
End IF


'''''''''''''''''''''''''''''''''��Ŀ�޸ģ��޸���Ӧ�����е���Ŀ����''''''''''''''''''''''''''''''''''''''''''''''
Sub edit(id,n)
	conn.execute("update zl_content set zl_news_lei_name='"&n&"'  where zl_news_Lei="&cint(id))
End Sub

sub small_lei()'��Ϣ������Ŀ
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_small_class where zl_b_id=2"
	Rs.open Sql,conn,1,1
	do while not rs.eof
	if cint(request("id"))=rs("zl_s_id")then
	response.write "<a href=?ID="&rs("zl_s_id")&"><b><font color=#ff0000>"&rs("zl_s_name")&"</font></B></a>&nbsp;"
	else
	response.write "<a href=?ID="&rs("zl_s_id")&">"&rs("zl_s_name")&"</a>&nbsp;"
	end if
	rs.movenext
	loop
	rs.close
End sub

sub small_lei2()'��Ϣ������Ŀ
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_small_class where zl_b_id=1"
	Rs.open Sql,conn,1,1
	do while not rs.eof
	if cint(request("id"))=rs("zl_s_id")then
	response.write "<a href=?ID="&rs("zl_s_id")&"&i=1><b><font color=#ff0000>"&rs("zl_s_name")&"</font></B></a>&nbsp;"
	else
	response.write "<a href=?ID="&rs("zl_s_id")&"&i=1>"&rs("zl_s_name")&"</a>&nbsp;"
	end if
	rs.movenext
	loop
	rs.close
End sub

'''''''''''''''''''''''''������������'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
IF request.querystring("Class")="Message" Then
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_message"
	Rs.open Sql,conn,1,3
	If request.form("m_nr")="" or request.form("m_na")="" or request.form("m_nid")="" Then
    response.write "<script language='javascript'>alert('��*�ŵ�Ϊ����');history.back(1);</script>"
	Else
	Rs.Addnew
	Rs("m_ip") = Request.ServerVariables("REMOTE_ADDR")
	Rs("m_nr") = Request.Form("m_nr")
	Rs("m_n") = Request.Form("m_na")
	Rs("m_nid") = Request.Form("m_nid")
	Rs.update	
			Set Rs_nc=server.createobject("ADODB.Recordset")'����ԭ·��
			Sql_nc="select * from zl_content where zl_news_id="&Request.Form("m_nid")
			Rs_nc.open Sql_nc,conn,1,1
			ncjm=c(Rs_nc("zl_news_lei_name"))
			nrt=year(rs_nc("zl_news_time"))
			nrid=nrt+Request.Form("m_nid")
			rs_nc.close				
	response.write"<script language='javascript'>alert('�ɹ�����,��˺���ʾ');window.location.href='/"&ncjm&"/"&nrid&".html';</script>"
	Rs.close
	End if
End IF


Sub back(bid)'ͨ����Ϣ���е�ƴ�������ض���������Ϣ
	Set Rs_n=server.createobject("ADODB.Recordset")
	Sql_n="select * from zl_small_class where zl_s_id="&bid
	Rs_n.open Sql_n,conn,1,1
	response.write rs_n("zl_s_name")
	Rs_n.close	
End sub

Sub pinglunn(ppid)'ͳ��δ������
	Set Rs_pl=server.createobject("ADODB.Recordset")
	Sql_pl="select * from zl_message where m_nid="&ppid&" and m_sh=0 "
	Rs_pl.open Sql_pl,conn,1,3
	pinglun=Rs_pl.recordcount
	Response.write pinglun
	Rs_pl.close	
End sub
Sub pinglunz(zpid)'ͳ����������
	Set Rs_pz=server.createobject("ADODB.Recordset")
	Sql_pz="select * from zl_message where m_nid="&zpid
	Rs_pz.open Sql_pz,conn,1,3
	pinglzn=Rs_pz.recordcount
	Response.write pinglzn
	Rs_pz.close	
End sub

Sub nrbt(nrid)'�������ݱ��еı���
	Set Rs_n=server.createobject("ADODB.Recordset")
	Sql_n="select * from zl_content where zl_news_id="&nrid
	Rs_n.open Sql_n,conn,1,1
	response.write rs_n("zl_news_title")
	Rs_n.close	
End sub

Sub pinglunhz()'ͳ����������
	Set Rs_hz=server.createobject("ADODB.Recordset")
	Sql_hz="select * from zl_message where m_sh=0"
	Rs_hz.open Sql_hz,conn,1,3
	pinglhz=Rs_hz.recordcount
	response.write"����<font color=#ff0000>"&pinglhz&"</font>��δ�ظ�<br><br>"
	Do while not rs_hz.eof
	response.write rs_hz("m_nid")&"&nbsp;"
	rs_hz.movenext
	loop
	Rs_hz.close	
End sub



'��̨����
IF Request.Querystring("Class")="userpas" Then
	Dim username,password
	username=replace(trim(request.Form("name")),"'","")
	password=md5(replace(trim(request.Form("password")),"'",""))
		Set Rs=server.createobject("adodb.recordset")
		Sql="select  * from  zl_admin  where zl_user='"&username&"' and zl_password='"&password&"'"
		Rs.open sql,conn,1,1
			IF Rs.eof or Rs.bof then
				  response.write "<script language='javascript'>alert('�û�����������󣡣���');top.location.href='index.asp';</script>"
				  rs.close
				  set rs=nothing
				  response.end()
			Else 
				session("username")=rs("zl_user")
				session("userpas")=rs("zl_password")
				session.Timeout=60
			End IF
		Response.Redirect("ziliang_admin.asp")
		Rs.close
		Set Rs=nothing
End IF

'��ȫ�˳�
IF request.querystring("class")="Out" Then

	session("username") = ""
	session("userpas") = ""
	response.write "<script language='javascript'>alert('��ȫ�˳������ص�½ҳ��');top.location.href='index.asp';</script>"
End IF

''''''''''''''''''''�޸�ҳ������'''''''''''''''''''''
sub shencheng(shengchengid)'�޸�����ҳ��
	set rs_s=Server.CreateObject("ADODB.Recordset")
	sql_s="select * from zl_content where zl_news_id="&shengchengid
	rs_s.open sql_s,conn,1,1
	if not rs.eof then
		idd=rs_s("zl_news_id")
		pp=c(rs_s("zl_news_lei_name"))
		yy=year(rs_s("zl_news_time"))
		'mm=month(rs_s("zl_news_time"))
		Nn=yy+idd'��+id��Ϊ�ļ���		
	Call Clkj_asptohtml(idd,pp,(Nn&".html"))
	response.write"<script language='javascript'>alert('�������');history.go(-1);</script>"
	end if
end sub

%>