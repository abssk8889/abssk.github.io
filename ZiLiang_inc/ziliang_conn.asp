<% '������ı�ƥ��
Function RemoveHTML(strHTML)
Dim objRegExp, Match, Matches 
Set objRegExp = New Regexp
strHTML=Replace(strHTML,vbCrLf,"")
strHTML=Replace(strHTML,"��","")
strHTML=Replace(strHTML," ","")
strHTML=Replace(strHTML,"&nbsp;","")
objRegExp.IgnoreCase = True
objRegExp.Global = True
'ȡ�պϵ�<>
objRegExp.Pattern = "<.+?>"
'����ƥ��
Set Matches = objRegExp.Execute(strHTML)
' ����ƥ�伯�ϣ����滻��ƥ�����Ŀ
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"")
Next
RemoveHTML=strHTML
Set objRegExp = Nothing
End Function
%>

<%
dim conn
dim connstr
dim db
db="/ZiLiang_DaTa/#B_ZiLiang.mdb" '���ݿ��ļ�λ��
on error resume next
connstr="DBQ="+server.mappath(db)+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
if err then
err.clear
else
conn.open connstr
end if
aliang="Powered by <a href='http://aliang.zjxyk.com/' target='_blank' title='��������'>ALiang-boke</a> "
%>
<%URl=Request.ServerVariables("SERVER_NAME")%>

<%''''''''''''''''''''''''''''''''''''''''''''��վ����''''''''''''''''''''''''''''''''''''''''
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_config where config_id=1"
	Rs.open Sql,conn,1,1
		web_bt=Rs("zl_web_title")'��վ����
		miaosu = Rs("zl_web_miaosu")'��վ����
		wangzi = Rs("zl_web_wangzi")'��վ��ַ
		banquan = aliang&Rs("zl_web_banquan")'��վ��Ȩ
		logo = Rs("zl_wel_logo")'��վ��־
		liebiao = cint(Rs("zl_wel_liebiao"))'�б�������ʾ
		flag = Rs("zl_wel_liebaio_flag")'�б���ʾ��ʽ��0����ʾ�������������֣�1���������ĵ�
		linkss = Rs("zl_web_links_off")'�������ӿ���
		h = Rs("zl_wel_height")'ͼƬ�߶�
		w = Rs("zl_wel_width")'ͼƬ���
		pr = Rs("zl_web_pr")'���ۿ���
		cd = Rs("zl_web_cd")'�浵����
		wz = Rs("zl_web_wz")'�������¿���
		zz = Rs("zl_web_zhanzhang")'վ��
		ikey=rs("zl_web_ikey")'��ҳ�ؼ���
		ides=rs("zl_web_ides")'��ҳ����
	Rs.close
	Set Rs=Nothing
	
   ''''''''''''''''''''''''''''''''''''''''''������ʾ''''''''''''''''''''''''''''''''''''''''''''''

	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_small_class where zl_b_id=1 and flag=1 order by zl_paixu ,zl_s_id asc"'��С��������
	Rs.open Sql,conn,1,1
	Do while Not Rs.Eof	
	
		s_id=Rs("zl_s_id")
		s_s_name =rs("zl_s_s_name")'�ļ����ӵ�ַ
		s_name=Rs("zl_s_name")'��Ŀ����
		ss_id=20009+s_id
		Nav=Nav&" | <a href="&chr(34)&"/"&s_s_name&".html"&chr(34)&">"&s_name&"</a>"
		
	Rs.Movenext
	Loop
	Rs.close
	Set Rs=Nothing
	
   ''''''''''''''''''''''''''''''''''''''''''��Ŀ��ʾ''''''''''''''''''''''''''''''''''''''''''''''	

	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_small_class where zl_b_id=2 and flag=1 order by zl_paixu ,zl_s_id asc"'��С��������
	Rs.open Sql,conn,1,1
	Do while Not Rs.Eof	
		s_id=Rs("zl_s_id")
		s_name=Rs("zl_s_name")'��Ŀ����		
		s_s_name =rs("zl_s_s_name")'�ļ����ӵ�ַ
		LanMu=LanMu&" <li><p><a href="&chr(34)&"/"&s_s_name&"/"&s_s_name&"1.html"&chr(34)&">"&s_name&"</a></p></li>"		
	Rs.Movenext
	Loop
	LanMuu="<li><h2>�ҵ���Ŀ</h2></li>"	
	L=LanMuu+LanMu'��Ŀ����		
	Rs.close
	Set Rs=Nothing
	
''''''''''''''''''''''''''''''''''''''''''�����ĵ�''''''''''''''''''''''''''''''''''''''''''''''		
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select top 10 * from zl_content where zl_news_lei_id=2 order by zl_news_id desc"
	Rs.open Sql,conn,1,1
	Do while Not Rs.Eof
	bt=left(rs("zl_news_title"),14)
	zl_news_lei_name=c(rs("zl_news_lei_name"))
	
	zl_news_id=rs("zl_news_id")
	zl_news_time=year(rs("zl_news_time"))
	
	wjm=zl_news_id+zl_news_time'�ļ�������id+���
	Zx=Zx&"<li><p><a href="&chr(34)&"/"&zl_news_lei_name&"/"&wjm&".html"&chr(34)&">"&bt&"</a></p></li>"
	Rs.Movenext
	Loop
	Zxw="<li><h2>��������</h2></li>"	
	Z=Zxw+Zx'�����ĵ���Ŀ	
	Rs.close
	Set Rs=Nothing
	
	
    ''''''''''''''''''''''''''''''''''''''''''��������''''''''''''''''''''''''''''''''''''''''''''''	
	
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_links order by zl_l_paixu asc,clkj_yqId desc"
	Rs.open Sql,conn,1,1
	Do while Not Rs.Eof
	title=rs("clkj_title")
	LinkURL=rs("clkj_LinkURL")
	l_ms=rs("zl_l_ms")	
	links=links&"<li><p><a href="&chr(34)&""&LinkURL&""&chr(34)&" title="&chr(34)&""&l_ms&""&chr(34)&">"&title&"</a></p></li>"
	Rs.Movenext
	Loop
	link="<li><h2>��������</h2></li>"	
	Y=link+links'��������		
	Rs.close
	Set Rs=Nothing
	
	
    ''''''''''''''''''''''''''''''''''''''''''��ҳ�б���ʾ''''''''''''''''''''''''''''''''''''''''''''''
    ziding=0
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_content where zl_news_lei_id=2 order by zl_news_zhiding,zl_news_id desc"'���ö�����
	Rs.open Sql,conn,1,1
	Do while not rs.eof and ziding<liebiao
	ziding=ziding+1
	if flag=0 then'���б���ʾ��ʽ
	zl_news_nr=left(RemoveHTML(rs("zl_news_nr")),150)
	else
	zl_news_nr=rs("zl_news_nr")
	end if
	zl_news_user=rs("zl_news_user")
	zl_news_title=rs("zl_news_title")
	zl_news_lei_name=c(rs("zl_news_lei_name"))
	zl_news_id=rs("zl_news_id")
	keys=Cstr(rs("zl_news_time"))
	zl_news_time=year(rs("zl_news_time"))
	wjmm=zl_news_id+zl_news_time
	hits=rs("hits")
	 
	 
	Set Rs_n=server.createobject("ADODB.Recordset")
	Sql_n="select * from zl_small_class where zl_s_id="&rs("zl_news_Lei")
	Rs_n.open Sql_n,conn,1,1
	bb_nn=rs_n("zl_s_name")
	Rs_n.close
	
	
			Set Rsp=server.createobject("ADODB.Recordset")'ͳ��������
			Sqlp="select * from zl_message where m_nid="&zl_news_id
			Rsp.open Sqlp,conn,1,3
			prcs=rsp.recordcount
			rsp.close
	
			mystr=rs("zl_news_key")'�������۷�
			a=Split(mystr,",")
			For Each Strs In a		
			key=" <a href="&chr(34)&"/"&zl_news_lei_name&"/"&wjmm&".html"&chr(34)&" title="&zl_news_title&">"&Strs&"</a> "	
			next
			
	index_1=index_1&"<li><strong><a href="&chr(34)&"/"&zl_news_lei_name&"/"&wjmm&".html"&chr(34)&">"&zl_news_title&"</a></strong></li><li><span>"&zl_news_nr&"</span></li> <li><font>������ "&keys&" <a href="&chr(34)&"/"&zl_news_lei_name&"/"&wjmm&".html"&chr(34)&">�鿴</a>("&hits&") <a href="&chr(34)&"/"&zl_news_lei_name&"/"&wjmm&".html"&chr(34)&">����</a>("&prcs&") <span style="&chr(34)&"cursor:pointer"&chr(34)&" onClick="&chr(34)&"window.external.addFavorite('http://"&URl&"/"&zl_news_lei_name&"/"&wjmm&".html','"&zl_news_title&"')"&chr(34)&" title="&chr(34)&""&zl_news_title&""&chr(34)&">�ղ�</span> | <input type="&chr(34)&"text"&chr(34)&" class="&chr(34)&"inputys"&chr(34)&" value="&chr(34)&"http://"&URl&"/"&zl_news_lei_name&"/"&wjmm&".html"&chr(34)&" id="&chr(34)&""&zl_news_id&""&chr(34)&" readonly="&chr(34)&"readonly"&chr(34)&"> <a href="&chr(34)&"#"&chr(34)&" onmousedown="&chr(34)&"copyurl1("&zl_news_id&")"&chr(34)&">����</a> | ����:<a href="&chr(34)&"/"&zl_news_lei_name&"/"&zl_news_lei_name&"1.html"&chr(34)&">"&bb_nn&"</a> ������:"&key&"</font></li><li><hr></li>"	        
			Rs.movenext
			Loop
			
	Rs.close
	Set Rs=Nothing
	
''''''''''''''''''''''''''''''''''''''''''������Ŀ���ɵ�������''''''''''''''''''''''''''''''''''''''''''''''

	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_small_class where zl_s_id="&cint(request("zl_s_id"))
	Rs.open Sql,conn,1,1
    zl_s_s_name=rs("zl_s_s_name")'�ļ����ӵ�ַ
	zl_s_key=rs("zl_s_key")'��Ŀ������
	zl_s_des=rs("zl_s_des")'��Ŀ����
	zl_adress=rs("zl_adress")'ģ�����ӵ�ַ
	zl_s_name=rs("zl_s_name")'��������
	zl_s_bt=rs("zl_s_bt")
	rs.close
	Set rs=nothing
	
%>