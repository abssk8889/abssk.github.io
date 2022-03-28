<% '输出纯文本匹配
Function RemoveHTML(strHTML)
Dim objRegExp, Match, Matches 
Set objRegExp = New Regexp
strHTML=Replace(strHTML,vbCrLf,"")
strHTML=Replace(strHTML,"　","")
strHTML=Replace(strHTML," ","")
strHTML=Replace(strHTML,"&nbsp;","")
objRegExp.IgnoreCase = True
objRegExp.Global = True
'取闭合的<>
objRegExp.Pattern = "<.+?>"
'进行匹配
Set Matches = objRegExp.Execute(strHTML)
' 遍历匹配集合，并替换掉匹配的项目
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
db="/ZiLiang_DaTa/#B_ZiLiang.mdb" '数据库文件位置
on error resume next
connstr="DBQ="+server.mappath(db)+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
if err then
err.clear
else
conn.open connstr
end if
aliang="Powered by <a href='http://aliang.zjxyk.com/' target='_blank' title='阿梁网络'>ALiang-boke</a> "
%>
<%URl=Request.ServerVariables("SERVER_NAME")%>

<%''''''''''''''''''''''''''''''''''''''''''''网站参数''''''''''''''''''''''''''''''''''''''''
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_config where config_id=1"
	Rs.open Sql,conn,1,1
		web_bt=Rs("zl_web_title")'网站名称
		miaosu = Rs("zl_web_miaosu")'网站描述
		wangzi = Rs("zl_web_wangzi")'网站网址
		banquan = aliang&Rs("zl_web_banquan")'网站版权
		logo = Rs("zl_wel_logo")'网站标志
		liebiao = cint(Rs("zl_wel_liebiao"))'列表数量显示
		flag = Rs("zl_wel_liebaio_flag")'列表显示方式，0，显示标题与少量文字，1，标题与文档
		linkss = Rs("zl_web_links_off")'友情链接开启
		h = Rs("zl_wel_height")'图片高度
		w = Rs("zl_wel_width")'图片宽度
		pr = Rs("zl_web_pr")'评论开关
		cd = Rs("zl_web_cd")'存档开关
		wz = Rs("zl_web_wz")'最新文章开关
		zz = Rs("zl_web_zhanzhang")'站长
		ikey=rs("zl_web_ikey")'首页关键词
		ides=rs("zl_web_ides")'首页描述
	Rs.close
	Set Rs=Nothing
	
   ''''''''''''''''''''''''''''''''''''''''''导航显示''''''''''''''''''''''''''''''''''''''''''''''

	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_small_class where zl_b_id=1 and flag=1 order by zl_paixu ,zl_s_id asc"'由小到大排序
	Rs.open Sql,conn,1,1
	Do while Not Rs.Eof	
	
		s_id=Rs("zl_s_id")
		s_s_name =rs("zl_s_s_name")'文件链接地址
		s_name=Rs("zl_s_name")'栏目名称
		ss_id=20009+s_id
		Nav=Nav&" | <a href="&chr(34)&"/"&s_s_name&".html"&chr(34)&">"&s_name&"</a>"
		
	Rs.Movenext
	Loop
	Rs.close
	Set Rs=Nothing
	
   ''''''''''''''''''''''''''''''''''''''''''栏目显示''''''''''''''''''''''''''''''''''''''''''''''	

	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_small_class where zl_b_id=2 and flag=1 order by zl_paixu ,zl_s_id asc"'由小到大排序
	Rs.open Sql,conn,1,1
	Do while Not Rs.Eof	
		s_id=Rs("zl_s_id")
		s_name=Rs("zl_s_name")'栏目名称		
		s_s_name =rs("zl_s_s_name")'文件链接地址
		LanMu=LanMu&" <li><p><a href="&chr(34)&"/"&s_s_name&"/"&s_s_name&"1.html"&chr(34)&">"&s_name&"</a></p></li>"		
	Rs.Movenext
	Loop
	LanMuu="<li><h2>我的栏目</h2></li>"	
	L=LanMuu+LanMu'栏目汇总		
	Rs.close
	Set Rs=Nothing
	
''''''''''''''''''''''''''''''''''''''''''最新文档''''''''''''''''''''''''''''''''''''''''''''''		
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select top 10 * from zl_content where zl_news_lei_id=2 order by zl_news_id desc"
	Rs.open Sql,conn,1,1
	Do while Not Rs.Eof
	bt=left(rs("zl_news_title"),14)
	zl_news_lei_name=c(rs("zl_news_lei_name"))
	
	zl_news_id=rs("zl_news_id")
	zl_news_time=year(rs("zl_news_time"))
	
	wjm=zl_news_id+zl_news_time'文件名等于id+年份
	Zx=Zx&"<li><p><a href="&chr(34)&"/"&zl_news_lei_name&"/"&wjm&".html"&chr(34)&">"&bt&"</a></p></li>"
	Rs.Movenext
	Loop
	Zxw="<li><h2>最新文章</h2></li>"	
	Z=Zxw+Zx'最新文档栏目	
	Rs.close
	Set Rs=Nothing
	
	
    ''''''''''''''''''''''''''''''''''''''''''友情链接''''''''''''''''''''''''''''''''''''''''''''''	
	
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
	link="<li><h2>友情链接</h2></li>"	
	Y=link+links'友情链接		
	Rs.close
	Set Rs=Nothing
	
	
    ''''''''''''''''''''''''''''''''''''''''''首页列表显示''''''''''''''''''''''''''''''''''''''''''''''
    ziding=0
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_content where zl_news_lei_id=2 order by zl_news_zhiding,zl_news_id desc"'按置顶排序
	Rs.open Sql,conn,1,1
	Do while not rs.eof and ziding<liebiao
	ziding=ziding+1
	if flag=0 then'文列表显示方式
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
	
	
			Set Rsp=server.createobject("ADODB.Recordset")'统计评论数
			Sqlp="select * from zl_message where m_nid="&zl_news_id
			Rsp.open Sqlp,conn,1,3
			prcs=rsp.recordcount
			rsp.close
	
			mystr=rs("zl_news_key")'关链词折分
			a=Split(mystr,",")
			For Each Strs In a		
			key=" <a href="&chr(34)&"/"&zl_news_lei_name&"/"&wjmm&".html"&chr(34)&" title="&zl_news_title&">"&Strs&"</a> "	
			next
			
	index_1=index_1&"<li><strong><a href="&chr(34)&"/"&zl_news_lei_name&"/"&wjmm&".html"&chr(34)&">"&zl_news_title&"</a></strong></li><li><span>"&zl_news_nr&"</span></li> <li><font>发布于 "&keys&" <a href="&chr(34)&"/"&zl_news_lei_name&"/"&wjmm&".html"&chr(34)&">查看</a>("&hits&") <a href="&chr(34)&"/"&zl_news_lei_name&"/"&wjmm&".html"&chr(34)&">评论</a>("&prcs&") <span style="&chr(34)&"cursor:pointer"&chr(34)&" onClick="&chr(34)&"window.external.addFavorite('http://"&URl&"/"&zl_news_lei_name&"/"&wjmm&".html','"&zl_news_title&"')"&chr(34)&" title="&chr(34)&""&zl_news_title&""&chr(34)&">收藏</span> | <input type="&chr(34)&"text"&chr(34)&" class="&chr(34)&"inputys"&chr(34)&" value="&chr(34)&"http://"&URl&"/"&zl_news_lei_name&"/"&wjmm&".html"&chr(34)&" id="&chr(34)&""&zl_news_id&""&chr(34)&" readonly="&chr(34)&"readonly"&chr(34)&"> <a href="&chr(34)&"#"&chr(34)&" onmousedown="&chr(34)&"copyurl1("&zl_news_id&")"&chr(34)&">分享</a> | 分类:<a href="&chr(34)&"/"&zl_news_lei_name&"/"&zl_news_lei_name&"1.html"&chr(34)&">"&bb_nn&"</a> 关链词:"&key&"</font></li><li><hr></li>"	        
			Rs.movenext
			Loop
			
	Rs.close
	Set Rs=Nothing
	
''''''''''''''''''''''''''''''''''''''''''导航栏目生成调用数据''''''''''''''''''''''''''''''''''''''''''''''

	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_small_class where zl_s_id="&cint(request("zl_s_id"))
	Rs.open Sql,conn,1,1
    zl_s_s_name=rs("zl_s_s_name")'文件链接地址
	zl_s_key=rs("zl_s_key")'栏目关链词
	zl_s_des=rs("zl_s_des")'栏目描述
	zl_adress=rs("zl_adress")'模版链接地址
	zl_s_name=rs("zl_s_name")'导航名称
	zl_s_bt=rs("zl_s_bt")
	rs.close
	Set rs=nothing
	
%>