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
	%>