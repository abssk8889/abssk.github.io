		<%
			Server.ScriptTimeOut=50000
				
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'  版式权所有:杭州成楼网络科技有限公司                 '
			'  联系电话：0571-86980893                             '
			'  联系人：子 氵良                                     '
			'  联系QQ：41864438                                    '
			'  msn: zjxyk@hotmail.com                              '
			'  网址：http://aliang.zjxyk.com/                      '
			'                                                      '
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Sub Clkj_asptohtml(c_id,Html_m,Html_FileName)
			
			Dim Template,TempFile,fso,RFile,TempPath,rsInfo,rsTemp,pic,p_name
			Set fso=Server.CreateObject("Scripting.FileSystemObject")
			Set TempFile= fso.OpenTextFile(Server.MapPath("../../ZiLiang_moban/ziliang_view.html"))
			Template=TempFile.ReadAll
			TempFile.Close
			Set fso=Nothing

			'''''''''''''''''''''''''下，上''''''''''''''''''''''''''''''
			exec="select * from zl_content where zl_news_id="&c_id-1
			set rs=server.createobject("adodb.recordset")
			rs.open exec,conn,1,1
			if not rs.Eof then
			y3=year(rs("zl_news_time"))
			id3=rs("zl_news_id")
			z3=y3+id3
			lei3=c(rs("zl_news_lei_name"))
				if z3<>0 then
				shang="<a href="&chr(34)&"/"&lei3&"/"&z3&".html"&chr(34)&">上一篇</a>"
				else shang="没有了"
				end if
			else
			shang="<a href='/' >返回首页</a>"
			end if
			
			'===========================读出下一id======================================
				  
			exec="select * from zl_content where zl_news_id="&c_id+1
			set rs=server.createobject("adodb.recordset")
			rs.open exec,conn,1,1
			if not rs.bof then
			y2=year(rs("zl_news_time"))
			id2=rs("zl_news_id")
			z2=y2+id2
			lei2=c(rs("zl_news_lei_name"))
			xia="<a href="&chr(34)&"/"&lei2&"/"&z2&".html"&chr(34)&">下一篇</a>"
			else
			xia="<a href='/' >返回首页</a>"
			end if
				

        ''''''''''''''''''''''''''''''''''详细信息''''''''''''''''''''''''''''''''''''''''''''
			
			Set Rs=server.createobject("ADODB.Recordset")
			Sql="select * from zl_content where zl_news_id="&c_id
			Rs.open Sql,conn,1,1
			
			v_id=rs("zl_news_id")
			v_name=rs("zl_news_lei_name")'分类名称
			v_name_n=c(rs("zl_news_lei_name"))'中文找拼音
			v_title=rs("zl_news_title")'标题
			v_key=rs("zl_news_key")'关链词
			v_maiosu=rs("zl_news_maiosu")'描述
			v_nr=rs("zl_news_nr")'内容
			v_time=rs("zl_news_time")'时间
			v_wjm=v_id+year(rs("zl_news_time"))
			
			    mystrr=v_key
				aa=Split(mystrr,",")
				For Each Strss In aa
				keyy=keyy&" <a href="&chr(34)&"/"&v_name_n&"/"&v_wjm&".html"&chr(34)&" alt="&v_title&">"&Strss&"</a> "	
				next
'''''''''''内容显示'''''''''''''''''''''''''''''
v_lei="<a href="&chr(34)&"/"&zl_news_lei_name&"/"&zl_news_lei_name&"1.html"&chr(34)&">"&bb_nn&"</a>"'分类				
v_view="<li><strong>"&v_title&"</strong></li><li><font> "&v_time&" "&shang&" / "&xia&" | 分类:"&v_lei&" | <a href="&chr(34)&"/"&zl_news_lei_name&"/"&v_wjm&".html"&chr(34)&">查看</a>(<script src="&chr(34)&"/ZiLiang_inc/cont.asp?id="&v_id&"&ck=ck"&chr(34)&"></script>) | <a href="&chr(34)&"#l_y"&chr(34)&">评论</a>(<script src="&chr(34)&"/ZiLiang_inc/cont.asp?id="&v_id&"&pr=pr"&chr(34)&"></script>) | <span style="&chr(34)&"cursor:pointer"&chr(34)&" onClick="&chr(34)&"window.external.addFavorite('http://"&URl&"/"&v_name_n&"/"&v_wjm&".html','"&v_title&"')"&chr(34)&" title="&chr(34)&""&v_title&""&chr(34)&">收藏</span> <input type="&chr(34)&"text"&chr(34)&" class="&chr(34)&"inputys"&chr(34)&" value="&chr(34)&"http://"&URl&"/"&v_name_n&"/"&v_wjm&".html"&chr(34)&" id="&chr(34)&""&v_id&""&chr(34)&" readonly="&chr(34)&"readonly"&chr(34)&"> | <a href="&chr(34)&"#"&chr(34)&" onmousedown="&chr(34)&"copyurl1("&v_id&")"&chr(34)&">分享</a></font></li><li><span>"&v_nr&"</span></li><li><font>标签："&keyy&" | <span style="&chr(34)&"cursor:pointer"&chr(34)&" onClick="&chr(34)&"window.external.addFavorite('http://"&URl&"/"&v_name_n&"/"&v_wjm&".html','"&v_title&"')"&chr(34)&" title="&chr(34)&""&v_title&""&chr(34)&">收藏</span> <input type="&chr(34)&"text"&chr(34)&" class="&chr(34)&"inputys"&chr(34)&" value="&chr(34)&"http://"&URl&"/"&v_name_n&"/"&v_wjm&".html"&chr(34)&" id="&chr(34)&""&v_id&""&chr(34)&" readonly="&chr(34)&"readonly"&chr(34)&"> | <a href="&chr(34)&"#"&chr(34)&" onmousedown="&chr(34)&"copyurl1("&v_id&")"&chr(34)&">分享</a> | "&shang&" / "&xia&"</font></li>"


	'''''''' 留言内容''''''''''''''''''''''			
				
prr="<form name="&chr(34)&"msg"&chr(34)&" action="&chr(34)&"/ZiLiang_inc/inc.asp?Class=Message"&chr(34)&" method="&chr(34)&"post"&chr(34)&" onsubmit="&chr(34)&"return tj()"&chr(34)&"><li id="&chr(34)&"l_y"&chr(34)&"><font>请谨慎发帖，本网站会记录您的IP地址。请注意，根据我国法律，网站会将有关您发帖内容、发帖时间以及您发帖时的IP地址的记录保留至少60天，并且只要接到合法请求，即会将这类信息提供给有关政府机构</font></li><li><p>内容:<textarea name="&chr(34)&"m_nr"&chr(34)&" cols="&chr(34)&"45"&chr(34)&" rows="&chr(34)&"6"&chr(34)&" id="&chr(34)&"m_nr"&chr(34)&"></textarea></p></li><li><p>昵称:<input name="&chr(34)&"m_na"&chr(34)&" type="&chr(34)&"text"&chr(34)&" id="&chr(34)&"m_na"&chr(34)&" size="&chr(34)&"18"&chr(34)&" /><input name="&chr(34)&"m_nid"&chr(34)&" type="&chr(34)&"hidden"&chr(34)&" id="&chr(34)&"m_nid"&chr(34)&" value="&chr(34)&""&v_id&""&chr(34)&"/> 验证码:<input name="&chr(34)&"num1"&chr(34)&" id="&chr(34)&"num1"&chr(34)&" size="&chr(34)&"1"&chr(34)&" style="&chr(34)&" border-color:#FFFFFF; background-color:#FFFFFF; border-left:1px dashed #cccccc;border-top:1px dashed #cccccc;border-bottom:1px dashed #cccccc; border-right:1px dashed #ffffff;"&chr(34)&"><input name="&chr(34)&"num2"&chr(34)&" id="&chr(34)&"num2"&chr(34)&" size="&chr(34)&"1"&chr(34)&" style="&chr(34)&" border-color:#FFFFFF; background-color:#FFFFFF; border:1px dashed #cccccc; color:#FF0000;"&chr(34)&"> <input name="&chr(34)&"num3"&chr(34)&" id="&chr(34)&"num3"&chr(34)&" size="&chr(34)&"3"&chr(34)&" style="&chr(34)&" border-color:#FFFFFF; background-color:#FFFFFF; border:1px dashed #cccccc; background-color:#ffff; color:#FF9900;"&chr(34)&"></li><li><input type="&chr(34)&"submit"&chr(34)&" name="&chr(34)&"Submit"&chr(34)&" value="&chr(34)&"提交留言"&chr(34)&" /></p></li></form>"

            prhf="<script src="&chr(34)&"/ZiLiang_inc/cont.asp?id="&v_id&"&prl=prl"&chr(34)&"></script>"

			Html_Code=Template
			'===========================内容表中的相关内容======================================
			Html_Code=Replace(Html_Code,"$ziliang-title$",web_bt)'标题
			Html_Code=Replace(Html_Code,"$ziliang-nav$",Nav)'导航
			Html_Code=Replace(Html_Code,"$ziliang-zhanzhang$",zz)'站长	
			Html_Code=Replace(Html_Code,"$lm$",L)			
			if wz=1 then
			Html_Code=Replace(Html_Code,"$wz$",Z)'最新文档
			else 
			Html_Code=Replace(Html_Code,"$wz$",off)
			end if
			
			if pr=1 then
			Html_Code=Replace(Html_Code,"$pr$",prr)
			else
			Html_Code=Replace(Html_Code,"$pr$",off)'评论开启
			end if
			
			Html_Code=Replace(Html_Code,"$cd$",off)
			
			if linkss=1 then
			Html_Code=Replace(Html_Code,"$yq$",Y)
			else
			Html_Code=Replace(Html_Code,"$yq$",off)'友情链接
			end if
						
			Html_Code=Replace(Html_Code,"$ziliang-miaosu$",miaosu)'描述
			Html_Code=Replace(Html_Code,"$middle$",v_view)'内容表	  
			Html_Code=Replace(Html_Code,"$banquan$",banquan)'版权
			Html_Code=Replace(Html_Code,"$ziliang-logo$",logo)'图标				  
	        Html_Code=Replace(Html_Code,"$v_title$",v_title)'内容标题
			Html_Code=Replace(Html_Code,"$v_maiosu$",v_maiosu)'内容描述
			Html_Code=Replace(Html_Code,"$v_key$",v_key)'内容关键词
			Html_Code=Replace(Html_Code,"$v_lei$",v_lei)'分类
			Html_Code=Replace(Html_Code,"$prhf$",prhf)'回复内容
			'Html_Code=Replace(Html_Code,"$prhf$",prhf)
			'判断文件目录是否存在如果不存在，生成,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
			
			Set fso=Server.CreateObject("Scripting.FileSystemObject")	
			If Not fso.FolderExists(Server.MapPath("/"&Html_m))Then
			fso.CreateFolder(Server.MapPath("/"&Html_m))
			End If
				
			Set FOut = fso.CreateTextFile(Server.MapPath("/"&Html_m&"/"&Html_FileName))
			FOut.WriteLine Html_Code
			FOut.Close
			
			set FOut = nothing
			Set fso = Nothing
				
			End Sub	
	%>