		<!--#include file="../../ziliang_inc/inc.asp"-->
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
			Dim Template,TempFile,fso,RFile,TempPath,rsInfo,rsTemp,pic,p_name
			Set fso=Server.CreateObject("Scripting.FileSystemObject")
			Set TempFile= fso.OpenTextFile(Server.MapPath("/ZiLiang_moban/ziliang_nav.html"))
			Template=TempFile.ReadAll
			TempFile.Close
			Set fso=Nothing					
		
          ''''''''''''''''''''''''''''''''''详细信息''''''''''''''''''''''''''''''''''''''''''''
					
			Set Rs=server.createobject("ADODB.Recordset")
			Sql="select * from zl_content where zl_news_Lei="&request("zl_s_id")
			Rs.open Sql,conn,1,1
			if rs.eof and rs.bof then
			response.write "<script language='javascript'>alert('暂无数据可生成，返回添加数据');history.go(-1);</script>"
			else
			v_title=rs("zl_news_title")'标题
			v_key=rs("zl_news_key")'关链词
			v_maiosu=rs("zl_news_maiosu")'描述
			v_nr=rs("zl_news_nr")'内容
			v_wjm=v_id+year(rs("zl_news_time"))
			v_lei="<a href="&chr(34)&"/"&c(rs("zl_news_lei_name"))&".html"&chr(34)&">"&rs("zl_news_lei_name")&"</a>"
			v_scm=c(rs("zl_news_lei_name"))&".html"
			
			
           '''''''''''内容显示'''''''''''''''''''''''''''''
		   
            nav_nr="<li><strong>"&v_title&"</strong></li><li><span>"&v_nr&"</span></li>"				

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
			
			Html_Code=Replace(Html_Code,"$ziliang-navv$",zl_s_name)'栏目
			Html_Code=Replace(Html_Code,"$ziliang-miaosu$",miaosu)'描述
			Html_Code=Replace(Html_Code,"$middle$",nav_nr)'内容	  
			Html_Code=Replace(Html_Code,"$banquan$",banquan)'版权
			Html_Code=Replace(Html_Code,"$ziliang-logo$",logo)'图标				  
	        Html_Code=Replace(Html_Code,"$v_title$",v_title)'内容标题
			Html_Code=Replace(Html_Code,"$v_maiosu$",v_maiosu)'内容描述
			Html_Code=Replace(Html_Code,"$v_key$",v_key)'内容关键词
			Html_Code=Replace(Html_Code,"$v_lei$",v_lei)'分类
			
			'判断文件目录是否存在如果不存在，生成,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
			
			Set fso=Server.CreateObject("Scripting.FileSystemObject")	
			If Not fso.FolderExists(Server.MapPath("/"))Then
			fso.CreateFolder(Server.MapPath("/"))
			End If
				
			Set FOut = fso.CreateTextFile(Server.MapPath("/"&v_scm))
			FOut.WriteLine Html_Code
			FOut.Close
			
			set FOut = nothing
			Set fso = Nothing
			response.write "<script language='javascript'>alert('成功生成');history.go(-1);</script>"
			end if
			
	%>