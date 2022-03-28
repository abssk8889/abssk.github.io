		<!--#include file="../../ziliang_inc/inc.asp"-->
			<%
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'  版式权所有:杭州成楼网络科技有限公司                 '
			'  联系电话：0571-86980893                             '
			'  联系人：子 氵良:阿梁                                '
			'  联系QQ：41864438                                   '
			'  msn: zjxyk@hotmail.com                              '
			'  网址：http://aliang.zjxyk.com/                      '
			'                                                      '
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						
		Dim Template,TempFile,fso,RFile,TempPath,rsInfo,rsTemp,BigClass,SmallClass,ShowTime,RndNum,Mb_path,off
		off=""
		Mb_path="/ZiLiang_moban/ziliang_index.html"
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		Set TempFile= fso.OpenTextFile(Server.MapPath(Mb_path))
		Template=TempFile.ReadAll
		TempFile.Close
		Set fso=Nothing					
		'===========================开始替换模版内容======================================	
			
		Html_Code=Template
		'===========================内容表中的相关内容======================================
		Html_Code=Replace(Html_Code,"$ziliang-title$",web_bt)'标题
		Html_Code=Replace(Html_Code,"$ziliang-key$",ikey)'关键词
		Html_Code=Replace(Html_Code,"$ides$",ides)'描述
		Html_Code=Replace(Html_Code,"$ziliang-nav$",Nav)'导航
		Html_Code=Replace(Html_Code,"$ziliang-zhanzhang$",zz)'站长	
	    Html_Code=Replace(Html_Code,"$lm$",L)
		if wz=1 then
		Html_Code=Replace(Html_Code,"$wz$",Z)'最新文档
		else 
		Html_Code=Replace(Html_Code,"$wz$",off)
		end if
		Html_Code=Replace(Html_Code,"$pr$",off)
	    Html_Code=Replace(Html_Code,"$cd$",off)
		if linkss=1 then
		Html_Code=Replace(Html_Code,"$yq$",Y)
		else
		Html_Code=Replace(Html_Code,"$yq$",off)'友情链接
		end if
		'Html_Code=Replace(Html_Code,"$ziliang-navv$",zl_s_name)'栏目
		Html_Code=Replace(Html_Code,"$ziliang-miaosu$",miaosu)'描述
		Html_Code=Replace(Html_Code,"$middle$",index_1)'主页列表	  
	    Html_Code=Replace(Html_Code,"$banquan$",banquan)'版权
		Html_Code=Replace(Html_Code,"$ziliang-logo$",logo)'图标		  
			  
		'===========================判断文件目录是否存在如果不存在，生成=====================
		
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		If Not fso.FolderExists(Server.MapPath("/"))Then
		fso.CreateFolder(Server.MapPath("/"))
		End If
		
		'=========================== 按年文件 ======================================
		
		Set FOut = fso.CreateTextFile(Server.MapPath("/index.html"))
		FOut.WriteLine Html_Code
		FOut.Close
		set FOut = nothing
		Set fso = Nothing
		conn.close
		set conn=nothing
		response.write"<script language='javascript'>alert('成功生成，请返回');history.go(-1);</script>"
		%>