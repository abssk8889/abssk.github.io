		<!--#include file="../../ziliang_inc/inc.asp"-->
			<%
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'  ��ʽȨ����:���ݳ�¥����Ƽ����޹�˾                 '
			'  ��ϵ�绰��0571-86980893                             '
			'  ��ϵ�ˣ��� ����:����                                '
			'  ��ϵQQ��41864438                                   '
			'  msn: zjxyk@hotmail.com                              '
			'  ��ַ��http://aliang.zjxyk.com/                      '
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
		'===========================��ʼ�滻ģ������======================================	
			
		Html_Code=Template
		'===========================���ݱ��е��������======================================
		Html_Code=Replace(Html_Code,"$ziliang-title$",web_bt)'����
		Html_Code=Replace(Html_Code,"$ziliang-key$",ikey)'�ؼ���
		Html_Code=Replace(Html_Code,"$ides$",ides)'����
		Html_Code=Replace(Html_Code,"$ziliang-nav$",Nav)'����
		Html_Code=Replace(Html_Code,"$ziliang-zhanzhang$",zz)'վ��	
	    Html_Code=Replace(Html_Code,"$lm$",L)
		if wz=1 then
		Html_Code=Replace(Html_Code,"$wz$",Z)'�����ĵ�
		else 
		Html_Code=Replace(Html_Code,"$wz$",off)
		end if
		Html_Code=Replace(Html_Code,"$pr$",off)
	    Html_Code=Replace(Html_Code,"$cd$",off)
		if linkss=1 then
		Html_Code=Replace(Html_Code,"$yq$",Y)
		else
		Html_Code=Replace(Html_Code,"$yq$",off)'��������
		end if
		'Html_Code=Replace(Html_Code,"$ziliang-navv$",zl_s_name)'��Ŀ
		Html_Code=Replace(Html_Code,"$ziliang-miaosu$",miaosu)'����
		Html_Code=Replace(Html_Code,"$middle$",index_1)'��ҳ�б�	  
	    Html_Code=Replace(Html_Code,"$banquan$",banquan)'��Ȩ
		Html_Code=Replace(Html_Code,"$ziliang-logo$",logo)'ͼ��		  
			  
		'===========================�ж��ļ�Ŀ¼�Ƿ������������ڣ�����=====================
		
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		If Not fso.FolderExists(Server.MapPath("/"))Then
		fso.CreateFolder(Server.MapPath("/"))
		End If
		
		'=========================== �����ļ� ======================================
		
		Set FOut = fso.CreateTextFile(Server.MapPath("/index.html"))
		FOut.WriteLine Html_Code
		FOut.Close
		set FOut = nothing
		Set fso = Nothing
		conn.close
		set conn=nothing
		response.write"<script language='javascript'>alert('�ɹ����ɣ��뷵��');history.go(-1);</script>"
		%>