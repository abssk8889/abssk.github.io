		<!--#include file="../../ziliang_inc/inc.asp"-->
		<%
			Server.ScriptTimeOut=50000
				
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'  ��ʽȨ����:���ݳ�¥����Ƽ����޹�˾                 '
			'  ��ϵ�绰��0571-86980893                             '
			'  ��ϵ�ˣ��� ����                                     '
			'  ��ϵQQ��41864438                                    '
			'  msn: zjxyk@hotmail.com                              '
			'  ��ַ��http://aliang.zjxyk.com/                      '
			'                                                      '
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''			
			Dim Template,TempFile,fso,RFile,TempPath,rsInfo,rsTemp,pic,p_name
			Set fso=Server.CreateObject("Scripting.FileSystemObject")
			Set TempFile= fso.OpenTextFile(Server.MapPath("/ZiLiang_moban/ziliang_nav.html"))
			Template=TempFile.ReadAll
			TempFile.Close
			Set fso=Nothing					
		
          ''''''''''''''''''''''''''''''''''��ϸ��Ϣ''''''''''''''''''''''''''''''''''''''''''''
					
			Set Rs=server.createobject("ADODB.Recordset")
			Sql="select * from zl_content where zl_news_Lei="&request("zl_s_id")
			Rs.open Sql,conn,1,1
			if rs.eof and rs.bof then
			response.write "<script language='javascript'>alert('�������ݿ����ɣ������������');history.go(-1);</script>"
			else
			v_title=rs("zl_news_title")'����
			v_key=rs("zl_news_key")'������
			v_maiosu=rs("zl_news_maiosu")'����
			v_nr=rs("zl_news_nr")'����
			v_wjm=v_id+year(rs("zl_news_time"))
			v_lei="<a href="&chr(34)&"/"&c(rs("zl_news_lei_name"))&".html"&chr(34)&">"&rs("zl_news_lei_name")&"</a>"
			v_scm=c(rs("zl_news_lei_name"))&".html"
			
			
           '''''''''''������ʾ'''''''''''''''''''''''''''''
		   
            nav_nr="<li><strong>"&v_title&"</strong></li><li><span>"&v_nr&"</span></li>"				

			Html_Code=Template
			'===========================���ݱ��е��������======================================
			Html_Code=Replace(Html_Code,"$ziliang-title$",web_bt)'����
			Html_Code=Replace(Html_Code,"$ziliang-nav$",Nav)'����
			Html_Code=Replace(Html_Code,"$ziliang-zhanzhang$",zz)'վ��
			Html_Code=Replace(Html_Code,"$lm$",L)
			
			if wz=1 then
			Html_Code=Replace(Html_Code,"$wz$",Z)'�����ĵ�
			else 
			Html_Code=Replace(Html_Code,"$wz$",off)
			end if
			
			if pr=1 then
			Html_Code=Replace(Html_Code,"$pr$",prr)
			else
			Html_Code=Replace(Html_Code,"$pr$",off)'���ۿ���
			end if
			
			Html_Code=Replace(Html_Code,"$cd$",off)
			
			if linkss=1 then
			Html_Code=Replace(Html_Code,"$yq$",Y)
			else
			Html_Code=Replace(Html_Code,"$yq$",off)'��������
			end if
			
			Html_Code=Replace(Html_Code,"$ziliang-navv$",zl_s_name)'��Ŀ
			Html_Code=Replace(Html_Code,"$ziliang-miaosu$",miaosu)'����
			Html_Code=Replace(Html_Code,"$middle$",nav_nr)'����	  
			Html_Code=Replace(Html_Code,"$banquan$",banquan)'��Ȩ
			Html_Code=Replace(Html_Code,"$ziliang-logo$",logo)'ͼ��				  
	        Html_Code=Replace(Html_Code,"$v_title$",v_title)'���ݱ���
			Html_Code=Replace(Html_Code,"$v_maiosu$",v_maiosu)'��������
			Html_Code=Replace(Html_Code,"$v_key$",v_key)'���ݹؼ���
			Html_Code=Replace(Html_Code,"$v_lei$",v_lei)'����
			
			'�ж��ļ�Ŀ¼�Ƿ������������ڣ�����,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
			
			Set fso=Server.CreateObject("Scripting.FileSystemObject")	
			If Not fso.FolderExists(Server.MapPath("/"))Then
			fso.CreateFolder(Server.MapPath("/"))
			End If
				
			Set FOut = fso.CreateTextFile(Server.MapPath("/"&v_scm))
			FOut.WriteLine Html_Code
			FOut.Close
			
			set FOut = nothing
			Set fso = Nothing
			response.write "<script language='javascript'>alert('�ɹ�����');history.go(-1);</script>"
			end if
			
	%>