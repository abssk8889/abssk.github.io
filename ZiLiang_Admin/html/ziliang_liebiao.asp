<!--#include file="../../ziliang_inc/inc.asp"-->

<%
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'  ��ʽȨ����:���ݳ�¥����Ƽ����޹�˾                 '
	'  ��ϵ�绰��0571-86980893                             '
	'  ��ϵ�ˣ��� ����                                     '
	'  ��ϵQQ��41864438                                    '
	'  msn: zjxyk@hotmail.com                              '
	'  ��ַ��http://aliang.zjxyk.com/                      '
	'                                                      '
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Server.ScriptTimeOut=50000
lanid=request("zl_s_id")
%>
<%
'����FSOֱ�Ӷ�ȡģ��
function fsow(filename)
   on error resume next
   dim mfo,rtf,body
   set mfo=Server.CreateObject("Scripting.FileSystemObject")
   set rtf=mfo.OpenTextFile(server.mappath(filename),1)
   body=rtf.readall
   fsow=body
   set rtf=nothing
   set mfo=nothing
end function

function CreateFolder(Folder)
	Dim Fso, F
	Set Fso = Server.CreateObject("Scripting.FileSystemObject")
	If Fso.FolderExists(Server.MapPath(Folder)) Then Exit function
	Set F = Fso.CreateFolder(Server.MapPath(Folder))
	Set F = Nothing
	Set Fso = Nothing
    CreateFolder=Folder
End function


'��Ʒ��ҳ
function GetNews(ntop,nIndex,nMaxPerPage,npx)
  		set rs=server.createobject("adodb.recordset")
		exec="select * from zl_content where zl_news_Lei= "&lanid&" order by zl_news_paixu,zl_news_id desc"
		rs.open exec,conn,1,1
		do while not rs.eof	and num<clng(nIndex*nMaxPerPage)
		if num >= clng((nIndex-1)*nMaxPerPage) then
		
		if flag=0 then'���б���ʾ��ʽ
		zl_news_nr=left(RemoveHTML(rs("zl_news_nr")),150)
		else
		zl_news_nr=rs("zl_news_nr")
		end if
		zl_news_user=rs("zl_news_user")
		zl_news_title=rs("zl_news_title")
		zl_news_lei_name=rs("zl_news_lei_name")
		zl_news_id=rs("zl_news_id")
		keys=Cstr(rs("zl_news_time"))
		zl_news_time=year(rs("zl_news_time"))
		wjmm=zl_news_id+zl_news_time
		hits=rs("hits")
	   
				Set Rs_n=server.createobject("ADODB.Recordset")
				Sql_n="select * from zl_small_class where zl_s_id="&rs("zl_news_Lei")
				Rs_n.open Sql_n,conn,1,1
				bb_nc=rs_n("zl_s_name")
				Rs_n.close	
	
			mystr=rs("zl_news_key")'�������۷�
			a=Split(mystr,",")
			For Each Strs In a		
			key="<a href="&chr(34)&"/"&zl_s_s_name&"/"&wjmm&".html"&chr(34)&" title="&zl_news_title&">"&Strs&"</a> "	
			next
       
		tmplink=tmplink&"<li><strong><a href="&chr(34)&"/"&zl_s_s_name&"/"&wjmm&".html"&chr(34)&">"&zl_news_title&"</a></strong></li><li><span>"&zl_news_nr&"</span></li> <li><font>������ "&keys&" <a href="&chr(34)&"/"&zl_s_s_name&"/"&wjmm&".html"&chr(34)&">�鿴</a>("&hits&") <a href="&chr(34)&"/"&zl_s_s_name&"/"&wjmm&".html"&chr(34)&">����</a>("&prcs&") <span style="&chr(34)&"cursor:pointer"&chr(34)&" onClick="&chr(34)&"window.external.addFavorite('http://"&URl&"/"&zl_s_s_name&"/"&wjmm&".html','"&zl_news_title&"')"&chr(34)&" title="&chr(34)&""&zl_news_title&""&chr(34)&">�ղ�</span> | <input type="&chr(34)&"text"&chr(34)&" class="&chr(34)&"inputys"&chr(34)&" value="&chr(34)&"http://"&URl&"/"&zl_s_s_name&"/"&wjmm&".html"&chr(34)&" id="&chr(34)&""&zl_news_id&""&chr(34)&" readonly="&chr(34)&"readonly"&chr(34)&"> <a href="&chr(34)&"#"&chr(34)&" onmousedown="&chr(34)&"copyurl1("&zl_news_id&")"&chr(34)&">����</a> | ����:<a href="&chr(34)&"/"&zl_s_s_name&"/"&zl_s_s_name&"1.html"&chr(34)&">"&bb_nc&"</a> ������:"&key&"</font></li><li><hr></li>"
			
	  end if
      num=num+1
		rs.movenext
        loop
        rs.close
   set rs=nothing  
   GetNews=tmplink
end function
%>
<!---------------------------------------------------------------------------->
<%
dim moban_file  'ģ���ļ�

moban_file="/ZiLiang_moban/ziliang_liebiao.html"'ģ�����ӵ�ַ

startime=timer()
'������ҳ�ļ�Ŀ¼

strfoldname="/"&zl_s_s_name&"/" '·��
fileName_Name=zl_s_s_name'�ļ���

CreateFolder(strfoldname)  

'�ļ��еı���Ԫ��
dim NewsTitle,NewsPage

dim strfilename,strfilename2    '����Ŀ¼�ļ���
dim filecount      '�ļ���
filecount=0

'��ȡ�����ļ�����
dim okfile,dtime1,tmpfile,file1
okfile=fsow(moban_file)
  
   '��ƷĿ¼�б�
   dim nMaxPerPage,ntop,bHavedata,totalPut,nIndex,nPageCount,nLinkNum
   nIndex=1                         '��¼��ҳ��
   totalPut=0                       '�ܼ�¼��
   bHavedata=0                      '����Ƿ��������
   nMaxPerPage=liebiao          'ÿҳ��¼��
   nPageCount=1                     '��ҳ��
   nLinkNum=5                      'ÿҳ������ʾ������
   sql1="select * from zl_content where zl_news_Lei="&lanid
   set rs1= Server.CreateObject("ADODB.Recordset")
   rs1.open sql1,conn,1,1
   if not rs1.eof then
      totalPut=rs1.recordcount                    '�ܼ�¼��
      if totalPut mod nMaxperPage=0 then 
          nPageCount= totalPut \ nMaxperPage      '��ҳ��
      else 
          nPageCount= totalPut \ nMaxperPage+1    '��ҳ��
      end if 
           
      for nIndex=1 to nPageCount    'ѭ��д��ÿһҳ
         strfilename2=fileName_Name+cstr(nIndex)+".html"

         'ѭ����ȡ��Ʒ��¼������д���ļ�
         dim okfile1        
         okfile1=okfile
         NewsTitle=GetNews(0,nIndex,nMaxPerPage,0)          '����������¼��
         okfile1 = replace(okfile1,"$middle$",NewsTitle)       '�滻ģ�����
         
         'ҳ���ҳ������
         dim strLink,npv,mPages         '�����ַ���ҳ�油��������ҳ�ֶ���
         if nPageCount mod nLinkNum=0 then
            mPages=nPageCount\nlinkNum             '��ҳ�ֶ����������20����ҳ��ÿҳ��10�����ӣ���ô����20/10���ֶ�
         else
            mPages=nPageCount\nlinkNum+1
         end if
         npv=nLinkNum\2                            '����ķ�ҳ�磺��1ҳ1,2...10����2ҳ11,12...20����3ҳ21,22...30��
                                                   '�������  ����1ҳ1,2...10����2ҳ5,6,7...15����3ҳ10,11...20��                
         dim strLinkFile,strLastLinkFile,strNextLinkFile,strFirstLinkFile,strEndLinkFile
 
         if nIndex>1 then   '��һҳ����ҳ
            strLastLinkFile=fileName_Name+cstr(nIndex-1)+".html"
            strFirstLinkFile=fileName_Name+"1.html"
         else
            strLastLinkFile=fileName_Name+"1.html"
            strFirstLinkFile=fileName_Name+"1.html"
         end if
         
         if nIndex<nPageCount then   '��һҳ��ĩҳ
            strNextLinkFile=fileName_Name+cstr(nIndex+1)+".html"
            strEndLinkFile=fileName_Name+cstr(nPageCount)+".html"
         else
            strNextLinkFile=fileName_Name+cstr(nPageCount)+".html"
            strEndLinkFile=fileName_Name+cstr(nPageCount)+".html"
         end if
                  
         '��ҳ��С��ÿҳ����ʾ��������
         strLink=""
         if nPageCount<=nLinkNum then
            if nIndex>1 then
               strLink=strLink+"<a href="+strLastLinkFile+">[Prev]</a>&nbsp;"
            end if
            for i=1 to nPageCount
               strLinkFile=fileName_Name+cstr(i)+".html"
               if i=nIndex then
                  strLink=strLink+"<b>"+cstr(i)+"</b>&nbsp;"          '��ǰҳ����������
               else
                  strLink=strLink+"<a href="+strLinkFile+">"                  
                  strLink=strLink+"["+cstr(i)+"]</a>"+"&nbsp;"
               end if
            next
            if nIndex<nPageCount then strLink=strLink+"&nbsp;<a href="+strNextLinkFile+">[Next]</a>"
         end if

         '��ҳ������ÿҳ��ʾ���Ӹ���
         if nPageCount>nLinkNum then
            dim j
            if nPageCount mod nLinkNum=0 then
               mPages=nPageCount\nlinkNum       '�����ҳ����
            else
               mPages=nPageCount\nlinkNum+1
            end if
            npv=nLinkNum\2
            
            for j=1 to mPages*2                 '*2��Ϊ���ֲ�npv  
               if nIndex>((j-1)*nLinkNum-((j-1)*npv)) and nIndex<=(j*nLinkNum-((j-1)*npv)) then      'ѡ��ֶ�ҳ��д�룬��������ر���Ҫ

                  strLink=""
                  if j=1 then'��һ��                     
                     if nIndex>1 then  
                        strLink=strLink+"<a href="+strFirstLinkFile+">[First]</a>&nbsp;"    
                        strLink=strLink+"<a href="+strLastLinkFile+">[Prev]</a>&nbsp;"
                     end if
                     for i=1 to nLinkNum
                        strLinkFile=fileName_Name+cstr(i)+".html"
                        if i=nIndex then
                           strLink=strLink+"<b>"+cstr(i)+"</b>&nbsp;"     '��ǰҳ����������
                        else
                           strLink=strLink+"<a href="+strLinkFile+">"
                           strLink=strLink+"["+cstr(i)+"]</a>"+"&nbsp;"
                        end if
                     next
                     if nIndex<nPageCount then
                        strLink=strLink+"&nbsp;<a href="+strNextLinkFile+">[Next]</a>"
                        strLink=strLink+"&nbsp;<a href="+strEndLinkFile+">[Last]</a>"
                     end if
                  else                                                    '��2�ο�ʼ
                     strLink=strLink+"<a href="+strFirstLinkFile+">[First]</a>&nbsp;"
                     strLink=strLink+"<a href="+strLastLinkFile+">[Prev]</a>&nbsp;"
                     for i=npv*(j-1) to npv*(j-1)+nlinkNum
                        if i> nPageCount then exit for                '���Ƴ���ҳ��������һ�����Ҫ
                        strLinkFile=fileName_Name+cstr(i)+".html"
                        if i=nIndex then
                           strLink=strLink+"<b>"+cstr(i)+"</b>&nbsp;"     '��ǰҳ����������
                        else
                           strLink=strLink+"<a href="+strLinkFile+">"
                           strLink=strLink+"["+cstr(i)+"]</a>"+"&nbsp;"
                        end if
                     next 
                     if nIndex<nPageCount then
                        strLink=strLink+"&nbsp;<a href="+strNextLinkFile+">[Next]</a>"   
                        strLink=strLink+"&nbsp;<a href="+strEndLinkFile+">[Last]</a>"
                     end if    
                  end if
               end if
            next         
         end if

  v_lei="<a href="&chr(34)&"/"&zl_s_s_name&"/"&zl_s_s_name&"1.html"&chr(34)&">"&zl_s_name&"</a>"
         '��ҳ�����Ӳ���д���ļ�
         NewsPage=strLink		 
			okfile1 = replace(okfile1,"$fenye$",NewsPage)			
			okfile1=Replace(okfile1,"$ziliang-title$",web_bt)'����
			okfile1=Replace(okfile1,"$ziliang-nav$",Nav)'����
			okfile1=Replace(okfile1,"$ziliang-zhanzhang$",zz)'վ��	
			okfile1=Replace(okfile1,"$lm$",L)
			if wz=1 then
			okfile1=Replace(okfile1,"$wz$",Z)'�����ĵ�
			else 
			okfile1=Replace(okfile1,"$wz$",off)
			end if
			okfile1=Replace(okfile1,"$pr$",off)
			okfile1=Replace(okfile1,"$cd$",off)
			if linkss=1 then
			okfile1=Replace(okfile1,"$yq$",Y)
			else
			okfile1=Replace(okfile1,"$yq$",off)'��������
			end if
			okfile1=Replace(okfile1,"$ziliang-navv$",v_lei)'��Ŀ
			okfile1=Replace(okfile1,"$ziliang-miaosu$",miaosu)'����	  
			okfile1=Replace(okfile1,"$banquan$",banquan)'��Ȩ
			okfile1=Replace(okfile1,"$ziliang-logo$",logo)'ͼ��	
			okfile1=Replace(okfile1,"$v_title$",zl_s_bt)'���ݱ���
			okfile1=Replace(okfile1,"$v_maiosu$",zl_s_des)'��������
			okfile1=Replace(okfile1,"$v_key$",zl_s_key)'���ݹؼ���
			okfile1=Replace(okfile1,"$v_lei$",v_lei)'�������
         tmpfile=strfoldname+strfilename2
         file1=server.mappath(tmpfile)
         '�½�ҳ���ļ�
         Set fso = Server.CreateObject("Scripting.FileSystemObject")
         Set fout = fso.Createtextfile(file1,true)
         fout.writeline okfile1
         fout.close
         set fout=nothing
         set fso=nothing
         filecount=filecount+1
%>
<a href="<%=strfoldname+strfilename2%>" target="_blank"><%=cstr(filecount)+"��"+strfilename2%>&nbsp;&nbsp;�鿴>>></a> <a href="javascript:history.back(1)">��������������Ŀ</a><br>

  <%
		next 
		bHavedata=1
		end if
		rs1.close
		set rs1=nothing		
		endtime=timer()
		response.write ("������"+cstr(filecount)+"���ļ���ҳ��ִ��ʱ��:"&FormatNumber(endtime-startime,3)&"��")
		conn.close
		set conn=nothing	
  %>