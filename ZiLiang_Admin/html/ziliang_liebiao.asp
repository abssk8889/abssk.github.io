<!--#include file="../../ziliang_inc/inc.asp"-->

<%
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'  版式权所有:杭州成楼网络科技有限公司                 '
	'  联系电话：0571-86980893                             '
	'  联系人：子 氵良                                     '
	'  联系QQ：41864438                                    '
	'  msn: zjxyk@hotmail.com                              '
	'  网址：http://aliang.zjxyk.com/                      '
	'                                                      '
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Server.ScriptTimeOut=50000
lanid=request("zl_s_id")
%>
<%
'利用FSO直接读取模板
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


'产品分页
function GetNews(ntop,nIndex,nMaxPerPage,npx)
  		set rs=server.createobject("adodb.recordset")
		exec="select * from zl_content where zl_news_Lei= "&lanid&" order by zl_news_paixu,zl_news_id desc"
		rs.open exec,conn,1,1
		do while not rs.eof	and num<clng(nIndex*nMaxPerPage)
		if num >= clng((nIndex-1)*nMaxPerPage) then
		
		if flag=0 then'文列表显示方式
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
	
			mystr=rs("zl_news_key")'关链词折分
			a=Split(mystr,",")
			For Each Strs In a		
			key="<a href="&chr(34)&"/"&zl_s_s_name&"/"&wjmm&".html"&chr(34)&" title="&zl_news_title&">"&Strs&"</a> "	
			next
       
		tmplink=tmplink&"<li><strong><a href="&chr(34)&"/"&zl_s_s_name&"/"&wjmm&".html"&chr(34)&">"&zl_news_title&"</a></strong></li><li><span>"&zl_news_nr&"</span></li> <li><font>发布于 "&keys&" <a href="&chr(34)&"/"&zl_s_s_name&"/"&wjmm&".html"&chr(34)&">查看</a>("&hits&") <a href="&chr(34)&"/"&zl_s_s_name&"/"&wjmm&".html"&chr(34)&">评论</a>("&prcs&") <span style="&chr(34)&"cursor:pointer"&chr(34)&" onClick="&chr(34)&"window.external.addFavorite('http://"&URl&"/"&zl_s_s_name&"/"&wjmm&".html','"&zl_news_title&"')"&chr(34)&" title="&chr(34)&""&zl_news_title&""&chr(34)&">收藏</span> | <input type="&chr(34)&"text"&chr(34)&" class="&chr(34)&"inputys"&chr(34)&" value="&chr(34)&"http://"&URl&"/"&zl_s_s_name&"/"&wjmm&".html"&chr(34)&" id="&chr(34)&""&zl_news_id&""&chr(34)&" readonly="&chr(34)&"readonly"&chr(34)&"> <a href="&chr(34)&"#"&chr(34)&" onmousedown="&chr(34)&"copyurl1("&zl_news_id&")"&chr(34)&">分享</a> | 分类:<a href="&chr(34)&"/"&zl_s_s_name&"/"&zl_s_s_name&"1.html"&chr(34)&">"&bb_nc&"</a> 关链词:"&key&"</font></li><li><hr></li>"
			
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
dim moban_file  '模版文件

moban_file="/ZiLiang_moban/ziliang_liebiao.html"'模版链接地址

startime=timer()
'建立分页文件目录

strfoldname="/"&zl_s_s_name&"/" '路径
fileName_Name=zl_s_s_name'文件名

CreateFolder(strfoldname)  

'文件中的变量元素
dim NewsTitle,NewsPage

dim strfilename,strfilename2    '生成目录文件名
dim filecount      '文件数
filecount=0

'读取摸版文件内容
dim okfile,dtime1,tmpfile,file1
okfile=fsow(moban_file)
  
   '产品目录列表
   dim nMaxPerPage,ntop,bHavedata,totalPut,nIndex,nPageCount,nLinkNum
   nIndex=1                         '记录集页号
   totalPut=0                       '总记录数
   bHavedata=0                      '检测是否具有数据
   nMaxPerPage=liebiao          '每页记录数
   nPageCount=1                     '总页数
   nLinkNum=5                      '每页允许显示连接数
   sql1="select * from zl_content where zl_news_Lei="&lanid
   set rs1= Server.CreateObject("ADODB.Recordset")
   rs1.open sql1,conn,1,1
   if not rs1.eof then
      totalPut=rs1.recordcount                    '总记录数
      if totalPut mod nMaxperPage=0 then 
          nPageCount= totalPut \ nMaxperPage      '总页数
      else 
          nPageCount= totalPut \ nMaxperPage+1    '总页数
      end if 
           
      for nIndex=1 to nPageCount    '循环写入每一页
         strfilename2=fileName_Name+cstr(nIndex)+".html"

         '循环读取产品记录集，并写入文件
         dim okfile1        
         okfile1=okfile
         NewsTitle=GetNews(0,nIndex,nMaxPerPage,0)          '反向搜索记录集
         okfile1 = replace(okfile1,"$middle$",NewsTitle)       '替换模板变量
         
         '页面分页超连接
         dim strLink,npv,mPages         '连接字符，页面补差数，分页分段数
         if nPageCount mod nLinkNum=0 then
            mPages=nPageCount\nlinkNum             '分页分段数，如果有20个分页，每页有10个连接，那么就有20/10个分段
         else
            mPages=nPageCount\nlinkNum+1
         end if
         npv=nLinkNum\2                            '常规的分页如：第1页1,2...10；第2页11,12...20；第3页21,22...30；
                                                   '补差后变成  ：第1页1,2...10；第2页5,6,7...15；第3页10,11...20；                
         dim strLinkFile,strLastLinkFile,strNextLinkFile,strFirstLinkFile,strEndLinkFile
 
         if nIndex>1 then   '上一页，首页
            strLastLinkFile=fileName_Name+cstr(nIndex-1)+".html"
            strFirstLinkFile=fileName_Name+"1.html"
         else
            strLastLinkFile=fileName_Name+"1.html"
            strFirstLinkFile=fileName_Name+"1.html"
         end if
         
         if nIndex<nPageCount then   '下一页，末页
            strNextLinkFile=fileName_Name+cstr(nIndex+1)+".html"
            strEndLinkFile=fileName_Name+cstr(nPageCount)+".html"
         else
            strNextLinkFile=fileName_Name+cstr(nPageCount)+".html"
            strEndLinkFile=fileName_Name+cstr(nPageCount)+".html"
         end if
                  
         '分页数小于每页可显示的连接数
         strLink=""
         if nPageCount<=nLinkNum then
            if nIndex>1 then
               strLink=strLink+"<a href="+strLastLinkFile+">[Prev]</a>&nbsp;"
            end if
            for i=1 to nPageCount
               strLinkFile=fileName_Name+cstr(i)+".html"
               if i=nIndex then
                  strLink=strLink+"<b>"+cstr(i)+"</b>&nbsp;"          '当前页，无需连接
               else
                  strLink=strLink+"<a href="+strLinkFile+">"                  
                  strLink=strLink+"["+cstr(i)+"]</a>"+"&nbsp;"
               end if
            next
            if nIndex<nPageCount then strLink=strLink+"&nbsp;<a href="+strNextLinkFile+">[Next]</a>"
         end if

         '分页数大于每页显示连接个数
         if nPageCount>nLinkNum then
            dim j
            if nPageCount mod nLinkNum=0 then
               mPages=nPageCount\nlinkNum       '计算分页段数
            else
               mPages=nPageCount\nlinkNum+1
            end if
            npv=nLinkNum\2
            
            for j=1 to mPages*2                 '*2是为了弥补npv  
               if nIndex>((j-1)*nLinkNum-((j-1)*npv)) and nIndex<=(j*nLinkNum-((j-1)*npv)) then      '选择分段页面写入，这句描述特别重要

                  strLink=""
                  if j=1 then'第一段                     
                     if nIndex>1 then  
                        strLink=strLink+"<a href="+strFirstLinkFile+">[First]</a>&nbsp;"    
                        strLink=strLink+"<a href="+strLastLinkFile+">[Prev]</a>&nbsp;"
                     end if
                     for i=1 to nLinkNum
                        strLinkFile=fileName_Name+cstr(i)+".html"
                        if i=nIndex then
                           strLink=strLink+"<b>"+cstr(i)+"</b>&nbsp;"     '当前页，无需连接
                        else
                           strLink=strLink+"<a href="+strLinkFile+">"
                           strLink=strLink+"["+cstr(i)+"]</a>"+"&nbsp;"
                        end if
                     next
                     if nIndex<nPageCount then
                        strLink=strLink+"&nbsp;<a href="+strNextLinkFile+">[Next]</a>"
                        strLink=strLink+"&nbsp;<a href="+strEndLinkFile+">[Last]</a>"
                     end if
                  else                                                    '第2段开始
                     strLink=strLink+"<a href="+strFirstLinkFile+">[First]</a>&nbsp;"
                     strLink=strLink+"<a href="+strLastLinkFile+">[Prev]</a>&nbsp;"
                     for i=npv*(j-1) to npv*(j-1)+nlinkNum
                        if i> nPageCount then exit for                '限制超出页面数，这一句很重要
                        strLinkFile=fileName_Name+cstr(i)+".html"
                        if i=nIndex then
                           strLink=strLink+"<b>"+cstr(i)+"</b>&nbsp;"     '当前页，无需连接
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
         '将页面连接参数写入文件
         NewsPage=strLink		 
			okfile1 = replace(okfile1,"$fenye$",NewsPage)			
			okfile1=Replace(okfile1,"$ziliang-title$",web_bt)'标题
			okfile1=Replace(okfile1,"$ziliang-nav$",Nav)'导航
			okfile1=Replace(okfile1,"$ziliang-zhanzhang$",zz)'站长	
			okfile1=Replace(okfile1,"$lm$",L)
			if wz=1 then
			okfile1=Replace(okfile1,"$wz$",Z)'最新文档
			else 
			okfile1=Replace(okfile1,"$wz$",off)
			end if
			okfile1=Replace(okfile1,"$pr$",off)
			okfile1=Replace(okfile1,"$cd$",off)
			if linkss=1 then
			okfile1=Replace(okfile1,"$yq$",Y)
			else
			okfile1=Replace(okfile1,"$yq$",off)'友情链接
			end if
			okfile1=Replace(okfile1,"$ziliang-navv$",v_lei)'栏目
			okfile1=Replace(okfile1,"$ziliang-miaosu$",miaosu)'描述	  
			okfile1=Replace(okfile1,"$banquan$",banquan)'版权
			okfile1=Replace(okfile1,"$ziliang-logo$",logo)'图标	
			okfile1=Replace(okfile1,"$v_title$",zl_s_bt)'内容标题
			okfile1=Replace(okfile1,"$v_maiosu$",zl_s_des)'内容描述
			okfile1=Replace(okfile1,"$v_key$",zl_s_key)'内容关键词
			okfile1=Replace(okfile1,"$v_lei$",v_lei)'分类标题
         tmpfile=strfoldname+strfilename2
         file1=server.mappath(tmpfile)
         '新建页面文件
         Set fso = Server.CreateObject("Scripting.FileSystemObject")
         Set fout = fso.Createtextfile(file1,true)
         fout.writeline okfile1
         fout.close
         set fout=nothing
         set fso=nothing
         filecount=filecount+1
%>
<a href="<%=strfoldname+strfilename2%>" target="_blank"><%=cstr(filecount)+"、"+strfilename2%>&nbsp;&nbsp;查看>>></a> <a href="javascript:history.back(1)">返回生成其它栏目</a><br>

  <%
		next 
		bHavedata=1
		end if
		rs1.close
		set rs1=nothing		
		endtime=timer()
		response.write ("共生成"+cstr(filecount)+"个文件，页面执行时间:"&FormatNumber(endtime-startime,3)&"秒")
		conn.close
		set conn=nothing	
  %>