<!--#include file="../ziliang_inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="backdoor.css" type="text/css" rel="stylesheet">
<title>��������</title>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td height="39" class="tr_bg"><strong>&nbsp;��������</strong> <font color="#ff0000"><%=request.querystring("Lei")%></font></td>
  </tr>
</table>
<table width="100%" cellpadding="2" cellspacing="0">
<form action="?Class=config" method="post" name="myform">
<%
	Set Rs=server.createobject("ADODB.Recordset")
	Sql="select * from zl_config where config_id=1"
	Rs.open Sql,conn,1,1
	%>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">վ���ǳ�</td> 
    <td height="30"><input name="zl_web_zhanzhang" type="text" id="name" value="<%=rs("zl_web_zhanzhang")%>" size="10" /></td>
  </tr>
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">��վ����</td> 
    <td height="30"><input name="name" type="text" id="name" value="<%=rs("zl_web_title")%>" size="40" /></td>
  </tr>
    <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">��ҳ�ؼ���</td> 
    <td height="30"><input name="ikey" type="text" id="ikey" value="<%=rs("zl_web_ikey")%>" size="40" /></td>
  </tr>
   <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">��ҳ����</td> 
    <td height="30"><textarea name="ides" cols="50" rows="3" id="ides"><%=rs("zl_web_ides")%></textarea></td>
  </tr>
   <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">����˵��</td> 
    <td height="30"><textarea name="maiosu" cols="50" rows="4" id="maiosu"><%=rs("zl_web_miaosu")%></textarea></td>
  </tr>
   <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">��վ��ַ</td> 
    <td height="30"><input name="wangzhi" type="text" id="wangzhi" value="<%=rs("zl_web_wangzi")%>" size="30" /></td>
  </tr>
 
  <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9"><p>��վ��Ȩ</p>
    </td> 
    <td height="30"><textarea name="banquan" cols="60" rows="4" id="banquan"><%=rs("zl_web_banquan")%></textarea></td>
  </tr>
   <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">��վlogo</td> 
    <td height="30"><input name="logo" type="text" id="logo" value="<%=rs("zl_wel_logo")%>" size="50" readonly="readonly">
&nbsp;
<input type="button" name="Submit2" value="�ϴ�ͼƬ" onClick="window.open('ziliang_upload.asp?formname=myform&editname=logo&uppath=/ZiLiang_Images/upfile&filelx=jpg','','status=no,scrollbars=no,top=400,left=400,width=600,height=165')">
ͼƬ�ϴ�h:128px w:128px </td>
  </tr>
     <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">logo�鿴</td> 
    <td height="30"><img src="<%=rs("zl_wel_logo")%>" alt="<%=rs("zl_web_title")%>" /></td>
   </tr>
   <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">�б�����</td> 
    <td height="30">�б�����
     <input name="liebaio" type="text" id="liebaio" value="<%=rs("zl_wel_liebiao")%>" size="4" /></td>
   </tr>
     <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">��ʾ��ʽ</td> 
    <td height="30">ֻ��ʾ����
     <input type="radio" name="xianshi" value="0" <% if rs("zl_wel_liebaio_flag")=0 then response.Write("checked") end if%> />
     ������������ʾ
     <input type="radio" name="xianshi" value="1" <% if rs("zl_wel_liebaio_flag")=1 then response.Write("checked") end if%> /></td>
  </tr>
     <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">��Ŀ����</td> 
    <td height="30">��������
      <input name="links" type="checkbox" id="links" value="yes" <% if rs("zl_web_links_off")=1 then response.Write("checked") end if%>/> 
      �浵  
      <input name="cd" type="checkbox" id="cd" value="yes" <% if rs("zl_web_cd")=1 then response.Write("checked") end if%>/>
      ����  
      <input name="pr" type="checkbox" id="pr" value="yes" <% if rs("zl_web_pr")=1 then response.Write("checked") end if%>/> 
      ��������  
      <input name="wz" type="checkbox" id="wz" value="yes" <% if rs("zl_web_wz")=1 then response.Write("checked") end if%>/></td>
  </tr>
       <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">ͼƬ����</td> 
    <td height="30">�߶�
      <input name="height" type="text" id="height" value="<%=rs("zl_wel_height")%>" size="5" />
      ���
      <input name="width" type="text" id="width" value="<%=rs("zl_wel_width")%>" size="5" /></td>
  </tr>
   <tr>
    <td width="15%" height="30" align="center" valign="middle" bgcolor="#F9F9F9">&nbsp;</td> 
    <td height="30" bgcolor="#F9F9F9"><input type="submit" name="Submit" value="�޸Ĳ���" />
    <input type="hidden" name="config_id" value="<%=rs("config_id")%>"></td>
  </tr>
  </form>
</table>

</body>
</html>
