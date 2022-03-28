<!--#include file="../ziliang_inc/inc.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>阿梁：博客管理系统</title>
<LINK href="backdoor.css" type=text/css rel=stylesheet>
</head>
<body bgcolor="#E9F1FE">
<table width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td height="39" background="../ziliang_images/back_image/back_images_title.gif" widht=100%>&nbsp;<b>欢迎使用阿梁博客管理系统</b></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100%>&nbsp;版本号: 20100102 </td>
  </tr>
</table>
<br>
<table width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td height="39" background="../ziliang_images/back_image/back_images_title.gif" widht=100% colspan=2>&nbsp;<b>服务器信息</b></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;服务器接受语言:<%=Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")%></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;Server地址:<%=Request.ServerVariables("SERVER_NAME")%></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;访问者IP:<%=Request.ServerVariables("REMOTE_ADDR")%></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;访问者浏览器版本及系统:<%=Request.ServerVariables("HTTP_USER_AGENT")%></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;服务器可接受文件:<%=Request.ServerVariables("SERVER_PROTOCOL")%></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;WEB服务器软件及版本信息:<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;路由端口:<%=Request.ServerVariables("REMOTE_PORT")%></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;服务器http端口:<%=Request.ServerVariables("LOCAL_PORT")%></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;WEB目录名称:<%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;当前WEB页位置:<%=Request.ServerVariables("PATH_TRANSLATED")%></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;请求方式:<%=Request.ServerVariables("REQUEST_METHOD")%></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;传输协议:<%=Request.ServerVariables("SERVER_PROTOCOL")%></td>
  </tr>
</table>
<br>
<table width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td height="39" background="../ziliang_images/back_image/back_images_title.gif" widht=100% colspan=2>&nbsp;<b>联系我们</b></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;产品开发：杭州成楼网络科技有限公司</td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;项目开发：QQ 41864438 </td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;总机电话：0571-86980893</td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;产品咨询：0571-86980893</td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;传&nbsp;&nbsp;&nbsp;&nbsp;真：0571-86980893</td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;联 系 人：阿梁</td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;官方网站：<a href="http://www.speed-web.cn/">http://aliang.zjxyk.com/</a></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;帮助中心：<a href="http://www.speed-web.cn/">http://aliang.zjxyk.com/</a><a href="http://www.tradedoor.net/"></a></td>
  </tr>
</table>
<table width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td height="39" background="../ziliang_images/back_image/back_images_title.gif" widht=100% colspan=2>&nbsp;<b>FSO组件开起方法</b></td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50% style="padding:10px;"><strong>win2000／ＸＰ系统：</strong><BR></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50% style="padding:10px;"><strong>&nbsp;windows2003系统</strong></td>
  </tr>
  <tr>
    <td height="60" bgcolor="#E9F1FE" widht=100% width=50% style="padding:10px;">在CMD命令行状态输入以下命令：<BR>
      关闭命令：RegSvr32 /u   C:\\windows\\SYSTEM32\\scrrun.dll<BR>
    打开命令：RegSvr32 C:\\windows\\SYSTEM32\\scrrun.dll</td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50% style="padding:10px;">在CMD命令行状态输入以下命令：<BR>
      关闭命令：RegSvr32 /u   C:\\WINDOWS\\SYSTEM32\\scrrun.dll<BR>
    打开命令：RegSvr32   C:\\WINDOWS\\SYSTEM32\\scrrun.dll</td>
  </tr>
  <tr>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50% style="padding:10px;"><strong>windows98系统</strong></td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;</td>
  </tr>
  <tr>
    <td height="60" bgcolor="#E9F1FE" widht=100% width=50% style="padding:10px;">在DOS命令行状态输入以下命令：<BR>
      关闭命令：RegSvr32 /u   C:\\WINDOWS\\SYSTEM\\scrrun.dll<BR>
    打开命令：RegSvr32 C:\\WINDOWS\\SYSTEM\\scrrun.dll</td>
    <td height="30" bgcolor="#E9F1FE" widht=100% width=50%>&nbsp;</td>
  </tr>
</table>

</body>
</html>
