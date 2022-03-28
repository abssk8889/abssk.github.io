<!--#include file="../ziliang_inc/inc.asp"-->
<html>
<head>
<LINK href="backdoor.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<TABLE  cellSpacing=0 cellPadding=0 border=0  background="../ziliang_images/back_image/back_line_bg.png" width=189>

  <tr>
  <td width=21 align="center" valign="middle"><img src="../ziliang_images/back_image/back_left_7.gif" width="13" height="11"></td>
    <td height="30" width=168 align="left" valign="middle" bgcolor="#F7F7F7" background="../ziliang_images/back_image/back_left_3.png" class="sec_menuu" id="imgmenu3" style="cursor:hand" onClick="showsubmenu(3)">联系我们</td>
  </tr>
  <tr>
  <td width=21></td>
   <td id="submenu3">
		<div class="sec_menu" style="WIDTH: 167px">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border:1px solid #F7F7F7">
      <!--DWLayoutTable-->
	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../ziliang_images/back_image/back_left_6.gif" width="7" height="10"> 电话：0571-86980893</td>
      </tr>
	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../ziliang_images/back_image/back_left_6.gif" width="7" height="10"> 联系人：<a href=http://aliang.zjxyk.com/ target=mainFrame>阿梁</a> </td>
      </tr>
	  	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../ziliang_images/back_image/back_left_6.gif" width="7" height="10"> QQ：41864438 </td>
      </tr>
	  	        <tr>
        <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../ziliang_images/back_image/back_left_6.gif" width="7" height="10"><a href=http://aliang.zjxyk.com/ target=mainFrame> 开发商：阿梁(Aliang)</a></td>
      </tr>
	  <tr>
     <td width="150" height="27" align="left" valign="middle" onMouseOut="this.className='bg2';" onMouseOver="this.className='bg';" class="bg2"><img src="../ziliang_images/back_image/back_left_6.gif" width="7" height="10"> <strong>回复区域</strong></td>
      </tr>  
	  <td width="150" height="27" align="left" valign="middle"><img src="../ziliang_images/back_image/back_left_6.gif" width="7" height="10"> <%call pinglunhz()%></td>
      </tr> 
    </table>
	</div>
	<br>
	<br>
	<br>
	<br><br><br>
	<br>
	<br>
	<br>
	<br>
  <tr>
	</td>
  </tr>
  
</TABLE>
<script>

function aa(Dir)
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//这里100为滚动速度
function StopScroll(){if(Timer!=null)clearTimeout(Timer)}

function initIt(){
divColl=document.all.tags("DIV");
for(i=0; i<divColl.length; i++) {
whichEl=divColl(i);
if(whichEl.className=="child")whichEl.style.display="none";}
}
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="../ziliang_images/back_image/back_left_1.png";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="../ziliang_images/back_image/back_left_3.png";
}
}

function loadingmenu(id){
var loadmenu =eval("menu" + id);
if (loadmenu.innerText=="Loading..."){
document.frames["hiddenframe"].location.replace("LeftTree.asp?menu=menu&id="+id+"");
}
}
</script>
</body>
</html>