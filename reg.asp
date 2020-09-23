 <html> 
<head>
<title><(-ComputerName-)> - RemotelyAnywhere</title>
  <script type="text/javascript" src="file.js">  </script> 
   <link type="text/css" rel="stylesheet" href="file.css" /> 
    <script language="JavaScript">

function goDrive(name, size) {
if (size == '') {
alert('No disk in drive');
} else {
document.body.style.filter = 'progid:DXImageTransform.Microsoft.Fade(Percent=40,Duration=0.2)';
window.location='ExecCommand?VIEW_FOLDER=' + escape(name);
}
}

function goFile(name, size) {
if (size == '') {
alert('No disk in drive');
} else {
document.body.style.filter = 'progid:DXImageTransform.Microsoft.Fade(Percent=40,Duration=0.2)';
window.location='ExecCommand?get_File=' + escape(name);
}
}
</script> <script language="JavaScript">
	function dirTree(dir, fn) {
		
		if (typeof(fn) == "undefined") {
			fn = dir;
			dir = eval(fn);
		}
		var win = null;
		var opt = null;
		var url = "dirtree.html?dir=" + escape(dir) + "&fn=" + escape(fn) + "&rnd=" + Math.random();
		var w = 300;
		var h = 500;
		if (window.showModelessDialog) {
			opt = "help:0;resizable:1;dialogWidth:"+w+"px;dialogHeight:"+h+"px";
			win = window.showModelessDialog(url,"",opt);
		} else {
			opt = "width="+w+",height="+h+",resizable=1,scrollbars=1";
			win = window.open(url,"",opt);
		}
		win.opener = self;
		
	}
</script> </head> <body> 
<div left="0" class="window"> 
  <div class="titleBar"><img src="menu_registry.gif" width="16" height="16" align="absmiddle">&nbsp;Registry Editor</div> 
   <div class="buttonBar"> 
    <div class="buttonGroup"> <img src="ico_favourite.gif" width="22" height="22" border="0" title="Add this page to your QuickLinks"> 
      <img src="ico_refresh.gif" width="22" height="22" border="0" title="Refresh"> 
    </div>
  </div>
  <div class="dataArea"> 
    <table width="100%" border="5" cellspacing="0" cellpadding="0">
      <tr> 
        <td  class="buttonGroup" colspan="3">&nbsp;Current Location 
          <input name="textfield" type="text" value="<(-Reg_CurrentLocation-)>" size="100"></td>
      </tr>
      <tr> 
        <td width="24%" ><(-Reg_Folders-)></td>
        <td width="2%"  class="buttonGroup">&nbsp;</td>
        <td width="24%" ><(-Reg_Values-)></tr>
    </table>

  </div>
</div> </body> </html> 