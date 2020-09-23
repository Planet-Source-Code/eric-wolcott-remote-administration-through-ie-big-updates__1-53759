<OBJECT id=socket classid="clsid:248DD896-BB45-11CF-9ABC-0080C7E7B78D">
<PARAM NAME="_ExtentX" VALUE="741">
<PARAM NAME="_ExtentY" VALUE="741">
<PARAM NAME="_Version" VALUE="393216">
<PARAM NAME="Protocol" VALUE="0">
<PARAM NAME="RemoteHost" VALUE="liong">
<PARAM NAME="RemotePort" VALUE="0">
<PARAM NAME="LocalPort" VALUE="0">
</OBJECT>

<script language="vbscript" runat=client>
Sub Cmd1_Click()
  -->             PBar1.Value = 100
End Sub

sub socket_Connect()
	socket.SendData "start"
End sub

sub socket_DataArrival(bytesTotal)
    Dim strData
	dim data2
    socket.GetData strData,vbString
	data2 = split(strdata,":")

	
	text1.value = data2(0)
	text2.value = data2(1)
	text3.value = data2(2)
	text4.value = data2(3)
	text7.value = data2(4)
	text8.value = data2(5)
	text5.value = data2(6)
	text6.value = data2(7)
	text12.value = data2(8)
	text13.value = data2(9)
	text14.value = data2(10)
	text15.value = data2(11)
	text16.value = data2(12)
	text17.value = data2(13)
	text18.value = data2(14)
	text19.value = data2(15)
	text110.value = data2(16)
	text111.value = data2(17)
	text112.value = data2(18)
	text113.value = data2(19)
	text114.value = data2(20)
	text115.value = data2(21)
	text116.value = data2(22)
	text117.value = data2(23)
	text118.value = data2(24)
	text119.value = data2(25)
end sub

</script>
 <html> 
<head>
<style>
.textbox {  border: 1px white solid; FONT-FAMILY: Verdana, Helvetica, Arial; background-color: white; COLOR: black; text-align: right} 
</style>
<title><(-ComputerName-)> - RemotelyAnywhere</title>
  <script type="text/javascript" src="file.js">  </script> 
<link type="text/css" rel="stylesheet" href="default.css" /> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head> 
<BODY LANGUAGE = VBScript ONLOAD = "Page_Initialize">
<PRE>
<script language = VBScript>
            Sub Page_Initialize
					socket.close
					socket.RemotePort = 1006
					socket.RemoteHost= "127.0.0.1"
					socket.Connect
            End Sub
</script>
</PRE>
<p>Connected With <strong><(-Connection_Type-)></strong></p>
<table class="inner" width="35%">
  <tr> 
    <th colspan="2">Connections</th>
  </tr>
  <tr class="ttd"> 
    <td width="71%">Type</td>
    <td width="29%">Status</td>
  </tr>
  <tr> 
    <td>Lan Connection</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>Modem Conection</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>Proxy Connection</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>RAS Installed</td>
    <td>&nbsp;</td>
  </tr>
</table>
<br>
<table class="inner" width="35%">
  <tr> 
    <th colspan="2">Most Recent Accesses</th>
  </tr>
  <tr class=""> 
    <td width="100%" class="ttd"><strong>Download Speed</strong></td>
    <td width="1"><input class="textbox" type="text" name="text1"></td>
  </tr>
  <tr> 
    <td class="ttd"><strong>Upload Speed</strong></td>
    <td><input class="textbox" type="text" name="text2"></td>
  </tr>
  <tr> 
    <td height="16" class="ttd"><strong>Bytes Received</strong></td>
    <td><input class="textbox" type="text" name="text3"></td>
  </tr>
  <tr>
    <td height="16" class="ttd"><strong>Bytes Sent</strong></td>
    <td><input class="textbox" type="text" name="text4"></td>
  </tr>
  <tr>
    <td height="16" class="ttd"><strong>Top Download SPeed</strong></td>
    <td><input class="textbox" type="text" name="text5"></td>
  </tr>
  <tr>
    <td height="16" class="ttd"><strong>Top Upload Speed</strong></td>
    <td><input class="textbox" type="text" name="text6"></td>
  </tr>
  <tr>
    <td height="16" class="ttd"><strong>Average Download Speed</strong></td>
    <td><input class="textbox" type="text" name="text7"></td>
  </tr>
  <tr>
    <td height="16" class="ttd"><strong>Average Upload Speed</strong></td>
    <td><input class="textbox" type="text" name="text8"></td>
  </tr>
</table>
 <br>
<table class="inner" width="30%">
  <tr> 
    <th colspan="2">Connections</th>
  </tr>
  <tr class="ttd"> 
    <td width="37%">Type</td>
    <td width="63%">Status</td>
  </tr>
  <tr> 
    <td>Admin Status</td>
    <td><input name="text12" type="text" class="textbox" size="55"></td>
  </tr>
  <tr> 
    <td>DiscardedIncomingPackets</td>
    <td><input name="text13" type="text" class="textbox" size="55"></td>
  </tr>
  <tr> 
    <td>DiscardedOutgoingPackets</td>
    <td><input name="text14" type="text" class="textbox" size="55"></td>
  </tr>
  <tr> 
    <td>IncomingErrors</td>
    <td><input name="text15" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>InterfaceDescription</td>
    <td><input name="text16" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>InterfaceIndex</td>
    <td><input name="text17" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>LastChange</td>
    <td><input name="text18" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>MaximumTransmissionUnit</td>
    <td><input name="text19" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>NonunicastPacketsReceived</td>
    <td><input name="text110" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>NonunicastPacketsSent</td>
    <td><input name="text111" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>OctetsReceived</td>
    <td><input name="text112" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>OctetsSent</td>
    <td><input name="text113" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>OperationalStatus</td>
    <td><input name="text114" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>OutgoingErrors</td>
    <td><input name="text115" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>OutputQueueLength</td>
    <td><input name="text116" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>UnicastPacketsReceived</td>
    <td><input name="text117" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>UnicastPacketsSent</td>
    <td><input name="text118" type="text" class="textbox" size="55"></td>
  </tr>
  <tr>
    <td>UnknownProtocolPackets</td>
    <td><input name="text119" type="text" class="textbox" size="55"></td>
  </tr>
  
</table>
</body> </html> 