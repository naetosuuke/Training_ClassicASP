
<HTML>
  <meta charset="UTF-8">
  <BODY>

    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338729(v=msdn.10)"></a>クライアント側のスクリプトと対話する<br/>
<br/>
クライアント側のスクリプトと対話する

<%
  Dim dtmTime, strServerName, strServerSoftware, intGreeting

  dtmTime = Time()
  strServerName = Request.ServerVariables("SERVER_NAME")
  strServerSoftware = Request.ServerVariables("SERVER_SOFTWARE")

  'Generate a random number. 		
  Randomize
  intGreeting = int(rnd * 3)
%>

  <SCRIPT LANGUAGE="JScript">  
  
  //Script内ではJScriptを使用できる。活用すればラウンドトリップ(応答時間)の改善、サーバー側負担を減らせる


  //Call function to display greeting
  showIntroMsg()

  function showIntroMsg()
  {
    switch(<%= intGreeting %>)
    {
    case 0:
      msg =  "This is the <%= strServerName%> Web server running <%= strServerSoftware %>."
      break
    case 1:			
      msg = "Welcome to the <%= strServerName%> Web server. The local time is <%= dtmTime %>."
      break
    case 2:
      msg = "This server is running <%= strServerSoftware %>."
      break
    }

  document.write(msg)
  }


</SCRIPT>


  </BODY>
</HTML>