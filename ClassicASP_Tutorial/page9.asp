<HTML>
  <meta charset="UTF-8">
  <BODY>


  <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338739(v=msdn.10)">
オブジェクトのスコープを設定する(構造体)</a><br/>
<br/>

オブジェクトのインスタンスは基本ページスコープ<br/>
同じASPファイルの任意のスクリプトから呼び出し可能<br/>
.aspファイルが要求の処理(寿命)を終了すると解放<br/>
<br/>

ループ内でオブジェクトを作成する<br/>
一般に、ループ内でのオブジェクトの作成は避ける。<br/>
ループ内でオブジェクトを作成する必要がある場合は、<b>オブジェクトが使用するリソースを手動で解放してください。</b><br/>
<br/>

<%
  Dim objAd
  For i = 0 To 1000
  '  Set objAd = Server.CreateObject("MSWC.AdRotator")
	'
  '  objAd.GetAdvertisement
  '
  '  Set objAd = Nothing 'Nothingを代入することで解放できる。
  Next
%>



<b>オブジェクトにセッション スコープを与える</b><br/>
"セッション スコープ" を持つオブジェクトは、、セッションが終了すると解放される。<br>
<br>

オブジェクトにセッション スコープを与えるのは、必要な場合のみに限られます。<br>


オブジェクトにセッション スコープを与えるには、オブジェクトを ASP の Session 組み込みオブジェクトに格納します。<br>
セッション スコープを持つオブジェクトのインスタンスを生成するには、Global.asa ファイルで <br>
HTML の <OBJECT> タグを使用する方法と、ASP ページ上で Server.CreateObject メソッドを使用する方法があります。<br>
<a href="http://asp.style-mods.net/topic16.html">Global.asaファイルについては、ここから確認</a><br>
<br>



Global.asa で設定する方法<br>
<br>

登録(asaファイルに記載する内容)<br>
&lt;OBJECT RUNAT=SERVER SCOPE=Session ID=MyBrowser PROGID="MSWC.BrowserType">
&lt;/OBJECT>
<br>

登録後は、任意のページからオブジェクトにアクセス可能<br>
&lt;%= If MyBrowser.browser = "IE"  and  MyBrowser.majorver >= 4  Then . . .%>
<br>

ASP ページでは、Server.CreateObject メソッドを使って Session 組み込みオブジェクトにオブジェクトを格納できる<br>
<br>

<% Set Session("MyBrowser") = Server.CreateObject("MSWC.BrowserType") %> 

ほかのASPファイルにあるブラウザ情報を表示するには、セッションオブジェクト内 MyBrowserオブジェクトの<br>
インスタンスを取得してから、メソッドを呼び出す。<br>
<br>

<% Set MyBrowser = Session("MyBrowser") %>
<%= MyBrowser.browser %>

Objectタグを使用した場合は、スクリプトコマンドでオブジェクトが参照されるまで、インスタンス生成されない。<br>
一方、Server.CreateObjectの場合は、その場でインスタンス生成される。<br>
よって、Objectタグを使用した方が拡張性が高い。<br>
<br>


<b>オブジェクトにアプリケーション スコープを与える</b><br>
<br>

"アプリケーション スコープ" を持つオブジェクトは、アプリケーションの起動時に作成されるオブジェクトの単一のインスタンスです。<br>
このオブジェクトは、すべてのクライアント要求において共有されます<br>
<br>

ASP の Application 組み込みオブジェクトにオブジェクトを格納。<br>
インスタンスの作成には、これもGlobal.asa ファイルで OBJECT タグを使用する方法と<br>
.asp ファイル上で Server.CreateObject メソッドを使用する方法があります。<br>
<br>

objectタグの使用(Adrotatorが動かないので、ブラウザに表示のみ)<br>
<br>

&lt;OBJECT RUNAT=SERVER SCOPE=Application ID=MyAds PROGID="MSWC.AdRotator"><br>
&lt;/OBJECT><br>
&lt;%=MyAds.GetAdvertisement("CustomerAds.txt") %><br>


&lt;% Set Application("MyAds") = Server.CreateObject("MSWC.Adrotator")%> 
&lt;%Set MyAds = Application("MyAds") %> &lt;%=MyAds.GetAdvertisement("CustomerAds.txt") %>





<br>

</BODY>
</HTML>

