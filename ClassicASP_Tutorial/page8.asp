<HTML>
  <meta charset="UTF-8">
  <BODY>


    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338732(v=msdn.10)">
コンポーネントとオブジェクトを使用する(構造体)</a><br/>
<br/>


<b>コンポーネント</b><br/>
 -> ダイナミックリンク ライブラリ (.dll) または実行可能 (.exe) ファイルに含まれる<br/>
実行可能コードです。コンポーネントは、1 つまたは複数の "オブジェクト" を提供<br/>
<br/>

<b>オブジェクト </b><br/>
-> コンポーネント内に組み込まれて特定の機能を実行するコードの単位です。<br/>
オブジェクトを使用するには、インスタンスを作成し、変数名に割り当てる。<br/>
インスタンス作成には、ASPのServer.CreateObject メソッド、または HTML の <OBJECT> タグを使用<br/>
その際、オブジェクトの登録名 (PROGID) を指定する
<br/>

例 AdRotatorコンポーネントのAdRotatorオブジェクトを使用する場合<br/>
<h2>※※ AdRotatorコンポーネント 使用できない(設定が足りない？)※※</h2>

<br/>
<b>Server.CreateObjectを使用</b><br/>
VBScriptのCreateObject、JScriptのNewといった、新機関数の宣言メソッドは使わない<br/>
※"MSWC.AdRotator"がPROGID<br/>
<br/>

Set MyAds = Server.CreateObject("MSWC.AdRotator") 'VBScriptの場合 
var MyAds = Server.CreateObject("MSWC.AdRotator") //JScriptの場合

<b>HTML OBJECTタグを使用</b><br/>

PROGID使用<br/>
&lt;OBJECT RUNAT=Server ID=MyAds PROGID="MSWC.AdRotator"></OBJECT><br/>
<br/>

CLSID使用<br/>
&lt;OBJECT RUNAT=SERVER ID=MyAds<br/>
CLASSID="Clsid:1621F7C0-60AC-11CF-9427-444553540000"></OBJECT> <br/>
<br/>


自分でCOM コンポーネント(Windowsコンポーネント)を作成できる。<br/>

<!-- 

スクリプト例 (Cercius⇔Fahrenheit変換)

<SCRIPTLET>

<Registration
	Description="ConvertTemp"
	ProgID="ConvertTemp.Scriptlet"
	Version="1.00"
>
</Registration>

<implements id=Automation type=Automation>
	<method name=Celsius>
		<PARAMETER name=F/>
	</method>
	<method name=Fahrenheit>
		<PARAMETER name=C/>
	</method>
</implements>

<SCRIPT LANGUAGE=VBScript>

  Function Celsius(F)
	  Celsius = 5/9 * (F - 32)
  End Function

  Function Fahrenheit(C)
	  Fahrenheit = (9/5 * C) + 32
  End Function

</SCRIPT>
</SCRIPTLET>

-->
この Windows スクリプト コンポーネントを実装する前に、このファイルに .sct 拡張子を付けて保存し、<br/>
次に Windows エクスプローラでファイルを右クリックして [登録] をクリックする必要があります。<br/>
※現環境で登録しようとしたところエラー発生。原因が究明できないのでスキップ<br/>
<br/>
呼出スクリプト<br/>
<br/>
<%
  Dim objConvert
  Dim sngFvalue, sngCvalue

  sngFvalue = 50
  sngCvalue = 21

  'Set objConvert = Server.CreateObject("ConvertTemp.Scriptlet")
%>

<%= sngFvalue %> degrees Fahrenheit is equivalent to objConvert.Celsius(sngFvalue)(デリミタで囲い、実行結果を表示) degrees Celsius<BR>
<%= sngCvalue %> degrees Celsius is equivalent to objConvert.Fahrenheit(sngCValue)(デリミタで囲い、実行結果を表示) degrees Fahrenheit<BR>


<b>ASPの組み込みオブジェクトを使用する</b><br/>
<br/>

基本構文<br/>
Object.Method parameters<br/>
例<br/>
Response.Write "Hello World"<br/>
<br/>

<b>Java クラスからオブジェクトを作成する</b>
  
<%
  Dim dtmDate
  'Set dtmDate = GetObject("java:java.util.Date")
%>
The date is dtmDate.toString()(デリミタで囲い、実行結果を表示)


</BODY>
</HTML>
