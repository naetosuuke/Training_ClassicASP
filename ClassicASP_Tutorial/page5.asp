<%@ LANGUAGE=VBScript %>

<HTML>
  <meta charset="UTF-8">
  <BODY>

    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338740(v=msdn.10)">
プロシージャを記述する</a>プロシージャを記述する<br/>
<br/>


"プロシージャ" とは、特定の作業を実行し、値を返すことのできるスクリプト コマンドの集まりのことです。<br/>
スクリプトでは独自のプロシージャを定義することができ、定義したプロシージャを何度でも呼び出すことができます。<br/>
※VBScriptはSub ~ End Sub でくくり、JScriptではfunction ~(){ } でくくる。関数とほぼ<br/>
<br/>


プロシージャの定義は、プロシージャを呼び出す .asp ファイル自体に記述することができます。<br/>
また、いくつかのよく使うプロシージャを 1 つの共有 .asp ファイルで定義し、これらのプロシージャを呼び出す側の <br/>
.asp ファイルで #include ディレクティブを使用して、共有 .asp ファイルをインクルードすることもできます。<br/>
さらに、これらの機能を COM コンポーネントにカプセル化することもできます。<br/>
<br/>

<br/>
プロシージャに配列を渡す<br/>
VBScript でプロシージャに配列全体を渡すには、配列名の後に空のかっこを付けて指定します。<br/>
JScript の場合は、空の角かっこを付けて指定します。<br/>
<br/>


プロシージャの実行、定義方法

<% Echo 'VBScriptのプロシージャを実行 %>
  <BR>
<% printDate()'Jscriptのプロシージャを実行 %><br/>
注   VBScript から JScript 関数を呼び出す場合、大文字と小文字は区別され "ません "。
  <BR>
  </BODY>
</HTML>


<%
Sub Echo 'VBScriptのプロシージャを宣言　通常のデリミタ内でOK

  Response.Write "<TABLE>" & _
    "<TR><TH>Name</TH><TH>Value</TH></TR>"

  Set objQueryString = Request.QueryString

  For Each strSelection In objQueryString
    Response.Write "<TR><TD>" & p & "</TD><TD>" & _
    FormValues(strSelection) & "</TD></TR>"
  Next

  Response.Write "</TABLE>"

End Sub
%>

<SCRIPT LANGUAGE=JScript runat=server> 
//JScriptのプロシージャを宣言　Scriptタグで内包する必要がある(デリミタは不要)
//ScriptタグにLanguage=使用言語、Runat=server or clientでの実行 を指定する。

function printDate()
{
  var x
  x = new Date()
  Response.Write(x.toString())
}
</SCRIPT><br/>


