
<%@ LANGUAGE=VBScript %>
 %@ は処理ディレクティブ。.aspファイルの先頭行に記述しなければならない。<br/>
 %@ LANGUAGE= では、aspファイル上のコードをVBScript, Jscriptのどちらでコンパイルするかを決める<br/>
 (構文がまるっと違うので、基本的に同一ファイルで共存不可能)<br/>
<br/>

<HTML>
 <meta charset="UTF-8">
  <BODY>
  This page was last refreshed on <%= Now() %>.

<br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338728(v=msdn.10)"></a>ASP ページを作成する<br/>
<br/>


<%
 ' ブラウザ側で表示するまでコンパイルできるか不明。まめに更新してチェックする
 ' コメントアウトはシングルクオート
 ' ASPではVBScriptとJScript どちらを使用するか選択できる。Jscriptの場合は//でコメントアウト
  Dim dtmHour ' 変数の宣言

  dtmHour = Hour(Now())

  If dtmHour < 12 Then
    strGreeting = "Good Morning!"
  Else 	
    strGreeting = "Hello!"
  End If
%><br/>


<%= strGreeting %><br/>
出力ディレクティブ (ディレクティブ内ではコメントできない)

<%
 ' 同一ファイル上で一度宣言した変数を、もう一度宣言しなおすことはできない
 ' Dim dtmHour
   Dim dtmHours

  dtmHours = Hour(Now())
  If dtmHours < 12 Then ' 1 つのステートメントのセクション間に HTML テキストを挿入することができます。
%>
  Good Morning!
<% Else %>
  Hello!
<% End If %><br/>


<%
 'ブラウザにHTMLテキストを返す < Response.Write "文字">
  Dim dtmHourss

  dtmHourss = Hour(Now())

  If dtmHourss < 12 Then
    Response.Write "Good Morning!"
  Else 	
    Response.Write "Hello!"
  End If
%>


デリミタの間のスペースは無視される
<% Color = "Green" %>

<%Color="Green"%>

<%
Color = "Green"
%>

  </BODY>
</HTML>