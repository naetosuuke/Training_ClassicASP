<%@ LANGUAGE=JScript %>
 %@ は処理ディレクティブ。.aspファイルの先頭行に記述しなければならない。<br/>
 %@ LANGUAGE= では、aspファイル上のコードをVBScript, Jscriptのどちらでコンパイルするかを決める<br/>
 (構文がまるっと違うので、基本的に同一ファイルで共存不可能)<br/>
<br/>

<HTML>
  <meta charset="UTF-8">
  <BODY>
  This page was last refreshed on <%= Now() %>.


<% //JScriptコマンドを使用している場合、ステートメントのブロックを示す中かっこの間に、
   //HTMLタグ、およびテキストがあっても、そのまま挿入できる。

  if (screenresolution == "low")
  {
%>
This is the text version of a page.
<%
  }
  else
  {
%>
This is the multimedia version of a page.
<%
  }
%>


<% //別の書き方（Response.Writeで表示する）
  if (screenresolution == "low")
  {
    Response.Write("This is the text version of a page.")
  }
  else
  {
    Response.Write("This is the multimedia version of a page.")
  }
%>

  </BODY>
</HTML>