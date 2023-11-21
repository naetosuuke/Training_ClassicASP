<HTML>
  <meta charset="UTF-8">
  <BODY>

    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338733(v=msdn.10)">ブラウザにコンテンツを送信する(構造体)</a><br/>
<br/>



コンテンツを送信する<br>

<%
  If blnFirstTime Then 'blnFirstTime = 初めて来たユーザーかどうかをBoolで返す予約変数
    Response.Write "<H3 ALIGN=CENTER>Welcome to the Overview Page.</H3>"
  Else
    Response.Write "<H3 ALIGN=CENTER>Welcome Back to the Overview Page.</H3>"
  End If
%><br>
<br>

<H3 ALIGN=CENTER>
<% If blnFirstTime Then %>
  Welcome to the Overview Page.
<% Else %>
  Welcome Back to the Overview Page.
<% End If %>
</H3><br>
<br>

<% '埋め込んだ場合　（ ＆は、VBScriptで文字列を連結する文字 _は行を続けるための文字）
Response.Write "<TR><TD>" & Request.Form("FirstName") _
 & "</TD><TD>" & Request.Form("LastName") & "</TD></TR>"
%><br>
<br>


HTMLヘッダのコンテンツタイプ文字列を設定する。<br>
<br>
<br>例
text/plain (HTML ステートメントでなく、テキストとして返されるコンテンツ)、<br>
image/gif (GIF イメージ)、image/jpeg (JPEG イメージ)、<br>
video/quicktime (Apple QuickTime® 形式の映画)、text/xml (XML ドキュメント)<br>
<br>

&lt;% Response.ContentType = "" %><br>
<br>

ブラウザをリダイレクトする<br>
<br>

ユーザーにコンテンツが送信される前に処理されるサーバー側スクリプトは、"バッファ処理される" といわれます。<br>
Response.Redirect ステートメントはテキストまたは <HTML> タグよりも前の、ページの先頭に記述し、<br>
その前にブラウザには何も送信されないようにしてください。<br>
<br>

<% '顧客IDを持っていない場合は、登録画面にリダイレクトさせる。※テストページのため、持っている場合はリダイレクトに逆転
  If Session("CustomerID") <> "" Then
    Response.Redirect "Register.asp"
  End If
%><br>
<br>



.ASP ファイル間で転送する<br>
<br>
Response.Redirect メソッドではなく Server.Transfer メソッドを使って、<br>
1 つの .asp ファイルから同じサーバー上の別のファイルに転送することができます<br>
<br>

<%
  If Session("blnSaleCompleted") Then
    Server.Transfer("/Order/ThankYou.asp")
  Else
  End if
%><br>
<br>

ASP には、ファイルに転送し、コンテンツを実行して、転送を開始したファイルに戻す Server.Execute コマンドもあります<br>

<% '動的に.aspファイルを読み込んでいる。
  ' .
  If blnUseDHTML Then
    Server.Execute("DHTML.asp")
  Else
  End If
  ' .
  ' .
  ' .
%><br>
<br>


コンテンツのバッファ処理を行う<br>
<br>

特に指定しない限り、Web サーバーはページ上のすべてのスクリプト コマンドを処理してから、<br>
ユーザーにコンテンツを送信します。この処理は "バッファ処理" といわれます（無効化もできる）<br>
<br>

  &lt; % '認証確認
    If Request("CustomerStatus") = "" Then
      Response.Clear
      Server.Transfer("/CustomerInfo/Register.asp")
    Else
      Response.Write "Welcome back " & Request("FirstName") & "!"
		    ' .
		    ' .
		    ' .
    End If
  %>

プロキシ サーバーがページをキャッシュできるようにする<br>
<br>

アプリケーションがクライアントに送信するページは、"プロキシ サーバー" を経由することがあります<br>
ASP では、特に指定しない限り、ASP ページ自体はキャッシュしないようにプロキシ サーバーに指示します <br>
(画像、イメージ マップ、アプレット、およびそのページから参照されているその他の項目はキャッシュされます)<br>
<br>

Response.CacheControl プロパティを使用して Cache-Control HTTP ヘッダー フィールドを設定すると、<br>
特定のページに対してキャッシュを有効にすることができます。Response.CacheControl の既定値は文字列 "Private" であり、<br>
プロキシ サーバーによるページのキャッシュは無効になっています。キャッシュを有効にするには、Cache-Control ヘッダー <br>
フィールドを Public に設定します。<br>
<br>

<% Response.CacheControl = "Public" %>

ブラウザ側でキャッシュを禁止する<br>
<br>
<% Response.Expires = 0 'キャッシュされたページの有効期限を０にしている%>

動的なチャンネルを作成する<br>
<br>
チャンネルは、ユーザーのコンピュータが定期的にサーバーに接続し、<br>
更新された情報を取り出すようにスケジュールします。この検索プロセスは、一般に "クライアント プル" と呼ばれます<br>
特定の Web サイトで新しい情報が入手可能になると、コンテンツがブラウザのキャッシュにダウンロードされます。<br>
Web ベースの情報を配布するために、特にイントラネット上でチャンネルを有効に使用すると、<br>
情報の中央集中化とサーバー トラフィックの削減に効果的です。<br>
<br>

<P> Choose the channels you want. (選択されたフォームは.cdxファイル内のスクリプトを呼び出し、定義を作成)</P>
<FORM METHOD="POST" ACTION="chan.cdx">
<P><INPUT TYPE=CHECKBOX NAME=Movies> Movies
<P><INPUT TYPE=CHECKBOX NAME=Sports> Sports
<P><INPUT TYPE="SUBMIT" VALUE="SUBMIT">
</FORM>

Chan.cdx 内のスクリプトは、要求と共にユーザーから送信されたフォームの値に基づいて、チャンネル定義を作成します。<br>

&lt;% If Request.Form("Movies") <> "" Then %>
  <CHANNEL>
    channel definition statements for the movie pages
  </CHANNEL>
&lt;% End If %>

&lt;% If Request.Form("Sports") <> "" Then %>
  <CHANNEL>
    channel definition statements for the sports pages
  </CHANNEL>
&lt;% End If %>







</BODY>
</HTML>

