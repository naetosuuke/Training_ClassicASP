<HTML>
  <meta charset="UTF-8">
  <BODY>


    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338726(v=msdn.10)">
セッションを管理する(構造体)</a><br/>
<br/>

ASP の Session オブジェクト、およびサーバーによって生成された特殊なユーザー ID を使うと、<br>
各利用ユーザーを識別し、情報を収集する高機能なアプリケーションを作成できます<br>
<br>

ASP では、HTTP cookie という方法を使用してユーザー ID を割り当てます。<br>
<br>

セッションは、次の 4 とおりの方法で開始できます。<br>
<br>

新しいユーザーが、アプリケーションの .asp ファイルを識別する URL を要求し、<br>
そのアプリケーションの Global.asa ファイルに Session_OnStart プロシージャが含まれている。<br>
<br>

ユーザーが Session オブジェクトに値を格納する。<br>
<br>

サーバーが受け取った要求に含まれている SessionID cookie が無効である場合、新しいセッションが自動的に開始される。<br>
<br>

ユーザーがアプリケーションの .asp ファイルを要求し、アプリケーションの Global.asa ファイルが <OBJECT><br>
タグを使用してセッション スコープを持つオブジェクトのインスタンスを作成する<br>
<br>

指定された時間内に、ユーザーがアプリケーションのページを要求または更新しない場合、<br>
セッションは自動的に終了します。この値の既定値は 20 分です<br>

特定のセッションに対し、アプリケーションの既定のタイムアウトより短いタイムアウト間隔に設定するには、<br>
Session オブジェクトの Timeout プロパティを設定することもできます<br>
<%  Session.Timeout = 5  %><br>
<br>

セッションを終了させることも可能<br>
<% Session.Abandon %><br>
<br>

<b>SessionID と cookie について</b><br>
<br>

ユーザーが特定のアプリケーション内で最初に .asp ファイルを要求したときに、<br>
ASP は "SessionID" を生成します。それをブラウザにcookie として格納します。<br>
<br>

ASP はユーザーのブラウザに SessionID cookie を格納した後は、ユーザーが別の .asp ファイルを<br>
要求または、ほかのアプリケーションで実行されている .asp ファイルを要求する場合でも、<br>
同じcookieを再利用してセッションを追跡します。<br>
<br>

無効にする方法<br>
<br>

1. アプリケーションのセッション状態が無効に設定されている場合。<br>
2. ASP ページがセッションレスに定義されている場合<br>
&lt;%@ EnableSessionState=False %><br>


<%
  Session("FirstName") = "Jeff"
  Session("LastName") = "Smith"
%>

Welcome <%= Session("FirstName") %>

Session オブジェクトにユーザー設定を格納し、後でその設定情報にアクセスして、<br>
ユーザーに返すページを決定することができます<br>
<br>

<% If Session("ScreenResolution") = "Low" Then %>
  This is the text version of the page.
<% Else %>
  This is the multimedia version of the page.
<% End If %>

また、Session オブジェクトにオブジェクト インスタンスを格納することもできます。<br>
ただし、その場合、サーバーのパフォーマンスに影響することがあります<br>
<br>

Session オブジェクトの Contents コレクションには、セッションに格納されている変数、<br>
つまり HTML <OBJECT> タグを使わずに格納された変数すべてが含まれています。<br>
<br>

<% '保存した内容を削除　キーおよびindexで指定可能
  If Session.Contents("Purchamnt") <= 75 then
    Session.Contents.Remove("Discount")
  End If
%>

&lt;% 'すべて削除
Session.Content.RemoveAll()
%>

複数のサーバー全体のセッションを管理する<br>
<br>
負荷分散が行われている（サーバー複数ある）サイトで ASP セッション管理を使用するには、<br>
ユーザー セッション内のすべての要求を同じ Web サーバーに向ける必要があります。<br>

Session_OnStart プロシージャを作成し、Response オブジェクトを使用して、<br>
ユーザーのセッションが実行されている特定の Web サーバーにブラウザをリダイレクトするという方法があります。<br>
<br>


cookie の使い方<br>
<br>

cookie の値を設定するには、Response.Cookies を使用します。<br>
cookie がまだ存在しない場合、Response.Cookies は新しい cookie を作成します<br>
<br>

Cookieの発行<br>
<br>
<% Response.Cookies("VisitorID") = 49
'　Web ページの <HTML> タグより前に置く必要があります。 %>

Cookieの起源は、セッションが生きている間だけ<br>
ブラウザを再起動した後でもCookieを保持させるには、失効期限を設定する<br>
<br>
<%
  Response.Cookies("VisitorID") = 49
  Response.Cookies("VisitorID").Expires = "December 31, 2001"
%>

cookie には複数の値を設定でき、そのような cookie を "インデックス付き cookie"と呼ぶ<br>
<br>

<% Response.Cookies("VisitorID")("49") = "Travel" %>

Cookieの取得<br>
<br>
<%= Request.Cookies("VisitorID") %>

ASP によってユーザーの Web ブラウザに格納された各 cookie には、それぞれパス情報が入っています。<br>
基本的に、Webアプリケーションのaspファイルのディレクトリに準じて設定される。<br>
対応するWebアプリが開いたら、同じディレクトリにあるCookieに送られる。<br>
このパスは自分で設定もできる<br>
<br>

<%
  Response.Cookies("Purchases") = "12"
  Response.Cookies("Purchases").Expires = "January 1, 2001"
  Response.Cookies("Purchases").Path = "/SalesApp/Customer/Profiles/"
%>

常にユーザーの Web ブラウザが cookie を転送するようにできます。(セキュリティ状の問題を起こさないよう注意)<br>
<% Response.Cookies("Purchases").Path = "/" %>




<br>



</BODY>
</HTML>

