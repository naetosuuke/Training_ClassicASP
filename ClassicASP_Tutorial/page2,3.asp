<% Option Explicit %>

<HTML>
  <meta charset="UTF-8">
  <BODY>

  <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338754(v=msdn.10)"></a>スクリプト言語を使用する<br/>
<br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338744(v=msdn.10)"></a>変数と定数を使用する<br/>
<br/>

    VBScript<br/>
<br/>
    変数の宣言<br/>
    <% Dim UserName %><br/>

    <%
    Dim strUserName
    Public lngAccountNumber
    %>

    グローバルスコープ、ローカルスコープ<br/>
    変数のスコープをプロシージャに限定すると、パフォーマンスが向上します。<br/>
<br/>
    <%   
     '= 変数の明示的な宣言を強制
    Dim X

    X = 1
    SetLocalVariable '関数

    Response.Write X '出力は1　ローカル内の新規Yは影響しない。

    Sub SetLocalVariable
        Dim X 'ローカルスコープ内で新規にYを宣言
        X = 2
    End Sub
    %><br/>


    <%
    Dim Y : Y = 1

    SetLocalVariable

    Response.Write Y '出力は2 (関数内でグローバルスコープの変数を参照してたから)

    Sub SetLocalVariable
        Y = 2 '新規で宣言しないと、グローバルスコープの変数を参照する
    End Sub
    %>
<br/>

<br/>
セッション スコープを持つ変数は、ある特定のユーザーに要求された 1 つの ASP アプリケーション内の<br/>
すべてのページで利用できます。<br/>
<br/>
アプリケーション スコープを持つ変数は、任意のユーザーに要求された<br/>
1 つの ASP アプリケーション内のすべてのページで利用できます。<br/>
<br/>
セッション変数は、ユーザー設定、ユーザー名、ユーザーの識別など、ユーザーごとの情報を格納する<br/>
場合に便利です。<br/>
<br/>
アプリケーション変数は、アプリケーション固有のメッセージ、アプリケーションで必要になる<br/>
一般的な値など、あるアプリケーションのすべてのユーザーに共通する情報を格納する場合に便利です。<br/>


<% 'セッションの宣言 ユーザー毎に変数を保持
  Session("FirstName") = "Jeff"
  Session("LastName") = "Smith"
%>

  セッションの呼び出し(アプリケーション内ならどこでも可能)<br/>
  Welcome <%= Session("FirstName") %> 

<br/>
ユーザーの設定を Session オブジェクトに格納しておくと、格納した設定に後からアクセスして、<br/>
どのページをユーザーに返すかを決定することができます。たとえば、アプリケーションの最初のページで、<br/>
ユーザーがテキスト専用バージョンのコンテンツを指定できるようにしておけば、この設定を、<br/>
アプリケーション内でユーザーが引き続き利用するすべてのページに対して適用することができます。<br/>
<br/>
アプリケーションオブジェクトに格納<br/>
<% Application("Greeting") = "Welcome to the Sales Department!" %>

<br/>
アプリケーション内の後続の任意のページから名前付きエントリにアクセスできます。<br/>
次の例は、出力ディレクティブを使用して Application("Greeting") の値を表示します。<br/>
<%= Application("Greeting") %>
<br/>

スクリプトで繰り返しセッション/アプリケーション スコープを持つ変数を参照する場合は、<br/>
その変数をローカル変数に割り当てるとパフォーマンスが向上します<br/>



定数<br/>
ActiveX Data Objects (ADO) など、<br/>
ASP に付属する一部の基本コンポーネントでは、スクリプトで使用できる定数が定義されています<br/>
<br/>
コンポーネントは、"コンポーネント タイプ ライブラリ" で定数を宣言できます。
コンポーネント タイプ ライブラリとは、COM コンポーネントによってサポートされるオブジェクトと型についての<br/>
情報が格納されているファイルのことです。<br/>
.asp ファイルでタイプ ライブラリを宣言すると、同じ .asp ファイル内の任意のスクリプトで定義済み定数を使用できます。<br/>
<br/>
タイプ ライブラリを宣言するには、.asp ファイルまたは Global.asa ファイルで <METADATA> <br/>
タグを使用します。たとえば、ADO タイプ ライブラリを宣言するには、次のステートメントを使用します。<br/>


これで、ADOタイプライブラリを使用できる。<br/>
<!--METADATA NAME="Microsoft ActiveX Data Objects 2.5 Library" TYPE="TypeLib" UUID="{00000205-0000-0010-8000-00AA006D2EA4}"-->
ファイルパスでも可能<br/>
<!-- METADATA TYPE="typelib" FILE="c:\program files\common files\system\ado\msado15.dll"-->
<br/>
このように宣言した後は、タイプ ライブラリを宣言した .asp ファイルや、ADO タイプ ライブラリ宣言をした<br/>
Global.asa ファイルが含まれているアプリケーションに常駐する .asp ファイルで、ADO 定数を使用できます<br/>



<a href= "https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338716(v=msdn.10)">タイプライブラリの詳細</a>


<%
  'Create and Open Recordset Object. 
  Set rstCustomerList = Server.CreateObject("ADODB.Recordset")

  rstCustomerList.ActiveConnection = cnnPubs
  rstCustomerList.CursorType = adOpenKeyset 'adOpenKeyset=ADO定数
  rstCustomerList.LockType = adLockOptimistic 'adLockOptimistic=ADO定数
%>

一般に使用されるタイプライブラリと UUIDの一覧
Microsoft ActiveX Data Objects 2.5 Library	{00000205-0000-0010-8000-00AA006D2EA4}
Microsoft CDO 1.2 Library for Windows 2000 Server	{0E064ADD-9D99-11D0-ABE5-00AA0064D470}
MSWC Advertisement Rotator Object Library	{090ACFA1-1580-11D1-8AC0-00C0F00910F9}
MSWC IIS Log Object Library	{B758F2F9-A3D6-11D1-8B9C-080009DCC2FA}


  </BODY>
</HTML>