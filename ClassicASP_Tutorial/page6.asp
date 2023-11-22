<HTML>
  <meta charset="UTF-8">
  <BODY>

<h3>・・・・・・・・・・・コレクションを格納するオブジェクトの種類を覚える・・・・・・・・・・・</h3><br/>

    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338731(v=msdn.10)">
コレクションを使用する(構造体)</a><br/>
<br/>

  <br/>
ASP の組み込みオブジェクトの多くは、コレクションを提供しています。"コレクション" は、文字列、数値、オブジェクト、<br/>
およびその他の値を格納する配列と類似したデータ構造です。ただし、コレクションから項目が取り出されたり、<br/>
コレクションに項目が格納されたりすると、それに合わせて自動的にコレクションが拡大または縮小する点が、<br/>
配列とは異なります。コレクションに対して変更が加えられると、コレクション内での項目の位置も移動します。<br/>
一意の文字列キーおよびコレクション内でのインデックス (位置) を使用したり、コレクション内のすべての項目に対して<br/>
操作を繰り返し実行したりすることにより、コレクション内の項目にアクセスすることができます。<br/>
<br/>


コレクション内の特定の項目には、一意の文字列キー、またはその名前を使用してアクセスすることができます。<br/>
たとえば、Contents コレクションは、Session オブジェクトに格納されたすべての変数を保持しています。<br/>
また、Contents コレクションは、Server.CreateObject を呼び出してインスタンス生成されたオブジェクトも<br/>
すべて保持しています。<br/>
たとえば、次のユーザー情報を Session オブジェクトに格納したとします。<br/>
<br/>

<% 'キーバリューで記憶させる。
  Session.Contents("FirstName") = "Meng"
  Session.Contents("LastName") = "Phua"
  Session.Contents("Age") = 29
%>
 対象のコレクションにキーを入れると値が出る<br/>
<%= Session.Contents("FirstName") %> = "Meng"<br/>
<br/>
インデックス番号でも値が出る<br/>
※※コレクションの院デック番号は、0からでなく1から開始する。気を付ける!<br/>
<%= Session.Contents(2) %> = "Phua"<br/>
<br/>
オブジェクトにコレクションが1つしかなければ、オブジェクト名だけで呼出可能<br/>
<%= Session("FirstName") %><br/>
<br/>

Application オブジェクトまたは Session オブジェクトに格納された項目にアクセスする場合は、<br/>
通常、コレクション名を省略しても目的の項目に安全にアクセスできます。しかし、<br/>
Request オブジェクトの場合は、コレクション名を指定する方がより安全です。<br/>
これは、コレクションに重複した名前を持つ項目が含まれる可能性が高いためです。(オブジェクトの項を要参照)<br/>
<br/>


コレクションへの繰り返し処理<br/>
VBScript For Each<br/>
<%
  'Declare a counter variable.
  Dim strItem

  'VBScriptが標準で持つforeachステートメント。プロパティの方が違ってもごり押しで実行される？
  For Each strItem In Session.Contents
    Response.Write Session.Contents(strItem) & "<BR>"
  Next
%><br/>
<br/>

VBScript For..Next<br/>
<%
  'Declare a counter variable.
  Dim intItem

   'VBScriptが標準で持つFor...Next ステートメントでループ
  For intItem = 1 To 3
    Response.Write Session.Contents(intItem) & "<BR>"
  Next
%><br/>
<br/>

VBScript For..Next(ループ回数はカウントで取得した数)<br/>
<%
  'Declare a counter variable.
  Dim lngItem, lngCount

  lngCount = Session.Contents.Count 'コレクションはカウントプロパティをもってるので、これでForNext構文の最後の値を指定できる。

  For lngItem = 1 To lngCount
     Response.Write Session.Contents(lngItem) & "<BR>"
  Next
%><br/>
<br/>

JScript for<br/>
<Script language=JScript runat=SERVER>
  function printItems() {
    var intItemJS, intNumItems;
      
    intNumItems = Session.Contents.Count;
      
    for (intItemJS = 1; intItemJS <= intNumItems; intItemJS++)
    {
      Response.Write(Session.Contents(intItemJS) + "<BR>") 
      //スクリプトタグ内でResponse.Writeを実行してもブラウザに表示されない。なので関数に定義してデリミタで実行
    }
  }
</Script>
<%= printItems() %><br/>
<br/>


JScript Enumerator オブジェクトの利用<br/>
<br/>
Microsoft JScript では、Enumerator オブジェクトがサポートされており、<br/>
このオブジェクトを使って ASP のコレクションに対する繰り返し操作を実行することもできます。<br/>
atEnd メソッドは、コレクション内にまだ項目があるかどうかを示します。<br/>
moveNext メソッドは、コレクション内の次の項目に移動します。<br/>
<br/>

<Script language=JScript runat=SERVER>
function testEnumerator(){
    Session.Contents("Name") = "Suki White"
    Session.Contents("Department") = "Hardware"
    Session.Contents("Sex") = "Female"

    //numeratorオブジェクトに、Session.Contentsの中身を渡して初期化
    var mycollection = new Enumerator(Session.Contents);

    //反復処理
    while (!mycollection.atEnd()) //Mycollectionの中身が最後になるまで（最後になっても実行、繰り返しを止める。）
    {
      var x  = mycollection.item(); //itemメソッドを初期化
      Response.Write(Session.Contents(x) + "<BR>");
      mycollection.moveNext(); //次の中身に移る
    }
  }
</Script>
<%= testEnumerator() %><br/>
<br/>

<h3>・・・・・・・・・・・Cookie, セッションの内容が習熟できていない。PHPの動画すすめて覚える...?・・・・・・・・・・・</h3><br/>
<br/>
<br/>

サブキーを使用してコレクション全体に対し繰り返し操作を実行する<br/>
<br/>
スクリプトでは、ブラウザと Web サーバーの間でやり取りされる cookie の数を減らすために、<br/>
いくつかの関連する値を 1 つの cookie に埋め込むことがあります。このため、Request オブジェクトおよび <br/>
Response オブジェクトの Cookies コレクションは、1 つの項目内に複数の値を保持していることがあります。<br/>
これらのサブ項目 (サブキー) には、個別にアクセスできます。サブキーをサポートしているのは、<br/>
Request.Cookies コレクションと Response.Cookies コレクションのみです。Request.Cookies は読み取り操作のみを、<br/>
Response.Cookies は書き込み操作のみをサポートしています。<br/>
<br/>


次の例は、通常の cookie とサブキーを持つ cookie を作成します。<br/>
<br/>

  
<%
  'サーバーからブラウザにクッキーを送信.
  Response.Cookies("Fruit") = "Pineapple" 

  'サーバーからブラウザにクッキーを送信(サブキーあり).
  Response.Cookies("Mammals")("Elephant") = "African"
  Response.Cookies("Mammals")("Dolphin") = "Bottlenosed"
%><br/>


HTTP 応答内の cookie テキストはブラウザに送信され、次のように表示されます。<br/>
HTTP_COOKIE= Mammals=ELEPHANT=African&DOLPHIN;=Bottlenosed; Fruit=Pineapple;<br/>
<br/>

Request.Cookies コレクション内のすべての cookie、および 1 つの cookie 内のすべてのサブキーを、<br/>
列挙することもできます。ただし、サブキーを持たない cookie に対して、サブキーを使った繰り返し操作を実行しても、<br/>
何も出力されません。これを避けるには、Cookies コレクションの HasKeys 属性を使って cookie がサブキーを<br/>
持っているかどうかを先にチェックしておきます。この例を次に示します。<br/>
<br/>

<%
   'Declare counter variables.
   Dim Cookie, Subkey

   'Display the entire cookie collection.
   For Each Cookie In Request.Cookies 'クライアントから届くクッキー達
     Response.Write Cookie 'クッキーをを表示
     If Request.Cookies(Cookie).HasKeys Then 'キーを持つクッキーに対して
       Response.Write "<BR>"  'キーを表示
       'Display the subkeys.
       For Each Subkey In Request.Cookies(Cookie) '入れ子構造でサブキーをもつクッキーだけ表示させる。
       Response.Write " ->" & Subkey & " = " & Request.Cookies(Cookie)(Subkey) & "<BR>"
       Next
     Else
       Response.Write " = " & Request.Cookies(Cookie) & "<BR>"
     End If
   Next
%><br/>
<br/>

上記スクリプトの実行結果<br/>
  <br/>
Mammals  <br/>
->ELEPHANT = African  <br/>
->DOLPHIN = Bottlenosed  <br/>
Fruit = Pineapple  <br/>

キーの大文字小文字<br/>
<br/>
Cookies コレクション、Request コレクション、Response コレクション、および ClientCertificate コレクションは、<br/>
同じ一意の文字列キー名を参照することができます。たとえば、次のステートメントでは、<br/>
User という同じキー名を参照していますが、コレクションによって異なる値が返されます。<br/>
<br/>

<%
strUserID = Request.Cookies("User") 'サーバーにユーザー情報が送信される。
strUserName = Request.QueryString("User")
%><br/>

キー名の大文字と小文字は、キーに値を割り当てる最初のコレクションによって設定されます。たとえば、<br/>
次のスクリプトの場合を考えてみます。<br/>
<br/>

<%
  'Retrieve a value from QueryString collection using the UserID key.
  'UserIDのクエリはUserIDとして定義したが、下のCookie取得時のUserIdはディーが小文字のキーとして受け取っている。
  strUser = Request.QueryString("UserID")
				
  'Send a cookie to the browser, but reference the key, UserId, which has a different spelling.
  '正しくはUserIDとして定義されているが、コレクションの値への参照では、大文字小文字の区別がないのでこれでも問題ない
  Response.Cookies("UserId")= "1111-2222"
  Response.Cookies("UserId").Expires="December 31, 2000"
%><br/>
<br/>

<h3>・・・・・・・・・・・Objectが習熟できていない。ASPのチュートリアル進めた後戻ってくる・・・・・・・・・・・</h3><br/>
オブジェクトのコレクションに対して繰り返し操作を実行する<br/>
<br/>

Session コレクションと Application コレクションは、スカラ変数またはオブジェクト <br/>
インスタンスのいずれかを保持することができます。Contents コレクションは、スカラ変数、および <br/>
Server.CreateObject を呼び出して作成されたオブジェクトのインスタンスの両方を保持できます。<br/>
<br/>
StaticObjects コレクションは、Session オブジェクトのスコープ内で HTML の <OBJECT> タグを使用して<br/>
作成されたオブジェクトを保持できます。この方法でオブジェクトのインスタンスを作成する方法の詳細については、<br/>
<a href= "https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338739(v=msdn.10)"><br/>
「オブジェクトのスコープを設定する」</a> を参照してください。<br/>
<br/>

オブジェクトを含むコレクションに対して繰り返し操作を実行する場合は、オブジェクトの Session 状態または <br/>
Application 状態の情報にアクセスするか、あるいはオブジェクトのメソッドまたはプロパティにアクセスすることができます。<br/>
たとえば、ユーザー アカウントを作成するために、アプリケーションでいくつかのオブジェクトを使用していて、<br/>
各オブジェクトに初期化メソッドがある場合、StaticObjects コレクションに対して繰り返し操作を実行すると、<br/>
オブジェクトのプロパティを取得することができます。<br/>
<br/>

  
<% 'オブジェクトを含むコレクションも対応可能。
  For Each Object in Session.StaticObjects
    Response.Write Object & ": " & Session.StaticObjects(Object)
  Next
%>


  </BODY>
</HTML>

