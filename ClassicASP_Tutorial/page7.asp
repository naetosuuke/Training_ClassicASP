<HTML>
  <meta charset="UTF-8">
  <BODY>


    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338735(v=msdn.10)">
ユーザーからの入力を処理する(構造体)</a><br/>
<br/>

  ASP の Request オブジェクトを使用すると、HTML フォームから取得したデータの収集および<br/>
  処理を行う強力なスクリプトを簡単に作成できます。<br/>
 <br/>

<FORM METHOD="Post" ACTION="Profile.asp"> <!--Actionに入っている.aspファイルにデータが送信される。-->
<INPUT TYPE="Text" NAME="FirstName">
<INPUT TYPE="Text" NAME="LastName">
<INPUT TYPE="Text" NAME="Age">
<INPUT TYPE="Hidden" NAME="UserStatus" VALUE="New">
<INPUT TYPE="Submit" VALUE="Enter">
</FORM>

フォームの入力内容を取得する<br/>
<br/>

ASP の Request オブジェクトは、URL 要求として一緒に送信されるフォーム情報の取得作業を容易にする<br/>
 2 つのコレクションを提供しています。<br/>
<br/>

QueryString コレクション<br/>
<br/>

QueryString コレクションは、要求の URL の疑問符 (?) に続く文字列の形で Web サーバーに渡された<br/>
フォームの値を取得します。フォームの値を URL に追加するには、HTTP GET メソッドを使用する方法と、<br/>
URL にフォームの値を手動で追加する方法があります。<br/>
<br/>

たとえば、前のフォームの例で GET メソッド (METHOD="GET") を使用したとします。<br/>
ユーザーが「Clair」、「Hector」、「30」と入力すると、次の URL 要求がサーバーに送信されます。<br/>
<br/>

http://Workshop1/Painting/Profile.asp?FirstName=Clair&LastName;=Hector&Age;=30&UserStatus;=New<br/>
<br/>


※この場合、Web サーバーは、ユーザーの Web ブラウザに次のテキストを返します。<br/>
Hello Clair Hector. You are 30 years old! This is your first visit to this Web site!<br/>
<br/>

Profile.asp には、次のようなフォーム処理スクリプトを記述できます。<br/>
<br/>

Hello <%= Request.QueryString("FirstName") %> <%= Request.QueryString("LastName") %>.
You are <%= Request.QueryString("Age") %> years old!
<%
  If Request.QueryString("UserStatus") = "New" Then
    Response.Write "This is your first visit to this Web site!"
  End if	
%>


複数の項目を持つリスト ボックスを含むフォームでは、次のような要求が作成される場合があります。<br/>
http://OrganicFoods/list.asp?Food;=Apples&Food;=Olives&Food;=Bread<br/>
<br/>

次のコマンドで、複数の値を数えることができる。<br/>
<% 
Dim foodCount : foodCount = Request.QueryString("food").Count
Response.Write foodCount
 %><br/>
<br/>

List.asp に次のスクリプトを記述しておくと、複数の値の種類を表示できます。<br/>
<br/>

<%
  lngTotal = Request.QueryString("Food").Count
  For i = 1 To lngTotal
    Response.Write Request.QueryString("Food")(i) & "<BR>"
  Next
%><br/>
<br/>


出力結果。<br/>
<br/>

Apples  <br/>
Olives  <br/>
Bread  <br/>
<br/>

値の一覧をカンマで区切られた文字列として表示することも可能<br/>
<br/>
<% Response.Write Request.QueryString("Item") %><br/>
<br/>

出力結果<br/> 
<br/>
 
Apples, Olives, Bread<br/>
<br/>

Form コレクション<br/>
<br/>

GETのクエリ文字列は長すぎるとサーバーによっては切り捨てられる。
データが大きい場合はPOSTメソッドを使用
取得方法はRequest.Form()<br/>
<br/>

<%
  lngTotal = Request.Form("Food").Count
  For i = 1 To lngTotal
   Response.Write Request.Form("Food")(i) & "<BR>"
  Next
%><br/>
<br/>


フォームの入力内容を送信する前に、クライアント側でバリデーションを確認する。<br/>

<br/>

<SCRIPT LANGUAGE="JScript"> //スクリプトタグ デフォルト値はクライアント起動
// ユーザーが入力したアカウント番号が数字かどうかを判別するスクリプト

function CheckNumber()
{			
 if (isNumeric(document.UserForm.AcctNo.value))
   return true
 else
 {
   alert("Please enter a valid account number.")
   return false
 }		
}
	
//Function for determining if form value is a number.
//Note:　JScript の　isNaN method でも確認できるが、ブラウザによっては対応していないので
//下記メソッドを自作

function isNumeric(str)
{
  for (var i=0; i < str.length; i++)
		{
    var ch = str.substring(i, i+1) //substring 文字列の切出　引数1が開始地点、引数２が切り出す文字数
    if( ch < "0" || ch>"9" || str.length == null)
				{
      return false
    }
  }
  return true
}	
</SCRIPT>


<FORM METHOD="Get" ACTION="balance.asp" NAME="UserForm" ONSUBMIT="return CheckNumber()">
<INPUT TYPE="Text"   NAME="AcctNo"><br/> <!--入力フォーム-->
<INPUT TYPE="Submit" VALUE="Submit"><br/> <!--送信ボタン-->
</FORM>

入力検査スクリプトがデータベースへのアクセスを必要とする場合は、サーバー側でスクリプトを走らせる<br/>
HTMLフォームを持つ、同じaspファイル内でバリデーションを行えば、フォームの有用性、応答を著しく向上できる<br/>
(同ページ内でリクエストの回答が返ってくる)<br/>
<br/>

<%
  strAcct = Request.Form("Account")
  If Not AccountValid(strAcct) Then
    ErrMsg = "<FONT COLOR=Red>Sorry, you may have entered an invalid account number.</FONT>"
  Else
    'Process the user input
    Server.Transfer("Complete.asp") 'ページ遷移
  End If

  Function AccountValid(strAcct)
    'DB接続、バリデーションを行い、Boolを返すスクリプトを記述
  End Function
%>

<FORM METHOD="Post"  ACTION="Verify.asp"> <!--送信先は自分のポスト-->
Account Number:  <INPUT TYPE="Text" NAME="Account"> <%= ErrMsg %> <BR>
<INPUT TYPE="Submit"><br/>
</FORM>

JScriptでサーバー側でバリデーションを行う場合は、Requestコレクションに空のかっこ()をつけて値を渡す<br/>
そうしないとStrでなくObj型で値を返すから<br/>

<Script language=JScript runat=server>
   var Name = Request.Form("Name")();
   var Password = Request.Form("Password")();

  if(Name > "")
  {
     if(Name == Password)
      Response.Write("Your name and password are the same.")
  else
      Response.Write("Your name and password are different.");
  }
</Script>


VBScript,JScriptのどちらでも、リクエストが複数の値を持っている場合、インデックスを指定しないといけない<br/>
<br/>
<%
'var Name = Request.Form("Name")(1);
%>


</BODY>
</HTML>

