<HTML>
  <meta charset="UTF-8">
  <BODY>


    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338735(v=msdn.10)"></a>ユーザーからの入力を処理する(構造体)<br/>
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

たとえば、複数の項目を持つリスト ボックスを含むフォームでは、次のような要求が作成される場合があります。<br/>
http://OrganicFoods/list.asp?Food=Apples&Food;=Olives&Food;=Bread<br/>
<br/>

この場合、次のコマンドを使用すると、複数の値を数えることができます。<br/>
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

このスクリプトの出力結果は次のようになります。<br/>
<br/>

Apples  <br/>
Olives  <br/>
Bread  <br/>
<br/>


次のスクリプトを使用すると、すべての値の一覧をカンマで区切られた文字列として表示することもできます。<br/>
<br/>
<% Response.Write Request.QueryString("Item") %><br/>
<br/>

この出力結果はつぎの文字列になります。<br/> 
<br/>
 
Apples, Olives, Bread<br/>
<br/>

Form コレクション<br/>















  </BODY>
</HTML>

