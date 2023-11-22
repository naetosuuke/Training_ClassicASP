<%@ TRANSACTION=Required %>
トランザクションディレクティブを追加すると、ページ内の処理がトランザクションとして宣言される<br>
<br>

<HTML>
  <meta charset="UTF-8">
  <BODY>


    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338727(v=msdn.10)">トランザクションを理解する(構造体)</a><br/>
<br/>

"トランザクション" は、複数の処理のすべてが成功しなければすべて元に戻すように扱われる操作です。<br>
トランザクション処理は、データベースを更新する際の信頼性を高めるために使用されます。<br>
<br>

値に Required が指定された @TRANSACTION ディレクティブを含んだスクリプトが、 <br>
Server.Transfer メソッドまたは Server.Execute メソッドで呼び出された場合<br>
 呼び出し元の .asp ファイルのトランザクションが実行されている場合は、<br>
スクリプトは呼び出し元の .asp ファイルのトランザクションをそのまま続行します。<br>


&lt;%	<br>
  'End transaction.<br>
  Server.Transfer("/BookSales/EndTrans.asp")		<br>
%><br>

&lt;%<br>
  'Instantiate a custom component to close transactions.<br>
  Set objSale = Server.CreateObject("SalesTransacted.Complete")<br>
%><br>

トランザクションが中止された場合、コンポーネント サービスはトランザクションを<br>
サポートするリソースに行われた変更をすべてロールバックします。<br>
現在、データベース サーバーのみがトランザクションを完全にサポートしています。<br>
<br>

スクリプトでは、ObjectContext.SetAbort を呼び出して、トランザクションの中止を明示的に宣言できます。<br>
<br>


スクリプト自体は、トランザクションが成功したか失敗したかを判別できません。ただし、<br>
トランザクションがコミットまたは中止されるときに呼び出されるイベントを作成することはできます。<br>
<br>

&lt;%<br>
  'Buffer output so that different pages can be displayed.<br>
  Response.Buffer = True<br>
<br>
  'FALSE･･･バッファに格納しません。<br>
  'TRUE･･･現在のページの ASP スクリプトの処理がすべて完了するまで、<br>
  'あるいは Flush メソッドまたは End メソッドが呼び出されるまで、<br>
  'クライアントに出力を送信しません。(規定値)<br>
%><br>
<br>
<HTML>
  <BODY>
  <H1>Welcome to the online banking service</H1><br>
<br>

  &lt;%<br>
    Set BankAction = Server.CreateObject("MyExample.BankComponent")<br>
    BankAction.Deposit(Request("AcctNum"))<br>
  %><br>
<br>
  <P>Thank you.  Your transaction is being processed.</P><br>
  </BODY>
</HTML>

&lt;%<br>
  'トランザクションが成功したら呼び出す<br>
  Sub OnTransactionCommit()<br>
%><br>
  <HTML>
    <BODY>

    Thank you.  Your account has been credited.<br>

    </BODY>
  </HTML>

&lt;%<br>
  Response.Flush() 'バッファしている処理をクライアントに返す<br>
  End Sub<br>
%><br>
<br>

&lt;%<br>
  'トランザクションが失敗したら呼び出す<br>
  Sub OnTransactionAbort()<br>
    Response.Clear() 'バッファしている処理を消去する<br>
%>		<br>
  <HTML>
    <BODY>

    We are unable to complete your transaction.<br>
    <br>

    </BODY>
  </HTML>
  
&lt;%<br>
    Response.Flush()<br>
  End Sub<br>
%><br>
<br>

コンポーネントをトランザクションで利用するには、COM+ アプリケーションにコンポーネントを登録し、<br>
トランザクションを要求するように構成する必要があります。<br>
<br>
トランザクション コンポーネントを登録し、構成するには、コンポーネント サービス スナップインを使用します。<br>
コンポーネントは、COM+ アプリケーションに登録する必要があります。<br>

コンポーネント サービス スナップインを使用するには、<br>
コンポーネント サービスがインストールされている必要があります。<br>
※このあたり理解できない。別途コンポーネントの管理、使用を覚える必要がある。<br>
<br>

ASP スクリプトは、宣言されたトランザクションのルート (開始点) です。<br>
トランザクション ASP ページで使用されるすべての COM オブジェクトは、<br>
トランザクションの一部と見なされます。トランザクションが完了すると、ページで使用された<br>
 COM オブジェクトは、Session オブジェクトまたは Application オブジェクトに格納された<br>
 オブジェクトも含め、すべて非アクティブ化されます。<br>
<br>






</BODY>
</HTML>

