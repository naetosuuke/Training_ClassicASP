<HTML>
  <meta charset="UTF-8">
  <BODY>


    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338742(v=msdn.10)">
データ ソースにアクセスする</a><br/>
<br/>

<h2>※実際のDBへのコネクトは失敗している。別ファイルでASPファイルからの接続を試みる</h2>

ActiveX Data Objects (ADO) は、Web ページにデータベース アクセス機能を追加するための、<br>
使いやすく拡張性の高いテクノロジです。<br>
<br>

接続文字列を作成する<br>
<br>
ADO は接続文字列を使用して OLE DB "プロバイダ" を識別し、そのプロバイダをデータ ソースに結び付けます<br>
<br>

データ ソース: Microsoft SQL Server<br>
OLE DB 接続文字列: Provider=SQLOLEDB.1;Data Source=(DBのパス)<br>


Web データ アプリケーションの設計時に考慮すべき重要な問題<br>
<br>

リモート SQL Server データベースに接続する ASP データベース アプリケーションを開発する場合は、<br>
次の点にも注意する必要があります。<br>

SQL Server の接続方式の選択<br>
TCP/IP ソケット = SQL Serverによる認証<br>
名前付きパイプ = Windwsによる認証<br>
<br>

ODBC 80004005 エラー<br>
SQL Server セキュリティ<br>
<br>


データ ソースに接続し、Connection オブジェクトを使用して SQL クエリを実行する<br>
<br>

  <%
  'Define the OLE DB connection string.
  strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\TutorialDB.mdf"

  'Instantiate the Connection object and open a database connection.
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.Open strConnectionString

  'Define SQL SELECT statement.
  strSQL = "INSERT INTO dbo.Customers (Name, Location) VALUES ('Take','Japan')"

  '追加
  cnn.Execute strSQL,,adCmdText + adExecuteNoRecords

  '上書き
  cnn.Execute "UPDATE dbo.Customers SET Name = 'YOYO' WHERE Location = 'Japan' ",,adCmdText + adExecuteNoRecords

  '削除
  cnn.Execute "DELETE FROM dbo.Customers WHERE Location = 'Japan'",,adCmdText + adExecuteNoRecords


%>

クエリの実行に使用されるステートメントで、adCmdText と adExecuteNoRecords の<br>
2 つのパラメータが指定されていることに注意してください。省略可能な adCmdText パラメータは、<br>
コマンドの種類を指定することで、クエリ ステートメント (この場合は SQL クエリ) を、テキスト形式の<br>
コマンド定義として評価するようにプロバイダに指示します。adExecuteNoRecords パラメータは、<br>
アプリケーションに結果が返されない場合に、データ レコードのセットを作成しないように ADO に指示します。<br>
<br>
省略可能ですが、Execute メソッドを使用するときは、データ アプリケーションのパフォーマンスの向上のため<br>
にこれらのパラメータを使用することをお勧めします<br>
<br>


Recordset オブジェクトを使用して結果を操作する<br>
<h2>※※レコードセットオブジェクトの理解が足りない。別途組み込みオブジェクトの項で使用を確認※※</h2><br>
<br>

ADO では、データの取得、結果の検査、およびデータベースの変更に使用する <br>
Recordset オブジェクトが用意されています。次のサーバー側スクリプトでは、<br>
Recordset オブジェクトを使用して SQL の SELECT コマンドを実行します。<br>
SELECT コマンドは、クエリの制約条件に基づいて特定の情報を取得します。<br>
<br>

<%
  'DBと接続
  strConnectionString  = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\TutorialDB.mdf"
  Set cnn = Server.CreateObject("ADODB.Connection")
  cnn.Open strConnectionString

  'レコードセットオブジェクトを初期化
  Set rstCustomers = Server.CreateObject("ADODB.Recordset")

  'レコードセットを、SQL文とコネクションのインスタンスを使って開ける
  strSQL = "SELECT * FROM dbo.Customers WHERE Location = 'Japan' "
  rstCustomers.Open  strSQL, cnn	'

  '取得したレコードのうち、条件を満たすレコードのLocationの値を出力
   Set objLocation = rstCustomers("Location")
     Do Until rstCustomers.EOF 'このEOFプロパティは最後のデータという意味？
     Response.Write objLocation & "<BR>"
     rstCustomers.MoveNext
   Loop
%>

この例では、Connection オブジェクトがデータベース接続を確立し、Recordset オブジェクトが<br>
この接続を使用してデータベースから結果を取得しています。この方法は、データベースとのリンクの<br>
立方法を正確に構成する必要がある場合に有効です。たとえば、接続の試みを中止するまでの<br>
待機時間を指定する場合は、そのためのプロパティを Connection オブジェクトを使用して設定する必要があります。<br>
<br>

<%
  strConnectionString  = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\TutorialDB.mdf"
  strSQL = "SELECT FirstName, LastName FROM Customers WHERE LastName = 'Smith' "
  Set rstCustomers = Server.CreateObject("ADODB.Recordset")

  'Open a connection using the Open method
  'and use the connection established by the Connection object.
  rstCustomers.Open  strSQL, strConnectionString	

  'Cycle through the record set, display the results,
  'and increment record position with MoveNext method.
   Set objLocation = rstCustomers("Location")
   Do Until rstCustomers.EOF
      Response.Write objLocation & "<BR>"
      rstCustomers.MoveNext
   Loop
%>


ADO の既定の接続プロパティを使用して接続を確立するだけであれば、<br>
Recordset オブジェクトの Open メソッドを使用することもできます。<br>
Recordset オブジェクトの Open メソッドを使用して接続を確立する場合も、<br>
リンクを確保するために Connection オブジェクトを暗黙的に使用しています。<br>
<br>

&lt;% 'サンプル　レコードセット内の正確なレコード数をカウント<br>
	Set rs = Server.CreateObject("ADODB.Recordset")<br>
	rs.Open "SELECT * FROM NewOrders", "Provider=Microsoft.Jet.OLEDB.3.51;Data Source='C:\CustomerOrders\Orders.mdb'", adOpenKeyset, adLockOptimistic, adCmdText<br>
	<br>
	'Use the RecordCount property of the Recordset object to get count.<br>
	If rs.RecordCount >= 5 then<br>
	  Response.Write "We've received the following " & rs.RecordCount & " new orders<BR>"	
	<br>
	  Do Until rs.EOF<br>
	  	Response.Write rs("CustomerFirstName") & " " & rs("CustomerLastName") & "<BR>"
		Response.Write rs("AccountNumber") & "<BR>"
		Response.Write rs("Quantity") & "<BR>"		
		Response.Write rs("DeliveryDate") & "<BR><BR>"
	      	rs.MoveNext<br>
	  Loop<br>
<br>
  	Else<br>	    	
	  Response.Write "There are less than " & rs.RecordCount & " new orders."		<br>
	End If<br>
	
   rs.Close<br>
%>



Command オブジェクトを使用してクエリを改良する<br>
ADO の Command オブジェクトでは、Connection オブジェクトと <br>
Recordset オブジェクトを使用してクエリを実行する場合と同じようにクエリを実行できます。<br>


Web ベースの在庫システムで定期的に供給情報およびコスト情報を更新する必要がある場合は、<br>
次のようにクエリをあらかじめ定義しておくことができます<br>
<br>

&lt;% サンプル
    'Open a connection using Connection object. Notice that the Command object
    'does not have an Open method for establishing a connection.
    strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Data\Inventory.mdb"
    Set cnn = Server.CreateObject("ADODB.Connection")
    cnn.Open strConnectionString

    'コマンドオブジェクトを初期化
    'コマンドオブジェクトのActiveConnectionに接続済みのConnectionを代入して、使用できるようにする
    Set cmn= Server.CreateObject("ADODB.Command")
    Set cmn.ActiveConnection = cnn

    'クエリを定義.
    cmn.CommandText = "INSERT INTO Inventory (Material, Quantity) VALUES (?, ?)"

    'Save a prepared (or pre-compiled) version of the query specified in CommandText
    'property before a Command object's first execution.
    cmn.Prepared = True

    'クエリのパラメーターを設定(Parameters配列に?のオブジェクトが入っているので、設定を足していく)
    cmn.Parameters.Append cmn.CreateParameter("material_type",adVarChar, ,255 )
    cmn.Parameters.Append cmn.CreateParameter("quantity",adVarChar, ,255 )

    'パラメーターを投入して実行
    cmn("material_type") = "light bulbs"
    cmn("quantity") = "40"
    cmn.Execute ,,adCmdText + adExecuteNoRecords

    '違うパラメーターを投入して実行
    cmn("material_type") = "fuses"
    cmn("quantity") = "600"
    cmn.Execute ,,adCmdText + adExecuteNoRecords
    .
    .
  %>

Command オブジェクトを使用してクエリをコンパイルすると、文字列と変数を連結して <br>
SQLクエリを形成する場合に発生する問題を回避できるという利点もあります<br>
<br>
  strSQL = "INSERT INTO Customers (FirstName, LastName) VALUES ('Robert','O'Hara')"<br>
この例のラスト ネーム「O'Hara」に含まれる一重引用符 (') は、SQL の VALUES キーワードでデータを表すために使われる一重引用符と競合します。<br>


重要   adCmdText などの ADO パラメータは、単なる変数です。つまり、データ アクセス スクリプトで <br>
ADO パラメータを使用する前に、その値を定義しておく必要があります。ADO では多数のパラメータを使用するため、<br>
ADO のすべてのパラメータおよび定数の定義を収めたファイルである "コンポーネント タイプ ライブラリ" を使用して、<br>
パラメータを定義すると簡単です。ADO タイプ ライブラリを実装する方法については、「変数と定数を使用する」の<br>
「定数を使用する」を参照してください。<br>
  <h2>※※タイプライブラリの使用方法を覚える※※</h2><br>
  <br>



HTML フォームとデータベース アクセスを結合する<br>
<br>

<%
  'Open a connection using Connection object. The Command object
  'does not have an Open method for establishing a connection.
   strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\CompanyCatalog\Seeds.mdb"
 	Set cnn = Server.CreateObject("ADODB.Connection")
 	cnn.Open strConnectionString

  'Instantiate Command object
  'and  use ActiveConnection property to attach
  'connection to Command object.
  Set cmn= Server.CreateObject("ADODB.Command")
  Set cmn.ActiveConnection = cnn

  'Define SQL query.
  cmn.CommandText = "INSERT INTO MySeedsTable (Type) VALUES (?)"

  'Define query parameter configuration information.
  cmn.Parameters.Append cmn.CreateParameter("type",adVarChar, ,255)

  'Assign input value and execute update.
  cmn("type") = Request.Form("SeedType") 'フォームで受け取ったデータを処理することも可能
  cmn.Execute ,,adCmdText + adExecuteNoRecords
%>



データベース接続を管理する<br>
。データベース接続を開いたままにしておくと、情報のやり取りが行われていなくても、<br>
データベース サーバーのリソースが圧迫され、接続に問題が生じる可能性があります。<br>
<br>
接続のタイムアウトを設定する<br>
Connection オブジェクトの ConnectionTimeout を使用すると、アプリケーションが接続の試行を放棄して<br>
エラー メッセージを表示するまで待機する時間を設定することができます<br>

<%
Set cnn = Server.CreateObject("ADODB.Connection")
cnn.ConnectionTimeout = 20
cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Data\Inventory.mdb"
%>


接続をプールする<br>
<br>
接続をプールすることで、Web アプリケーションは "プール" つまり再確立の不要な空いている接続の<br>
集まりの中から接続を使用できます。接続が作成され、プールに配置されると、その接続は接続処理を<br>
実行しなくてもアプリケーションで再利用できます。<br>

デフォルトで60秒間接続されている。レジストリキーを作成すると、任意の時間に変更可能<br>



複数のページにまたがって接続を使用する<br>

ASP の Application オブジェクトに接続を格納すると、複数ページにわたって接続を再利用できる。<br>
ただ、ASP の Application オブジェクトにデータベース接続文字列を格納し、<br>
複数の Web ページでその接続文字列を再利用する方がアプローチとして優れています。<br>
<br>
Global.asa ファイルの Application_OnStart イベント プロシージャで接続文字列を指定できます。<br>
Application("ConnectionString") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Data\Inventory.mdb"<br>
<br>
ASPファイルではObjectタグでインスタンス生成して利用する。<br>
&lt;OBJECT RUNAT=SERVER ID=cnn PROGID="ADODB.Connection"></OBJECT><br>
cnn.Open Application("ConnectionString")<br>
cnn.Close<br>

1人のユーザーが複数ページをまたぐときは、Sessionオブジェクトを使用するのが適している
※Applicationだと他ユーザーにも影響あり


接続を閉じる<br>
接続プールを活用するには、不要になったら明示的に接続を閉じるといい<br>
リソースが空く<br>
<br>








</BODY>
</HTML>

