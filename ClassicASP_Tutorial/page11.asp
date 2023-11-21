<HTML>
  <meta charset="UTF-8">
  <BODY>


    <br/>
<a href="https://learn.microsoft.com/ja-jp/previous-versions/iis/iis-5.0/cc338736(v=msdn.10)">ファイルをインクルードする</a><br/>
<br/>

サーバー側インクルードディレクティブ<br>
サーバー内の他ファイルを呼び出せる<br>
<br>

キーワード virtual とキーワード file は、ファイルをインクルードするときに使うパスの種類を表し、<br>
filename はインクルードするファイルのパスと名前を示します。<br>
<br>

キーワード Virtual を使用する<br>
"仮想ディレクトリ" で始まるパスを指定するときに使用します。<br>
<br>

<h2>※※ディレクトリ内に作成したincファイル 私の環境だとデフォルトではページ使用ができない(404.3)※※</h2>
<h2>※※解決に時間がかかりそうだったので、使用できまいまま次に進む※※</h2>



&lt;!-- #include virtual ="/include/test.inc" --><br>
<br>

キーワード File を使用する<br>
<br>

"相対" パスを指定するときに使用します。相対パスでは、インクルードする側のファイルのあるディレクトリがパスの起点になります。<br>
ＩＩＳスナップインで [親のパスを有効にする] オプションをオンにしている場合は、キーワード file を構文 (..\) と共に使用して、<br>
親ディレクトリ (1 つ上のディレクトリ) のファイルをインクルードすることもできる<br>
&lt;!-- #include file ="test.inc" -->
<br>

インクルードされる側のファイルが、ほかのファイルをインクルードすることもできます。<br>
.asp ファイルでは、同じファイルを何度もインクルードすることができます<br>
だし、これは #include ディレクティブがループを発生させない場合に限ります。<br>
<br>

<b>ASP は、スクリプト コマンドを実行する前にファイルをインクルードします。</b><br>
したがって、スクリプト コマンドを使って、インクルードされるファイルの名前を作成することはできません。<br>
<br>

<!--  This script will fail 変数をinclude名に使用しているため-->
&lt;% name=(header1 & ".inc") %>
&lt;!-- #include file="<%= name %>" -->


<!-- This script will fail スクリプト内でもダメ(コンパイル前だから)-->
&lt;% 
  For i = 1 To n
    statements in main file
    &lt!--  #include file="header1.inc" -->
  Next
%><br>

<!-- スクリプト外に逃がしていればOK-->

&lt;%
  For i = 1 to n
    statements in main file
%>
&lt!--  #include file="header1.inc"   -->
&lt;% Next %><br>

インクルードファイルは細かく分けて使う方がいい<br>
大きい1つのファイルで運用すると、必要ない構造体でリソースがなくなるから。<br>
<br>

スクリプトタグを使用してもインクルードできる。<br>
&lt;SCRIPT LANGUAGE="VBScript" RUNAT=SERVER SRC="Utils\datasrt.inc"></SCRIPT><br>
<br>

仮想、相対パスで構文が違う。<br>
パスの種類	構文	例<br>
相対	SRC="Path\Filename"	SRC="Utilities\Test.asp"<br>
仮想	SRC="/Path/Filename"	SRC="/MyScripts/Digital.asp"<br>
仮想	SRC="\Path\Filename"	SRC="\RegApps\Process.asp"<br>
<br>






</BODY>
</HTML>

