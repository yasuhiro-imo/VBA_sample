/*
************************************************************
MySQLをVBAで連携させてEXCELファイル上に転載するプログラム。
(ODBC接続プログラムをインストールしておく(mysql-connector-odbc-8.0.13-win32.msi))

当サンプルでは、「SERVER=localhost」内の「DATABASE=nyuumon」に、「UID=user01」というユーザー名でログインし、
SQL文（"SELECT * FROM tbl_employee"）を実行し、EXCEL上のセルに貼り付けている(Module1)。 
又、insert文を使ってマクロからデータベースに情報を追加している(Module2)。
************************************************************
*/


'[Module1]

Sub データ一覧表示()
  Dim adoCon As Object ' ADOコネクション
  Dim adoRs As Object  ' ADOレコードセット
  Dim SQL As String    ' SQL
  Dim i As Long
  Dim sheetobj As Worksheet
  
  Set sheetobj = ThisWorkbook.Worksheets("sheet2")
   
  ' ADOコネクションを作成
  Set adoCon = CreateObject("ADODB.Connection")
  
  ' ODBC接続
   adoCon.Open _
      "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
      " SERVER=localhost;" & _
      " DATABASE=nyuumon;" & _
      " UID=user01;" & _
      " PWD=pass01;"
      
  'MsgBox "データ一覧を表示します"
  
  ' SQL文
  SQL = "SELECT * FROM tbl_employee"
 
  ' SQLの実行
  Set adoRs = adoCon.Execute(SQL)
  
  ' レコードセット内の全ての行の読込が
  ' 終了するまで処理を繰り返す
  With sheetobj
  i = 1
  Do Until adoRs.EOF
    .Cells(i, 1) = adoRs!code
    .Cells(i, 2) = adoRs!Name
    .Cells(i, 3) = Format(adoRs!birthday, "yyyy/mm/dd")
    
    i = i + 1
    ' 次の行に移動する
    adoRs.MoveNext
  Loop
  End With
  
  ' 解放処理
  adoRs.Close
  adoCon.Close
  Set adoRs = Nothing
  Set adoCon = Nothing
End Sub
 

'[Module2]

Sub 新規追加()
  Dim adoCon As Object ' ADOコネクション
  Dim adoRs As Object  ' ADOレコードセット
  Dim SQL As String    ' SQL
  Dim RecordsAffected As Long  ' 変更された行数
  Dim i As Long
  Dim result As Long
  
  ' 定数
  Const adExecuteNoRecords = &H80
  
  Set sheetobj = ThisWorkbook.Worksheets("sheet2")
   
  ' ADOコネクションを作成
  Set adoCon = CreateObject("ADODB.Connection")
  
  ' ODBC接続
   adoCon.Open _
      "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
      " SERVER=localhost;" & _
      " DATABASE=nyuumon;" & _
      " UID=user01;" & _
      " PWD=pass01;"
      
  result = MsgBox("本当にデータを追加してもいいですか？", vbYesNo + vbExclamation + vbDefaultButton2)
  
  If result = vbYes Then
  ' SQL文
  SQL = "insert into  tbl_employee values(110,'車　高志', '1988-03-03', 30, 4, 101);"
        
  
  ' SQLの実行
  ' adExecuteNoRecordsは行を返さないのでパフォーマンスが向上
  adoCon.Execute SQL, RecordsAffected, adExecuteNoRecords
  
  ' RecordsAffectedには変更された行数が返される
  Debug.Print "変更された行数:" & CStr(RecordsAffected) & "行"
  
  End If
  
End Sub

      
