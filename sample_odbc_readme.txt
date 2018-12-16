MySQL��VBA�ŘA�g������EXCEL�t�@�C����ɓ]�ڂ���v���O�����B
(ODBC�ڑ��v���O�������C���X�g�[�����Ă���(mysql-connector-odbc-8.0.13-win32.msi))

���T���v���ł́A�uSERVER=localhost�v���́uDATABASE=nyuumon�v�ɁA�uUID=user01�v�Ƃ������[�U�[���Ń��O�C�����A
SQL���i"SELECT * FROM tbl_employee"�j�����s���AEXCEL��̃Z���ɓ\��t���Ă���B 

�ȉ��A�R�[�h�𔲐�����


[Module1]

Sub �f�[�^�ꗗ�\��()
  Dim adoCon As Object ' ADO�R�l�N�V����
  Dim adoRs As Object  ' ADO���R�[�h�Z�b�g
  Dim SQL As String    ' SQL
  Dim i As Long
  Dim sheetobj As Worksheet
  
  Set sheetobj = ThisWorkbook.Worksheets("sheet2")
   
  ' ADO�R�l�N�V�������쐬
  Set adoCon = CreateObject("ADODB.Connection")
  
  ' ODBC�ڑ�
   adoCon.Open _
      "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
      " SERVER=localhost;" & _
      " DATABASE=nyuumon;" & _
      " UID=user01;" & _
      " PWD=pass01;"
      
  'MsgBox "�f�[�^�ꗗ��\�����܂�"
  
  ' SQL��
  SQL = "SELECT * FROM tbl_employee"
 
  ' SQL�̎��s
  Set adoRs = adoCon.Execute(SQL)
  
  ' ���R�[�h�Z�b�g���̑S�Ă̍s�̓Ǎ���
  ' �I������܂ŏ������J��Ԃ�
  With sheetobj
  i = 1
  Do Until adoRs.EOF
    .Cells(i, 1) = adoRs!code
    .Cells(i, 2) = adoRs!Name
    .Cells(i, 3) = Format(adoRs!birthday, "yyyy/mm/dd")
    
    i = i + 1
    ' ���̍s�Ɉړ�����
    adoRs.MoveNext
  Loop
  End With
  
  ' �������
  adoRs.Close
  adoCon.Close
  Set adoRs = Nothing
  Set adoCon = Nothing
End Sub
 

[Module2]

Sub �V�K�ǉ�()
  Dim adoCon As Object ' ADO�R�l�N�V����
  Dim adoRs As Object  ' ADO���R�[�h�Z�b�g
  Dim SQL As String    ' SQL
  Dim RecordsAffected As Long  ' �ύX���ꂽ�s��
  Dim i As Long
  Dim result As Long
  
  ' �萔
  Const adExecuteNoRecords = &H80
  
  Set sheetobj = ThisWorkbook.Worksheets("sheet2")
   
  ' ADO�R�l�N�V�������쐬
  Set adoCon = CreateObject("ADODB.Connection")
  
  ' ODBC�ڑ�
   adoCon.Open _
      "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
      " SERVER=localhost;" & _
      " DATABASE=nyuumon;" & _
      " UID=user01;" & _
      " PWD=pass01;"
      
  result = MsgBox("�{���Ƀf�[�^��ǉ����Ă������ł����H", vbYesNo + vbExclamation + vbDefaultButton2)
  
  If result = vbYes Then
  ' SQL��
  SQL = "insert into  tbl_employee values(110,'�ԁ@���u', '1988-03-03', 30, 4, 101);"
        
  
  ' SQL�̎��s
  ' adExecuteNoRecords�͍s��Ԃ��Ȃ��̂Ńp�t�H�[�}���X������
  adoCon.Execute SQL, RecordsAffected, adExecuteNoRecords
  
  ' RecordsAffected�ɂ͕ύX���ꂽ�s�����Ԃ����
  Debug.Print "�ύX���ꂽ�s��:" & CStr(RecordsAffected) & "�s"
  
  End If
  
End Sub

      