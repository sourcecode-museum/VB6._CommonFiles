Attribute VB_Name = "mdlADO"
Option Explicit
Public gOConn As ADODB.Connection
Public gbBancoTipoMDB As Boolean

Public Function AbrirConexao(ByVal psConn As String) As Boolean
  Static bErrou As Boolean
  
  Set gOConn = New ADODB.Connection
  On Error GoTo TrataErro
  With gOConn
    .CursorLocation = adUseServer
    .Open psConn ', "SYSDBA", "masterkey"
  End With
  AbrirConexao = True
  Exit Function
  
TrataErro:
  If Not bErrou Then
    On Error GoTo 0
    On Local Error Resume Next
    
    bErrou = True
    Dim oS As New SisFuncoes.cSisFuncoes
    oS.CompactMDB mdlConfigINI.INIConexao.PathDB
    Set oS = Nothing
    Call AbrirConexao(mdlConfigINI.INIConexao.StrConexao)
  End If
  On Error GoTo 0
  
  Err.Clear
  
  Set gOConn = Nothing
  AbrirConexao = False
  MGShowErro "mdlADO.AbrirConexao"
  End
End Function

Public Sub FecharConexao()
  If Not gOConn Is Nothing Then
    gOConn.Close
    Set gOConn = Nothing
  End If
End Sub

Public Sub AbrirRS(ByRef pRS As ADODB.Recordset, ByVal pSQL As String, _
                   Optional ByVal pCursorLocation As CursorLocationEnum = adUseServer, _
                   Optional pCursorType As CursorTypeEnum = adOpenDynamic, _
                   Optional pLockType As LockTypeEnum = adLockOptimistic, _
                   Optional ByVal pDesconectado As Boolean = False, _
                   Optional bNotVazio As Boolean = True)
 
  On Error GoTo TrataErro
  If gOConn Is Nothing Then
    Call mdlADO.AbrirConexao(mdlConfigINI.INIConexao.StrConexao)
  End If
  
  If Not pRS Is Nothing Then
    On Error Resume Next
    If pRS.Status = adStateOpen Then pRS.Close
    Set pRS = Nothing
    On Error GoTo 0
  End If
  
  Set pRS = New ADODB.Recordset
  With pRS
    .CursorLocation = pCursorLocation
    .Open pSQL, gOConn, pCursorType, pLockType, adCmdText
    
    If pDesconectado Then Set .ActiveConnection = Nothing
  End With
  
  If pRS.RecordCount = 0 And bNotVazio Then
    pRS.Close
    Set pRS = Nothing
  End If
  
  Exit Sub
TrataErro:
  Set pRS = Nothing
  Call MGShowErro("mdlADO.AbrirRS")
End Sub

Public Sub AbrirDAT(ByRef pRS As ADODB.Recordset, ByVal pFullPath As String, Optional bNotVazio As Boolean = True)
 
  On Error GoTo TrataErro
 
  If Not pRS Is Nothing Then
    On Error Resume Next
    If pRS.Status = adStateOpen Then pRS.Close
    Set pRS = Nothing
    On Error GoTo 0
  End If
  
  Set pRS = New ADODB.Recordset
  pRS.Open pFullPath, , adOpenDynamic, adLockBatchOptimistic, adCmdFile
  
  If pRS.RecordCount = 0 And bNotVazio Then
    pRS.Close
    Set pRS = Nothing
  End If
  
  Exit Sub
TrataErro:
  Set pRS = Nothing
  Call MGShowErro("mdlADO.AbrirDAT")
End Sub

Public Sub Pesquisar(ByRef pRecordset As Object, ByVal DataField As String, ByVal Texto As String)
  Dim sTexto      As String
  Dim sCriterio   As String
  
  Dim RSConsulta  As ADODB.Recordset
  
  Static sStaticField As String
  Static sStaticCriterio As String
  
  sTexto = Texto
  If sTexto = "" Then Exit Sub
  
  On Error GoTo TrataErro
  Set RSConsulta = pRecordset.Clone(adLockReadOnly)
  With RSConsulta
    
    If pRecordset.Sort <> "" Then .Sort = pRecordset.Sort
    
    DataField = Replace(DataField, "[", "") 'Apenas para evitar erros
    DataField = Replace(DataField, "]", "")
    
    Select Case RSConsulta.Fields(DataField).Type
      Case Is = adChapter, adWChar, adVarChar, adVarWChar, adChar
        sCriterio = " LIKE '" & sTexto & "%'"
        
      Case Is = adNumeric, adInteger, adSmallInt, adDouble, adBinary
        sCriterio = " = " & sTexto
        
      Case Is = adBoolean
        Dim s As String
        
        sTexto = Trim$(sTexto)
        s = UCase(Left$(sTexto, 1))
        
        Select Case s
          Case Is = "S", "V", "-", "T" 'S=Sim , V=Verdadeiro, - = -1, T = True
            sTexto = "True"
          Case Else
            sTexto = "False"
        End Select
        sCriterio = " = " & sTexto
        
      Case Is = adDate, adDBDate, adDBTime
        sCriterio = " = #" & sTexto & "#"
        
      Case Else
        sCriterio = " LIKE '" & sTexto & "%'"
    End Select
    
    DataField = "[" & DataField & "]"
    
    Dim nStartFind As Integer
    
    If sStaticField <> DataField Then
      sStaticField = DataField
      nStartFind = 0
    Else
      If pRecordset.AbsolutePosition < pRecordset.RecordCount Then
        nStartFind = pRecordset.AbsolutePosition
      Else
        nStartFind = 0
      End If
    End If
    
    If sStaticCriterio <> sCriterio Then
      sStaticCriterio = sCriterio
      nStartFind = 0
    Else
      If pRecordset.AbsolutePosition < pRecordset.RecordCount Then
        nStartFind = pRecordset.AbsolutePosition
      Else
        nStartFind = 0
      End If
    End If
    
    On Error GoTo 0
    
    On Error Resume Next
    If nStartFind > 0 Then
      .Find DataField & sCriterio, nStartFind, adSearchForward
    Else
      'Neste caso inicia a busca a partir do primeiro registro
      .Find DataField & sCriterio, , adSearchForward, 1
    End If
    
    If Err.Number <> 0 Then
      GoTo TrataErro:
    End If
  
    If Not .BOF And Not .EOF Then
      pRecordset.AbsolutePosition = .AbsolutePosition
    End If
  End With
  
  RSConsulta.Close
  On Error GoTo 0

  Set RSConsulta = Nothing
  
  Exit Sub
  
TrataErro:
  Call MGShowErro("mdlADO.Pesquisar")

  On Error Resume Next
  RSConsulta.Close
  Set RSConsulta = Nothing
  On Error GoTo 0
End Sub

Public Function AutoPosition(ByRef pRecordset As Object, ByVal DataField As String, ByVal Texto As String, KeyAscii As Integer) As Long
  Dim sTexto      As String
  Dim sCriterio   As String
  Dim RSConsulta  As ADODB.Recordset

  Static sStaticField As String
  Static sStaticCriterio As String

  If KeyAscii = vbKeyEscape Then Exit Function

  On Error Resume Next
  Select Case KeyAscii
    Case Is = vbKeyBack
      sTexto = Mid$(Texto, 1, Len(Texto) - 1)
    Case Is = vbKeyReturn
      sTexto = Texto
    Case Else
      sTexto = Texto & Chr(KeyAscii)
  End Select
  On Error GoTo 0

  If sTexto = "" Then Exit Function

  On Error Resume Next

  Call AbrirRS(RSConsulta, pRecordset.Source, adUseServer)
  If RSConsulta Is Nothing Then Exit Function

  With RSConsulta
    Select Case RSConsulta.Fields(DataField).Type
      Case Is = adChapter, adWChar, adVarChar, adVarWChar, adChar
        sCriterio = " LIKE '" & sTexto & "%'"

      Case Is = adNumeric, adInteger, adSmallInt, adDouble
        sCriterio = " = " & sTexto

      Case Is = adBinary, adBoolean
        sCriterio = " = " & sTexto

      Case Is = adDate, adDBDate, adDBTime
        sCriterio = " = '" & sTexto & "'"

      Case Else
        sCriterio = " LIKE '" & sTexto & "%'"
    End Select

    DataField = "[" & DataField & "]"

    Dim nStartFind As Integer

    If sStaticField <> DataField Then
      sStaticField = DataField
      nStartFind = 0
    Else
      nStartFind = pRecordset.AbsolutePosition
    End If

    If sStaticCriterio <> sCriterio Then
      sStaticCriterio = sCriterio
      nStartFind = 0
    Else
      nStartFind = pRecordset.AbsolutePosition
    End If

    If nStartFind > 0 Then
      .Find DataField & sCriterio, nStartFind, adSearchForward
    Else
      'Neste caso inicia a busca a partir do primeiro registro
      .Find DataField & sCriterio, , adSearchForward, 1
    End If

    If Err.Number <> 0 Then
      GoTo TrataErro:
    End If

    If Not .BOF And Not .EOF Then
      pRecordset.AbsolutePosition = .AbsolutePosition
      .Close
    End If
  End With
  Set RSConsulta = Nothing

  On Error GoTo 0
  Exit Function

TrataErro:
  Call MGShowErro("mdlADO.AutoPosition")
End Function

''' Antigo AutoSearch
Public Sub AutoComplete(ByRef poTextBox As Control, ByVal psDataField As String, _
                        ByVal psTabela As String, ByRef KeyAscii As Integer)
                           
  Dim RS      As ADODB.Recordset
  Dim sBuffer As String, sTexto As String
  Dim sSearch As String, sSQL   As String
  
  If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then Exit Sub

  On Error Resume Next
  If KeyAscii = vbKeyBack Then
    sTexto = poTextBox.Text
    If sTexto = "" Then Exit Sub
    sTexto = Mid$(Len(sTexto), 1, Len(sTexto) - 1)
    If sTexto = "" Then Exit Sub
  End If
  On Error GoTo 0
  
  sTexto = poTextBox.Text & Chr(KeyAscii)
  sBuffer = Left(sTexto, poTextBox.SelStart) & Chr(KeyAscii)
  
  psDataField = "[" & psDataField & "]"
  sSearch = psDataField & " LIKE '" & sBuffer & "%'"
  
  sSQL = "SELECT " & psDataField & " FROM " & psTabela & " WHERE " & sSearch
  Call AbrirRS(RS, sSQL)
  If Not RS Is Nothing Then
    poTextBox.Text = RS.Fields(0)
    poTextBox.SelStart = Len(sBuffer)
    poTextBox.SelLength = Len(poTextBox.Text)
    KeyAscii = 0
    RS.Close
    Set RS = Nothing
  End If
  
End Sub

Public Function UpdateValores(ByVal psTabela As String, _
                              ByVal pTextArrayCMP As String, _
                              ByVal pTextArrayVAL As String, _
                              ByVal pSQLWhere As String) As Boolean
                              
  Dim SQL As String, SQLValores As String
  Dim aCampos() As String, aValores() As String
  Dim i As Integer
  
  On Error GoTo TrataErro
  aCampos = Split(pTextArrayCMP, "|")
  aValores = Split(pTextArrayVAL, "|")
  
  For i = 0 To UBound(aCampos)
    SQLValores = SQLValores & ", " & aCampos(i) & " = " & aValores(i)
  Next
  
  SQL = "UPDATE " & psTabela & " SET " & Mid(SQLValores, 3) & " WHERE " & pSQLWhere
  gOConn.Execute SQL
  UpdateValores = True
  Exit Function
TrataErro:
  Call MGShowErro("mdlADO.UpdateValores")

End Function

Public Function IDExiste(ByVal pID As Variant, ByVal pTabela As String)
  Dim RS As ADODB.Recordset
  Dim SQL As String
  
  If pID = 0 Or pID = "" Then Exit Function
  
  SQL = "SELECT ID FROM " & pTabela & " WHERE ID = " & pID
  Call AbrirRS(RS, SQL)
  If Not RS Is Nothing Then
    IDExiste = True
    RS.Close
    Set RS = Nothing
  End If
End Function

Public Function Posicionar(ByRef pObjRecordSet As Object, ByVal pID As Integer) As Boolean
   Dim nPos As Integer
   
   If pObjRecordSet.RecordCount <> 0 Then
      nPos = pObjRecordSet.AbsolutePosition
      pObjRecordSet.Find "ID = " & pID, , adSearchForward, 1
      
      Posicionar = Not pObjRecordSet.EOF
      If pObjRecordSet.EOF Then
         pObjRecordSet.AbsolutePosition = nPos
      End If
   End If
End Function

Public Function MaxValue(ByVal psTabela, ByVal psCampo As String, Optional ByVal psWhere As String) As Variant
On Error GoTo Finaliza
    Dim oRS   As New ADODB.Recordset
    Dim sSQL  As String
    
    sSQL = "SELECT MAX(" & psCampo & ") as MaxCod From " & psTabela
    If psWhere <> "" Then sSQL = sSQL & " WHERE " & psWhere
    Set oRS = gOConn.Execute(sSQL)
    
    If Not IsNull(oRS!MaxCod) Then
      MaxValue = oRS!MaxCod
    Else
      MaxValue = Null
    End If
  
Finaliza:
    If Err.Number <> 0 Then
      Call mdlGeral.MGShowErro("mdlADO.MaxValue")
    End If
    
    On Local Error Resume Next
    If oRS.State = adStateOpen Then oRS.Close
    Set oRS = Nothing
End Function

Public Function GerarCodSequencial(ByVal pNome As String, ByVal pIdentificador As String, _
                                   Optional ByVal pMask As String = "0", _
                                   Optional ByVal pMinCodSeq As Long) As String
On Error GoTo Finaliza
    Dim oRS     As New ADODB.Recordset
    Dim sSQL    As String
    Dim sWhere  As String
    Dim nReturn As Long
    
    sWhere = "SEQ_NOME = '%SEQNAME' AND SEQ_IDENTIFICADOR = '%SEQFIELD'"
    sWhere = Replace(sWhere, "%SEQNAME", Replace(pNome, "'", "''"))
    sWhere = Replace(sWhere, "%SEQFIELD", Replace(pIdentificador, "'", "''"))
    
    sSQL = "SELECT MAX(SEQ_VALOR) AS MaxCod FROM SIS_SEQUENCIA WHERE " & sWhere
    
    Set oRS = gOConn.Execute(sSQL)
    
    nReturn = IIf(IsNull(oRS!MaxCod), 1, Val(oRS!MaxCod) + 1)
    If pMinCodSeq > nReturn Then nReturn = pMinCodSeq
    
    Call UpdateValores("SIS_SEQUENCIA", "SEQ_VALOR", nReturn, sWhere) 'SALVA A PROXIMA SEQUENCIA
    
    GerarCodSequencial = Format$(nReturn, pMask)
      
Finaliza:
    If Err.Number <> 0 Then
      Call mdlGeral.MGShowErro("mdlADO.GerarCodSequencial")
    End If
    
    On Local Error Resume Next
    If oRS.State = adStateOpen Then oRS.Close
    Set oRS = Nothing
End Function

Public Function ExecQuery(SQL As String, ByRef RecordsAfetados As Integer) As Boolean
  On Error GoTo TrataErro:
  If gOConn Is Nothing Then
    Call mdlADO.AbrirConexao(mdlConfigINI.INIConexao.StrConexao)
  End If

  gOConn.Execute SQL, RecordsAfetados
  ExecQuery = True
  Exit Function
TrataErro:
  MGShowErro ("mdlADO.ExecQuery")
End Function

Public Function BuscaValores(ByVal pSQL As String, Optional ByVal psDelimitador As String = " - ") As String
  Dim RS As ADODB.Recordset
  Dim i As Integer, s As String
  
  Call AbrirRS(RS, pSQL, adUseServer, adOpenForwardOnly, adLockOptimistic, True, True)

  If Not RS Is Nothing Then
    For i = 0 To RS.Fields.Count - 1
      s = s & RS.Fields(i).value & psDelimitador
    Next
    BuscaValores = Mid$(s, 1, Len(s) - Len(psDelimitador))
    
    RS.Close
    Set RS = Nothing
  End If
End Function

