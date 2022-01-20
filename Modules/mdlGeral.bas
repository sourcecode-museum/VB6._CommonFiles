Attribute VB_Name = "mdlGeral"
Option Explicit
Global goArqINI As SisFuncoes.cArqINI
Private mColShellApp As Collection 'Colecao de nWnd dos Aplicativos Externos abertos

'=======#Inicio Declarações de API

'Manter os Aplicativos externos como Filhos do Aplicativo
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_HWNDNEXT = 2
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

'Sendo usada para fechar as janelas dos Aplicativos externos
'E para Drag em Forms
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_CLOSE = &H10
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'Usado para pegar a posicao do objeto
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'Sempre no Topo
Declare Function SetWindowPos Lib "user32" (ByVal h&, ByVal hb&, ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal f&) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
'=======#Fim Declarações de API


'Colocando os Controles em modo de Flat
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Consts used for Flat effects
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const BS_HOLLOW = 0

' Window Consts
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
'Private Const SWP_NOSIZE = &H1
'Private Const SWP_NOMOVE = &H2
'==========Fim Flat

Enum geDataPadrão
  [DT Atual] = 0
  [DT IniMês] = 1
  [DT FimMês] = 2
End Enum

' ******************** INICIO **********************
' COPIADO TUDO PARA mdlConfigINI
'***********************************************
'Type tEmpresa
'  NomeEmpresa    As String
'  IMGFundo       As String
'  IMGLogoMarca   As String
'End Type
'
'Private mtEmpresa As tEmpresa
'
Public gsUsuario            As String
Public gsEmpresa            As String
'
'Public Property Get Empresa() As tEmpresa
'   Empresa = mtEmpresa
'End Property
'Public Property Let Empresa(vNewValue As tEmpresa)
'   mtEmpresa = vNewValue
'End Property
'
'Public Sub LerInfoEmpresa()
'   Dim VarEmpresa As tEmpresa
'
'   Set goArqINI = New cArqINI
'   With goArqINI
'
'      .PathFile = mdlConfigINI.gsPathINI
'      VarEmpresa.NomeEmpresa = .Ler("SISTEMA", "NomeEmpresa", "Contato:heliomarpm@hotmail.com ")
'      VarEmpresa.IMGFundo = .Ler("SISTEMA", "IMGFundo", "Fundo.JPG")
'      VarEmpresa.IMGLogoMarca = .Ler("SISTEMA", "IMGLogo", "Logo.JPG")
'      Empresa = VarEmpresa
'   End With
'   Set goArqINI = Nothing
'End Sub

' ****************** FIM **********************

Public Sub EnterTab(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If Shift = vbShiftMask Then
      SendKeys "+{TAB}"
      KeyCode = 0
    ElseIf Shift = 0 Then
      SendKeys "{TAB}"
      KeyCode = 0
    End If
  End If
End Sub

'Função para abreviar nomes Proprios
'Ex: "Heliomar Pereira Marques dos Santos" => "Heliomar P. M. Dos Santos"
Public Function AbreviaNome(psNome As String, Optional TCase As VbStrConv = vbUpperCase)
  Dim aNomes() As String
  Dim i As Integer, nCount As Integer
  Dim sTemp As String
  
  psNome = Trim$(psNome)
  
  aNomes() = Split(psNome, " ") 'Array é de Base Zero
  nCount = UBound(aNomes)
  'Abreviar a partir do segundo nome, exceto o último.
  If nCount > 1 Then
    sTemp = aNomes(0) & Chr(32)
    For i = 1 To nCount - 1
      'Contém mais de 3 letras? (ignorar de, da, das, do, dos, etc.)}
      If Len(aNomes(i)) > 3 Then
        'Se não. Pega apenas a primeira letra do nome e coloca um ponto após.
        sTemp = sTemp & Left(aNomes(i), 1) & ". "
      Else
        sTemp = sTemp & aNomes(i) & Chr(32)
      End If
    Next
    sTemp = sTemp & aNomes(nCount)
  Else
    sTemp = psNome
  End If
    
  'UpperCase = 1 , LowerCase = 2, ProperCase = 3, Unicode = 64
  If (TCase > vbProperCase) And TCase <> vbUnicode Then TCase = vbProperCase
  sTemp = StrConv(sTemp, TCase)
  
  AbreviaNome = sTemp
End Function

Public Function CGCouCPF(sText As String) As String
  Dim nLen As Integer
  
  sText = Trim(sText)
  nLen = Len(sText)
  
  Select Case nLen
    Case Is = 11
      CGCouCPF = Format(sText, "000\.000\.000-00")
    Case Is = 12, 14
      CGCouCPF = Format(sText, "00\.000\.000/0000-00")
    Case Else
      CGCouCPF = Format(sText, String(14, "0"))
  End Select
End Function

Public Function ValorData(ByVal pnData As geDataPadrão, Optional ByVal pFormato As String = "dd/mm/yyyy") As String
   Dim sData As String
   
   Select Case pnData
      Case Is = [DT Atual]
         sData = Format$(Date, pFormato)
         
      Case Is = [DT IniMês]
         sData = "01" & Mid$(Format$(Date, pFormato), 3)
         
      Case Is = [DT FimMês]
         sData = UltimoDiaDoMes(Date)
         
      Case Else
         'é o próprio conteúdo do campo
   End Select
   ValorData = sData
End Function

Public Function UltimoDiaDoMes(ByVal psData As String) As String
   Dim sMês As String, sAno As String
   Dim sMêsAno As String
   
   sMês = Format$(psData, "mm")
   sAno = Format$(psData, "yyyy")
  
   Select Case sMês
    Case "01", "03", "05", "07", "08", "10", "12"
         UltimoDiaDoMes = "31"
    Case Else
       If sMês <> "02" Then
          UltimoDiaDoMes = "30"
       Else
          If sAno Mod 4 = 0 Then
             UltimoDiaDoMes = "29"
          Else
             UltimoDiaDoMes = "28"
          End If
       End If
   End Select
   sMêsAno = Format$(psData, "mm/yyyy")
   UltimoDiaDoMes = UltimoDiaDoMes & "/" & sMêsAno
End Function

Public Sub HabilitarEdicao(ByRef Campos As Object, _
                           Optional ByVal bEditar As Boolean = True, _
                           Optional QualDataSource As Object, _
                           Optional pBackColor As OLE_COLOR)
  Dim oC As Control
  Dim nValue As Integer
  
  On Error GoTo TrataErro:
  
  nValue = Abs(CInt(bEditar))   'Retorna 0 ou 1
 
  If QualDataSource Is Nothing Then
    For Each oC In Campos
      GoSub SetarAparencia
    Next
  Else
    For Each oC In Campos
      If oC.DataSource Is QualDataSource Then
        GoSub SetarAparencia
      End If
    Next
  End If
  Set oC = Nothing
    
  Exit Sub
  
SetarAparencia:
  Select Case TypeName(oC)
    Case Is = "TextBox", "ActiveText"
      If Not oC.Locked Then
'        oC.Appearance = nValue
        oC.Enabled = nValue
      End If
    Case Is = "CheckBox"
'      oC.Appearance = nValue
      oC.Enabled = nValue
'      If pBackColor <> 0 Then oC.BackColor = pBackColor
    Case Else
      oC.Enabled = nValue
  End Select
Return

TrataErro:
  Set oC = Nothing
  Call MGShowErro("mdlGeral.HabilitarEdicao")
  On Error GoTo 0
End Sub

Public Sub CaptionsGrid(ByRef pDataGrid As Control, _
                        ByVal sTxtArrayCaptions As String, _
                        Optional ColWidth As String, _
                        Optional SumirRestante As Boolean = True)
  Dim aCap() As String
  Dim aCW() As String
  Dim i As Integer, n As Integer
  
  aCW = Split(ColWidth, ",")
  aCap = Split(sTxtArrayCaptions, ",")
  
  For i = 0 To UBound(aCap)
    pDataGrid.Columns.Item(i).Caption = aCap(i)
    On Error Resume Next
    pDataGrid.Columns.Item(i).Width = aCW(i)
    On Error GoTo 0
  Next
  
  If SumirRestante Then
    For n = i To pDataGrid.Columns.Count - 1
      pDataGrid.Columns(n).Visible = False
    Next
  End If
End Sub

Public Sub CaptionsDataFieldGrid(ByRef pDataGrid As Control, _
                                 ByVal DataFieldCaption As String, _
                                 Optional ColWidth As String)
  Dim aDatCap() As String
  Dim aCols() As String
  Dim aCW() As String
  Dim i As Integer, n As Integer, bOK As Boolean
  
  aCols = Split(DataFieldCaption, ";")
  aCW = Split(ColWidth, ",")
  
  On Error Resume Next
  For n = 0 To pDataGrid.Columns.Count - 1
    bOK = False
    For i = UBound(aCols) To 0 Step -1
      aDatCap = Split(aCols(i), ",")
      With pDataGrid.Columns
        If UCase$(aDatCap(0)) = UCase$(.Item(n).DataField) Then
          .Item(n).Caption = aDatCap(1)
          .Item(n).Width = aCW(i)
          bOK = True
          Exit For
        End If
      End With
    Next
    pDataGrid.Columns(n).Visible = bOK
  Next
End Sub


Public Function MyShell(ByVal NewHandle As Long, ByVal pPathName As String) As Long
  Dim pID As Long
  Dim mWnd As Long
  
  On Error Resume Next
  pID = Shell(pPathName, vbMinimizedNoFocus)
  If pID = 0 Then
    MsgBox "Erro ao abrir Aplicativo!"
    
  ElseIf NewHandle <> 0 Then
    mWnd = BuscarhWnd(pID)
    MyShell = mWnd
    
    SetParent mWnd, NewHandle
    
    Call AddColShellApp(mWnd)
  End If
  On Error GoTo 0
End Function

Public Sub AddColShellApp(phWnd As Long)
  If mColShellApp Is Nothing Then Set mColShellApp = New Collection

  'Com essa colecao vamos testar se existe aplicativos externos abertos
  If phWnd <> 0 Then
    mColShellApp.Add phWnd, "K" & phWnd
  End If
End Sub

Public Sub FecharAppExternos(Optional phWnd As Long)
  Dim i As Integer
  Dim sErro As String
  
  If mColShellApp Is Nothing Then Exit Sub
  
  If phWnd <> 0 Then
    Call SendMessage(phWnd, WM_CLOSE, 0&, 0&)
  Else
    On Error Resume Next
    For i = 1 To mColShellApp.Count
      Call SendMessage(mColShellApp(i), WM_CLOSE, 0&, 0&)
      If Err.Number <> 0 Then
        sErro = sErro & "Erro: " & Err.Number & " - " & Err.Description & vbCrLf
        Err.Clear
      End If
    Next
    On Error GoTo 0
    Set mColShellApp = Nothing
  End If
  
  If sErro <> "" Then _
    MsgBox sErro, vbCritical, "Fechamento de Aplicativos Externos"
End Sub

Public Function CheckShellApp(phWnd As Long) As Boolean
  Dim i As Integer
  
  If mColShellApp Is Nothing Then Exit Function
  
  If phWnd <> 0 Then
    For i = 1 To mColShellApp.Count
      If mColShellApp(i) = phWnd Then
        CheckShellApp = True
        Exit For
      End If
    Next
  End If
End Function

Private Function BuscarhWnd(ByVal InstanciaID As Long) As Long
  Dim test_hwnd As Long
  Dim test_pid As Long
  Dim test_thread_id As Long
  
  'Find the first window
  test_hwnd = FindWindow(ByVal 0&, ByVal 0&)
  
  Do While test_hwnd <> 0
    'Check if the window isn't a child
    If GetParent(test_hwnd) = 0 Then
      'Get the window's thread
      test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)
      If test_pid = InstanciaID Then
        BuscarhWnd = test_hwnd
        Exit Do
      End If
    End If
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
  Loop
End Function

Public Function FileExist(pathFile As String) As Boolean
  FileExist = (Dir(pathFile, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) <> "")
End Function

Public Function DeleteFile(ByVal pathFile As String) As Boolean
On Error GoTo Finaliza
  Kill pathFile
  
  DeleteFile = True
Finaliza:
  If Err.Number <> 0 Then
    Err.Clear
  End If
End Function

Public Function RenameFile(ByVal pathFile As String, ByVal newPath) As Boolean
On Error GoTo Finaliza
  Call DeleteFile(newPath)
  FileCopy pathFile, newPath
  Call DeleteFile(pathFile)
  
  RenameFile = True
Finaliza:
  If Err.Number <> 0 Then
    Err.Clear
  End If
End Function

Public Function ExtractResData(sID As String, sType As String, PathArqDestino As String) As Boolean
  Dim mFreeFile As Integer, bExist As Boolean
  Dim FileTmp As String, Buffer As String
  Dim vData As Variant, b As Long

  If FileExist(PathArqDestino) Then
    If MsgBox("Arquivo já existe!" & vbCrLf & vbCrLf & "Deseja substitui-ló?", _
              vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        
      bExist = True
    Else
      Exit Function
    End If
  End If
  
  On Error GoTo TrataErro
  mFreeFile = FreeFile
  FileTmp = "File.tmp~"
  
  vData = LoadResData(sID, sType)

  Open FileTmp For Binary As mFreeFile
    Put mFreeFile, , vData
  Close mFreeFile

  b = FileLen(FileTmp)
  Buffer = String(b - 12, " ")

  Open FileTmp For Binary As mFreeFile
    Seek mFreeFile, 13
    Get mFreeFile, , Buffer
  Close mFreeFile
  Kill FileTmp

  'Salvando mo arquivo correto
  If bExist Then Kill PathArqDestino
  
  Open PathArqDestino For Binary As mFreeFile
    Put mFreeFile, , Buffer
  Close mFreeFile
  
  ExtractResData = True
  Exit Function
TrataErro:
  MGErrRaise App.ProductName
End Function

Public Function ValidaData(ByVal psData As String) As Boolean
  Dim s As String
  
  On Error GoTo ValidaData_Error
  s = FormatDateTime(psData, vbShortDate)
  ValidaData = True

  On Error GoTo 0
  Exit Function

ValidaData_Error:
  ValidaData = False
End Function

Public Sub DragForm(ByVal phWnd As Long)
  On Local Error Resume Next
  'Move the borderless form...
  Call ReleaseCapture
  Call SendMessage(phWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Public Sub MGErrRaise(Optional Titulo As String)
  If Titulo = "" Then
    MsgBox Err.Description & vbCrLf & Err.Source, vbCritical
  Else
    MsgBox Err.Description & vbCrLf & Err.Source, vbCritical, Titulo
  End If
  
  On Error GoTo 0
  Err.Clear
End Sub

Public Sub MGShowErro(ByVal pProcedure As String, Optional ByVal Mensagem As String, Optional ByVal Titulo As String)
  If Mensagem = "" Then
    Mensagem = "Nome Procedure: " & pProcedure & vbCrLf & vbCrLf & _
               "Número do Erro: " & Err.Source & vbCrLf & _
               "Descrição.....: " & Err.Description & vbCrLf & vbCrLf & _
               "Consulte o Administrador de Sistema e o informe sobre o erro!"
  End If
  
  If Titulo = "" Then
    Titulo = App.FileDescription
    If Titulo = "" Then Titulo = App.EXEName
  End If
    
  Dim oF As Form
  Set oF = New FormMessage

  oF.ShowMsgBox Mensagem, Titulo, , , , imCritical
  Set oF = Nothing
  Err.Clear
End Sub

Public Sub MGShowInfo(ByVal Mensagem As String, Optional ByVal Titulo As String)
  If Titulo = "" Then
    Titulo = App.FileDescription
    If Titulo = "" Then Titulo = App.EXEName
  End If
  
  Dim oF As Form
  Set oF = New FormMessage
  oF.ShowMsgBox Mensagem, Titulo, , , , imInformation
  Set oF = Nothing
End Sub

Public Function MGShowQuest(ByVal Mensagem As String, Optional ByVal Titulo As String, _
                            Optional ByVal psBotoes As String = "&OK") As String
  If Titulo = "" Then
    Titulo = App.FileDescription
    If Titulo = "" Then Titulo = App.EXEName
  End If
  
  Dim oF As Form
  Set oF = New FormMessage
  
  MGShowQuest = oF.ShowMsgBox(Mensagem, Titulo, psBotoes, , , imQuestion)
  Set oF = Nothing
End Function

'criptografa/descriptografa
Public Function Cripty(ByVal pString As String, ByVal Pw As String) As String
  Dim s As String
  Dim i As Integer, ii As Integer
  Dim iii As Integer, iv As Integer 'dimensiona
  
  iii = 0
  For i = 1 To Len(pString$)              'para cada caracter
    iii = iii + 1                         'incrementa ponteiro
    If iii > Len(Pw$) Then iii = 1        'testa e reseta, se for o caso
    iv = Asc(Mid$(Pw$, iii, 1)) Or 128    'pega char da senha evitando acima de 128
    ii = Asc(Mid$(pString$, i))           'pega char da string a encriptar
  
DeNovo:
    ii = ii Xor iv                        'encripta...
    If ii < 31 Then                       'se char de controle
      ii = (128 + ii)                     'somar 128 e
      GoTo DeNovo                         'ecripta novamente
    ElseIf ii > 127 And ii < 159 Then     'se nesta faixa pode ser char de controle
      ii = ii - 128                       'tira 128 e
      GoTo DeNovo                         'encripta novamente
    End If
    s$ = s$ + Chr$(ii)                    'concatena string encriptada
  Next                                    'próximo caracter a encriptar
  Cripty$ = s$                             'retorna a nova string
End Function

Public Sub FlatControles(ByRef Controles As Object, Optional ByVal pEnableControl As Boolean)
  Dim oCtr As Control
  
  On Local Error GoTo TrataErro
  For Each oCtr In Controles
    oCtr.BorderStyle = 0
    oCtr.Enabled = pEnableControl
    Call FlatBorder(oCtr.hwnd, True)
    Set oCtr = Nothing
  Next
  On Error GoTo 0
  Exit Sub
TrataErro:
  Set oCtr = Nothing
  Call MGShowErro("ModuloGeral.FlatControl")
End Sub

Public Sub FlatBorder(ByVal hwnd As Long, ByVal MakeControlFlat As Boolean)
  Dim TFlat As Long
  
  TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
  If MakeControlFlat Then
    TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
  Else
    TFlat = TFlat And Not WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
  End If
  SetWindowLong hwnd, GWL_EXSTYLE, TFlat
  SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Function CalculaIdade(ByVal XData) As String
  Dim sIdade As String
  Dim nAnos As Double, nMes As Double, nDias As Double
    
  If XData = "" Or IsNull(XData) And Not IsDate(XData) Then
    sIdade = "Não Informado!"
  Else
    On Error GoTo Sair:
    '--------Calculando Anos
    nAnos = Trim(Str(Int((Val(Format(Date, "yyyymmdd")) - Val(Format(XData, "yyyymmdd"))) / 10000)))
    
    '--------Calculando Mêses
    nMes = Int((Val(Format(Date, "mmdd")) - Val(Format(XData, "mmdd"))) / 100)
    If nMes < 0 Then nMes = Trim(Str(nMes + 12))
    
    '--------Calculando Dias
    nDias = Day(XData)
    If nDias > Day(Date) Then
      nDias = Trim(Str(30 - (nDias - Day(Date))))
    Else
      nDias = Trim(Str(Day(Date) - nDias))
    End If
        
    '--------Preenchendo Anos
    sIdade = ""
    If Val(nAnos) = 1 Then
      sIdade = "1 ano"
    ElseIf Val(nAnos) > 1 Then
      sIdade = nAnos & " anos"
    End If
    
    '--------Preenchendo Meses
    If Val(nMes) > 0 Then
      If sIdade <> "" Then sIdade = sIdade & IIf(Val(nDias) > 0, ", ", " e ")
      sIdade = sIdade & Val(nMes) & IIf(Val(nMes) = 1, " mês", " meses")
    End If
    
    '--------Preenchendo Dias
    If Val(nDias) > 0 Then
      If sIdade <> "" Then sIdade = sIdade & " e "
      sIdade = sIdade & Val(nDias) & IIf(Val(nDias) = 1, " dia", " dias")
    End If
  End If
Sair:
  CalculaIdade = sIdade
End Function

Public Function IsDebug() As Boolean
On Error GoTo TrataErro
    Debug.Print 1 \ 0 'Divisao por zero, ocorre erro apenas em mode debug
    IsDebug = False
    
    Exit Function
TrataErro:
    IsDebug = True
End Function

Public Function IfNull(ByVal value As Variant, ByVal default As Variant)
  IfNull = IIf(IsNull(value), default, value)
End Function

Public Sub SetFocus(ByRef pControl As Object)
On Error GoTo TrataErro
  If pControl.Enabled And pControl.Visible Then pControl.SetFocus
  
TrataErro:
  Err.Clear
  On Error GoTo 0
End Sub

Public Function Max(ByVal val1 As Double, ByVal val2 As Double)
  Max = IIf(val1 > val2, val1, val2)
End Function

Public Function SoNumeros(ByVal pValor As String) As String
  Dim ret As String
  Dim i As Integer
  
  ret = vbNullString
  pValor = Trim$(pValor)
  
  For i = 1 To Len(pValor)
    If IsNumeric(Mid(pValor, i, 1)) Then
      ret = ret & Mid(pValor, i, 1)
    End If
  Next
  
  SoNumeros = ret
End Function
 
Public Function IsValidCPF(ByVal pCPF As String) As Boolean
  
  'Remove formatacao
  pCPF = SoNumeros(Trim$(pCPF))
  
  If Len(pCPF) <> 11 Then
    Exit Function
  Else
    
    Dim n As Integer 'numero
    Dim d As Integer 'digito
    Dim m As Integer 'multiplo
    Dim x As Integer 'validacoes
    Dim p As Integer 'posicao do numero
    
    For x = 0 To 9
      Debug.Print String(11, x & "")
      If pCPF = String(11, x & "") Then Exit Function
    Next
    
    For x = 0 To 1
      n = 0
      For p = 0 To 8 + x
        m = (10 + x - p)
        'Debug.Print Val(Mid(pCPF, p + 1, 1)), m, n
        n = n + Val(Mid(pCPF, p + 1, 1)) * m
      Next

      d = 11 - (n - (Int(n / 11) * 11))
      If d = 10 Or d = 11 Then d = 0

      'Valida o digito verificador
      If d <> Val(Mid(pCPF, 10 + x, 1)) Then
        Exit Function
      End If

    Next

  End If
  
  IsValidCPF = True
End Function

Public Function IsValidCNPJ(ByVal pCNPJ As String) As Boolean

  Dim i As Integer
  Dim valida As Boolean
  
  'Remove formatacao
  pCNPJ = SoNumeros(Trim$(pCNPJ))
  
  If Len(pCNPJ) <> 14 Then
    Exit Function
  Else
    
    Dim n As Integer 'numero
    Dim d As Integer 'digito
    Dim m As Integer 'multiplo
    Dim x As Integer 'validacoes
    Dim p As Integer 'posicao do numero
    
    For x = 0 To 9
'      Debug.Print String(14, x & "")
      If pCNPJ = String(14, x & "") Then Exit Function
    Next
    
    For x = 0 To 1
      m = 5 + x
      n = 0
      For p = 1 To 12 + x
        If m < 2 Then m = 9
        
        n = n + Mid(pCNPJ, p, 1) * m
        m = m - 1
      Next
      
      d = 11 - (n - (Int(n / 11) * 11))
      If d = 10 Or d = 11 Then d = 0
      
      'Valida o digito verificador
      If d <> Val(Mid(pCNPJ, 13 + x, 1)) Then
        Exit Function
      End If
    Next
  End If
  
  IsValidCNPJ = True
End Function
