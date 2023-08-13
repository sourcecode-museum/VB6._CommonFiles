VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{7493D2DD-8190-4122-AEA8-67726C4A96F5}#4.0#0"; "ideFrame.ocx"
Object = "{AB4C3C68-3091-48D0-BB3D-8F92CD2CB684}#1.0#0"; "AButtons.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FormPesquisa 
   BackColor       =   &H80000009&
   Caption         =   "Pesquisa R�pida..."
   ClientHeight    =   4950
   ClientLeft      =   1965
   ClientTop       =   2685
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPesquisa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FormPesquisa.frx":058A
   ScaleHeight     =   4950
   ScaleWidth      =   7965
   StartUpPosition =   1  'CenterOwner
   Begin Insignia_Frame.ideFrame Panel 
      Align           =   2  'Align Bottom
      Height          =   405
      Index           =   3
      Left            =   0
      Top             =   3105
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   714
      BorderExt       =   6
      BorderPaint     =   10
      BackColor       =   16777215
      BackColorB      =   15987699
      GradientStyle   =   1
      Caption         =   "Clique duas vezes ou tecle [ENTER] para selecionar o registro"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   1  'Align Top
      Height          =   1920
      Left            =   0
      TabIndex        =   7
      Top             =   570
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   3387
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         ScrollBars      =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerResize 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4050
      Top             =   2505
   End
   Begin Insignia_Frame.ideFrame Panel 
      Align           =   2  'Align Bottom
      Height          =   300
      Index           =   2
      Left            =   0
      Top             =   4650
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   529
      BorderExt       =   6
      BorderWidth     =   5
      BackColor       =   15987699
      Caption         =   "App.LegalTrademarks"
      ForeColor       =   8421504
      CaptionAlign    =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Insignia_Frame.ideFrame Panel 
      Align           =   1  'Align Top
      Height          =   570
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   1005
      BorderExt       =   6
      BorderPaint     =   8
      BorderWidth     =   40
      BackColor       =   10711090
      BackColorB      =   16763528
      GradientStyle   =   4
      Caption         =   "Titulo"
      ForeColor       =   16777215
      CaptionAlign    =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Image Image1 
         Height          =   480
         Left            =   105
         Picture         =   "FormPesquisa.frx":08CC
         Top             =   60
         Width           =   480
      End
   End
   Begin Insignia_Frame.ideFrame Panel 
      Align           =   2  'Align Bottom
      Height          =   1140
      Index           =   1
      Left            =   0
      Top             =   3510
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   2011
      BorderExt       =   6
      BorderPaint     =   10
      BorderWidth     =   20
      BackColor       =   15987699
      BackColorB      =   16579836
      GradientStyle   =   4
      Caption         =   "Registros encontrados: "
      ForeColor       =   0
      CaptionAlign    =   4
      CaptionAlignPos =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin AButtons.AButton btnPesquisar 
         Height          =   525
         Left            =   6660
         TabIndex        =   6
         Top             =   345
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   926
         BTYPE           =   5
         TX              =   "&Pesquisar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         FCOL            =   0
      End
      Begin rdActiveText.ActiveText txtCampo 
         Height          =   315
         Index           =   0
         Left            =   2670
         TabIndex        =   3
         Top             =   555
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextCase        =   1
         RawText         =   0
         FontName        =   "Verdana"
         FontSize        =   8,25
      End
      Begin VB.ComboBox cmbCampos 
         Height          =   315
         Left            =   330
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   555
         Width           =   2295
      End
      Begin rdActiveText.ActiveText txtCampo 
         Height          =   315
         Index           =   1
         Left            =   4785
         TabIndex        =   5
         Top             =   555
         Width           =   1600
         _ExtentX        =   2831
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         TextMask        =   9
         TextCase        =   1
         RawText         =   9
         Mask            =   "##/##/####"
         FontName        =   "Verdana"
         FontSize        =   8,25
      End
      Begin VB.Image imgOrderBy 
         Height          =   240
         Left            =   60
         Picture         =   "FormPesquisa.frx":1196
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&at�"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4395
         TabIndex        =   4
         Top             =   585
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor de Pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2685
         TabIndex        =   1
         Top             =   300
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Pesquizar por:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   0
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.Menu menuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuPop 
         Caption         =   "Ordenar por:"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   1
      End
   End
End
Attribute VB_Name = "FormPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'########################################################################################
'# Exemplo de psCapFieldMask =
'#  "Caption       ,Field    ,Maskara    |Caption     ,Field    ,Maskara  "
'#  "Data do Pedido,DATA_PEDI,�dd/mm/yyyy|N� do Pedido,NUME_PEDI,999999;0;"
'#  {OBS: as mascaras:
'#   '�dd/mm/yyyy' = inicial com � significa que a pesquisa utilizara os campos de data
'#   '�P'          = significa que o campo n�o deve entrar no combo de pesquisa
'#
'#  Alt 164 = �  /  Alt 209 = �
'########################################################################################

Private Const cTXTWidth As Single = 3705

Private WithEvents mRSQuery  As ADODB.Recordset
Attribute mRSQuery.VB_VarHelpID = -1

Private msSQLPrimary    As String

Private msaIndexColReturn As String
Private msaValoresRetorno As String
Private msAliasFieldMask    As String

Private maFields()      As String
Private maMasks()       As String

Private Sub MontarTela(ByVal psAliasFieldMask As String, ByVal pAddCombo As Boolean)
  Dim sCapt As String, sField As String, sMask As String


  Dim aI() As String, aL() As String

  aI = Split(psAliasFieldMask, "|")

  Dim i As Byte
  For i = 0 To UBound(aI)
    aL = Split(aI(i), ",")

    sCapt = Trim$(aL(0))
    sField = Trim$(aL(1))
    sMask = LTrim$(aL(2))
 
    'Defini��o de Campos de Pesquisa e Ordem
    'Preenchimento do array de mascara
    If Mid$(sMask, 1, 2) <> "�P" And pAddCombo Then
      'Se <> ent�o adicionando na Combo os Campos
      With cmbCampos
        .AddItem sCapt
        ReDim Preserve maFields(.ListCount)
        ReDim Preserve maMasks(.ListCount)
        maFields(.ListCount) = sField
        maMasks(.ListCount) = sMask
      End With
      cmbCampos.ListIndex = 0
    End If
        
        'Formatando as colunas do grid
        Dim sngW1 As Single, sngW2 As Single
        
        If InStr(1, sMask, "$") > 0 Then
    '        codfmt.Format = "@.@@@-@"
    '        datfmt.Format = "mm-dd-yyyy"
    '        valfmt.Format = "R$##,##0.00"

            Dim dataFormat As StdFormat.StdDataFormat
            Set dataFormat = New StdFormat.StdDataFormat
            
            dataFormat.Format = "###,###,##0.00"
            Set DataGrid1.Columns(i).dataFormat = dataFormat
            DataGrid1.Columns(i).Alignment = dbgRight
            Set dataFormat = Nothing
            
            sngW1 = TextWidth(DataGrid1.Columns(i).Caption)
            sngW2 = TextWidth("###,###,##0.00")
            If sngW2 > sngW1 Then sngW1 = sngW2
            DataGrid1.Columns(i).Width = sngW1
        
        ElseIf Mid(UCase(sMask), 1, 1) = "�" Then
            DataGrid1.Columns(i).Alignment = dbgCenter
            
            sngW1 = TextWidth(DataGrid1.Columns(i).Caption)
            sngW2 = TextWidth("DD/MM/YYYY")
            If sngW2 > sngW1 Then sngW1 = sngW2
            DataGrid1.Columns(i).Width = sngW1
        End If
  Next
    
End Sub

Public Function ShowForm(ByVal pSQL As String, ByRef pConnection, _
                         ByVal psAliasFieldMask As String, ByVal psaIndexFieldsRetorno As String) As String

On Error GoTo TrataErro:
  Set mRSQuery = New ADODB.Recordset
    
  With mRSQuery
    .CursorLocation = adUseServer
    .Open pSQL, pConnection, adOpenForwardOnly, adLockReadOnly, adCmdText

    Set DataGrid1.DataSource = mRSQuery
        
    Dim i As Integer
    For i = 0 To DataGrid1.Columns.Count - 1
      Load mnuPop(i + 2)
      mnuPop(i + 2).Caption = Trim(DataGrid1.Columns(i).Caption)
      mnuPop(i + 2).Visible = True
      mnuPop(i + 2).Enabled = True
    Next
    
    Panel(1).Caption = "Registros Encontrados: " & .RecordCount
  End With

  msSQLPrimary = pSQL
  
  msAliasFieldMask = psAliasFieldMask
  Call MontarTela(psAliasFieldMask, True)
  
  msaIndexColReturn = psaIndexFieldsRetorno
    
  Call LerArquivoINI
    
  Me.Show vbModal
  ShowForm = msaValoresRetorno
  
  Unload Me
  On Error GoTo 0
  Exit Function
  
TrataErro:
  FormMessage.ShowMsgBox Err.Source & ":(" & Err.Description & ")", "FormPesquisa.ShowForm", , , , imCritical
  Unload Me
End Function

Private Sub cmbCampos_Click()
  Dim sMask As String
  
  'Se a Mascara for �dd/mm/yyyy entao
  'desabilita o Mask e habilita os Campos data
  If cmbCampos.ListIndex <> -1 Then
    sMask = maMasks(cmbCampos.ListIndex + 1)
    
        
    If Mid(sMask, 1, 1) = "�" Then
      txtCampo(0).TextMask = [Custom Mask]
      txtCampo(0).Width = 1600
      txtCampo(1).Enabled = True
    
      sMask = Mid(sMask, 2, Len(sMask)) 'Retirando o �
                sMask = Replace(LCase(sMask), "d", "#")
                sMask = Replace(LCase(sMask), "m", "#")
                sMask = Replace(LCase(sMask), "y", "#")
                
                txtCampo(1).Mask = sMask
    Else
      txtCampo(0).Width = cTXTWidth
      txtCampo(1).Enabled = False
          txtCampo(0).TextMask = [No Mask]
    End If
    
    If sMask = "$" Then
            txtCampo(0).TextMask = [Float Mask]
        Else
            txtCampo(0).Mask = sMask
        End If
    If Trim(sMask) = "" Then txtCampo(0).MaxLength = 0
    If txtCampo(0).Visible And txtCampo(0).Enabled Then
        txtCampo(0).SetFocus
    End If
  End If

'  mRS.Sort = maDataField(cmbCampos.ListIndex)
'  txtCampo.Text = ""
End Sub

Private Sub DataGrid1_DblClick()
  Dim a() As String
  Dim i As Byte

    If msaIndexColReturn <> "" And DataGrid1.Row > -1 Then
    a = Split(msaIndexColReturn, "|")
    For i = 0 To UBound(a)
      msaValoresRetorno = msaValoresRetorno & DataGrid1.Columns(a(i)).value & "|"
    Next
    msaValoresRetorno = Mid$(msaValoresRetorno, 1, Len(msaValoresRetorno) - 1)
    Unload Me
  End If
End Sub

Private Sub Pesquisar(ByVal psCampo As String, _
                      ByVal psValor1 As String, ByVal psValor2 As String, _
                      ByVal psMask As String, _
                      Optional ByVal pbData As Boolean = False)

  Dim sSQL As String

  If psCampo = "" Then
    MGShowInfo "� necess�rio informar o campo de pesquisa."
    Exit Sub
  End If

  'Se tiver pesquisa que envolva datas, n�o aceita que a duas fiquem vazias
  If pbData Then
    If psValor1 = "" And psValor2 = "" Then
      MGShowInfo "Informe pelo menos uma das datas."
      Exit Sub
    End If
  End If
  
  'Prepara a pesquisa c/ base nos dados fornecidos
  MousePointer = vbHourglass
  sSQL = msSQLPrimary

  If psValor1 <> "" Or psValor2 <> "" Then
    If InStr(UCase(sSQL), "WHERE") = 0 Then
      sSQL = sSQL + " WHERE "
    Else
      sSQL = sSQL + " AND "
    End If

    If Not pbData Then
      If psValor1 <> "" Then
        If Mid$(psMask, 1, 1) = "#" Then 'S� aceita numeros
          sSQL = sSQL & psCampo & " = " & psValor1
        Else
          sSQL = sSQL & psCampo & " LIKE '" & psValor1 & "%'"
        End If
      End If
      
    Else
      If ValidaData(psValor1) And Not ValidaData(psValor2) Then
        sSQL = sSQL & psCampo & " >= " & psValor1
        
      ElseIf Not ValidaData(psValor1) And ValidaData(psValor2) Then
        sSQL = sSQL & psCampo & " <= " & psValor2
      
      ElseIf ValidaData(psValor1) And ValidaData(psValor2) Then
        sSQL = sSQL & psCampo & " BETWEEN " & psValor1 & " AND " & psValor2
      End If
    End If
  End If

        sSQL = Trim(sSQL)
    'FINAL DA EXPRESSAO
    If Mid(sSQL, Len(sSQL) - Len("WHERE")) = " WHERE" Then
        sSQL = Mid(sSQL, 1, Len(sSQL) - Len("WHERE"))
    End If
    If Mid(sSQL, Len(sSQL) - Len("AND")) = " AND" Then
        sSQL = Mid(sSQL, 1, Len(sSQL) - Len("AND"))
    End If
    
  sSQL = sSQL & " ORDER BY " & psCampo

  With mRSQuery
    On Error Resume Next
    .Close
    On Error GoTo 0
    
    On Error GoTo TrataErro
    .Open sSQL
    Set DataGrid1.DataSource = mRSQuery
    
    ' Call MontaTela(msAliasFieldMask, False)
    
    Panel(1).Caption = "Registros Encontrados: " & .RecordCount
    
    If .RecordCount > 0 Then DataGrid1.SetFocus
  End With
  MousePointer = vbDefault
  Exit Sub
  
TrataErro:
  MousePointer = vbDefault
  MGShowErro "FormPesquisa.Pesquisar"
  Err.Clear
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then DataGrid1_DblClick
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = vbRightButton _
    And mRSQuery.RecordCount > 1 Then
    PopupMenu menuPop
  End If
  On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case Is = vbKeyEscape
      Unload Me
    Case Is = vbKeyReturn
      If TypeName(ActiveControl) = "DataGrid1" Then DataGrid1_DblClick
  End Select
End Sub

Private Sub Form_Load()
  menuPop.Visible = False
'  Me.Caption = MDIMain.Caption
  txtCampo(0).Width = cTXTWidth
  Panel(2).Caption = gtApp.URL
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  If mRSQuery.Status = adStateOpen Then mRSQuery.Close
  Set mRSQuery = Nothing
  On Error GoTo 0
  
  Call GravarArquivoINI
  Set FormPesquisa = Nothing
End Sub

Private Sub Form_Resize()
  TimerResize.Enabled = True  'isto � feito para evitar atrazos no resize do form
End Sub

Private Sub imgOrderBy_Click()
    PopupMenu menuPop
End Sub

Private Sub mnuPop_Click(Index As Integer)
  On Local Error Resume Next
  If Index >= 2 Then
    mRSQuery.Sort = "[" & mRSQuery.Fields(Index - 2).Name & "]"
  End If
End Sub

Private Sub TimerResize_Timer()
  On Error Resume Next
  DataGrid1.Move 0, Panel(0).Height, ScaleWidth
  DataGrid1.Height = (Panel(3).Top - DataGrid1.Top) + 10
  On Error GoTo 0
  
  TimerResize.Enabled = False
End Sub

Private Sub txtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If Index = 0 Then
      If Not txtCampo(1).Enabled Then
        btnPesquisar_Click
      Else
        On Error Resume Next
        txtCampo(1).SetFocus
        On Error GoTo 0
      End If
    Else
      btnPesquisar_Click
    End If
  End If
End Sub

Private Sub btnPesquisar_Click()
  Dim sValor1 As String, sValor2 As String, sMask As String
  Dim bPesqData As Boolean

  If Not txtCampo(1).Enabled Then  'N�o esta utilizando os 2 Campos
    sValor1 = txtCampo(0).Text
    sValor2 = ""
    sMask = txtCampo(0).Mask
    bPesqData = False
    
  Else 'visualiza os dois campos
    sMask = maMasks(cmbCampos.ListIndex + 1)
    sMask = Mid$(sMask, 2, Len(sMask) - 1) 'Retira o Caracter � da mascara de data
    
    Select Case sMask
      Case Is = "##/##/####"
        sMask = "dd/mm/yyyy"
      Case Is = "##/##"
        sMask = "dd/mm"
      Case Is = "##/####"
        sMask = "mm/yyyy"
    End Select

    sValor1 = Format(txtCampo(0).Text, sMask)
    sValor2 = Format(txtCampo(1).Text, sMask)
    sMask = ""
    bPesqData = True
  End If

  Call Pesquisar(maFields(cmbCampos.ListIndex + 1), _
                 sValor1, sValor2, sMask, bPesqData)
End Sub

Private Sub LerArquivoINI()
'  On Local Error Resume Next
'  modArquivoINI.PathFile = App.Path & "\IniFiles\USR" & Format(usuarioAtivo, "000000") & ".ini"
'
'  cmbCampos.ListIndex = modArquivoINI.Ler(Panel(0).Caption, "CAMPO", 0)
'  txtCampo(0).Text = modArquivoINI.Ler(Panel(0).Caption, "VALOR1", "")
'  txtCampo(1).Text = modArquivoINI.Ler(Panel(0).Caption, "VALOR2", "")
'  On Error GoTo 0
End Sub

Private Sub GravarArquivoINI()
'  On Local Error Resume Next
'  modArquivoINI.PathFile = App.Path & "\IniFiles\USR" & Format(usuarioAtivo, "000000") & ".ini"
'
'  modArquivoINI.Gravar Panel(0).Caption, "CAMPO", cmbCampos.ListIndex
'  modArquivoINI.Gravar Panel(0).Caption, "VALOR1", txtCampo(0).Text
'  modArquivoINI.Gravar Panel(0).Caption, "VALOR2", txtCampo(1).Text
'  On Error GoTo 0
End Sub
