VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{AB4C3C68-3091-48D0-BB3D-8F92CD2CB684}#1.0#0"; "AButtons.ocx"
Begin VB.Form FormMessage 
   BackColor       =   &H00E9FEFE&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin AButtons.AButton cmdButtons 
      Height          =   420
      Index           =   0
      Left            =   4230
      TabIndex        =   2
      Top             =   2130
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   741
      BTYPE           =   4
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12117227
      FCOL            =   0
   End
   Begin VB.CheckBox chkOpcao 
      BackColor       =   &H00E9FEFE&
      Caption         =   "Não mostrar esta mensagem novamente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   60
      TabIndex        =   4
      Top             =   2385
      Visible         =   0   'False
      Width           =   4110
   End
   Begin VB.Frame fraRodape 
      BackColor       =   &H00FCFCFC&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   5595
      Begin VB.Label lblRotulo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblRotulo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   45
         Width           =   735
      End
   End
   Begin VB.ComboBox cboOpcoes 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1695
      Visible         =   0   'False
      Width           =   1950
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3450
      Top             =   345
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":0000
            Key             =   "imgFichario"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":0E52
            Key             =   "imgNovoDoc"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":172C
            Key             =   "imgWin_Lixeira"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":2006
            Key             =   "imgEscrita"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":28E0
            Key             =   "img_FolderQuestion"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":3732
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":4584
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":4E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":5738
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":6012
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":7B1C
            Key             =   "imgCritical"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":852E
            Key             =   "imgQuestion"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":8F40
            Key             =   "imgInformation"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":9952
            Key             =   "imgExclamation"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":A22C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":AB06
            Key             =   "imgStop"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormMessage.frx":B3E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTexto 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   690
      TabIndex        =   1
      Top             =   1695
      Width           =   4860
   End
   Begin VB.TextBox txtTexto 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00212121&
      Height          =   1200
      Index           =   1
      Left            =   795
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "FormMessage.frx":BCBA
      Top             =   390
      Width           =   4710
   End
   Begin VB.PictureBox picMensagem 
      BackColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   690
      ScaleHeight     =   1260
      ScaleWidth      =   4800
      TabIndex        =   8
      Top             =   330
      Width           =   4860
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   45
      Index           =   3
      Left            =   0
      Top             =   210
      Width           =   5610
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H0049CAE0&
      Caption         =   " App.EXEName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5595
   End
   Begin VB.Image imgIcone 
      Height          =   480
      Left            =   75
      Picture         =   "FormMessage.frx":BCD4
      Top             =   360
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B8E4EB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   990
      Index           =   1
      Left            =   360
      Top             =   120
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0049CAE0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   1350
      Index           =   2
      Left            =   270
      Top             =   120
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B8E4EB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   1710
      Index           =   0
      Left            =   180
      Top             =   120
      Width           =   105
   End
End
Attribute VB_Name = "FormMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HWND_TOPMOST = -1
'Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'=================================
'NA FRENTE
'   SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'ATRAZ
'   SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'=================================

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Public Enum eImagem
  imEspacoTotal = -1
  imSemImagem = 0
  imFichario = 1
  imNovoDoc = 2
  imWin_Lixeira = 3
  imEscrever = 4
  imFolderQuestion = 5
  imFolderInformation = 6
  imFavorito = 7
  imMac_Salvar = 8
  imMac_Lixeira = 9
  imCheckTable = 10
  imCritical = 11
  imQuestion = 12
  imInformation = 13
  imExclamation = 14
  '    imMac_TrocaUsuario = 15
  imCheckMenu = 15
  imStop = 16
  imPrinter = 17
  imCustom = 99
End Enum

Public Enum enuResultInput
  inputCancel
  inputOK
End Enum
Private meResultInput As enuResultInput

Private mbShowInput As Boolean  'Se caixa de mensagem for um inputbox

'---------------------------------------------------------------------------------------
' Procedure : 28/08/2006 11:24 - ShowInput
' Author    : Heliomar P. Marques
' Purpose   : Como o InputBox tradicional, abre como modal esperando uma informacao do usuario
'---------------------------------------------------------------------------------------
Public Function ShowInput(ByVal psPrompt As String, _
                          Optional ByVal psTitulo As String, _
                          Optional ByVal psValorDefault As String = "", _
                          Optional ByVal XLeft As Long = -1, _
                          Optional ByVal YTop As Long = -1, _
                          Optional ByVal pbInputPassword As Boolean, _
                          Optional ByVal pImagem As eImagem = imEscrever) As String

  Call SetarValores(psPrompt, psTitulo, XLeft, YTop, pImagem)

  CampoInput.Text = psValorDefault

  If pbInputPassword Then
    With CampoInput
      .PasswordChar = "X"
      .Font.name = "Wingdings"
      .Font.Size = 10
    End With
  End If

  Call CriarBotoes("Con&firmar,&Cancelar")

  mbShowInput = True
  Me.Show vbModal
  ShowInput = CampoInput.Text
  Unload Me
End Function

'---------------------------------------------------------------------------------------
' Procedure : 28/08/2006 11:25 - ShowInputCombo
' Author    : Heliomar P. Marques
' Purpose   : Como o InputBox tradicional, e abre como um modal travando a tela a espera de uma informacao
'---------------------------------------------------------------------------------------
Public Function ShowInputCombo(ByVal psPrompt As String, _
                               ByVal psListaArray As String, _
                               Optional ByVal psTitulo As String, _
                               Optional ByVal pnListIndex As Byte, _
                               Optional ByVal XLeft As Long = -1, _
                               Optional ByVal YTop As Long = -1, _
                               Optional ByVal pImagem As eImagem = imFichario) As String
  Dim i As Byte
  Dim aLista() As String

  Call SetarValores(psPrompt, psTitulo, XLeft, YTop, pImagem)

  CampoInput.Visible = False

  On Error GoTo TrataErro
  aLista = Split(psListaArray, ",")
  For i = 0 To UBound(aLista)
    cboOpcoes.AddItem aLista(i)
  Next
  On Error GoTo 0

  On Local Error Resume Next
  cboOpcoes.Visible = True
  cboOpcoes.Move CampoInput.Left, CampoInput.Top, CampoInput.Width
  cboOpcoes.ListIndex = pnListIndex
  On Error GoTo 0

  Call CriarBotoes("Con&firmar,&Cancelar")

  mbShowInput = True
  Me.Show vbModal
  ShowInputCombo = CampoInput.Text
  Unload Me
  Exit Function
TrataErro:
  ShowMsgBox "Erro na Lista de Operações!", "Erro: ShowInputCombo", "Fechar", , , imCritical
  Unload Me
End Function

'---------------------------------------------------------------------------------------
' Procedure : 28/08/2006 11:23 - ShowMsgBox
' Author    : Heliomar P. Marques
' Purpose   : Substituir o msgbox padrao
'---------------------------------------------------------------------------------------
Public Function ShowMsgBox(ByVal psPrompt As String, _
                           Optional ByVal psTitulo As String, _
                           Optional ByVal psCapButtonsArray As String = "&OK", _
                           Optional ByVal XLeft As Long = -1, _
                           Optional ByVal YTop As Long = -1, _
                           Optional ByVal pImagem As eImagem = imSemImagem) As String

  Call SetarValores(psPrompt, psTitulo, XLeft, YTop, pImagem)

  Call ResizeCaixaMensagem(pImagem = imEspacoTotal)
  
  CampoInput.Visible = False

  Call CriarBotoes(psCapButtonsArray)

  mbShowInput = False
  Me.Show vbModal
  ShowMsgBox = CampoInput.Text
  Unload Me
End Function

'---------------------------------------------------------------------------------------
' Procedure : 28/08/2006 11:22 - ShowInfo
' Author    : Heliomar P. Marques
' Purpose   : Mensagens informativas nao Modal
'---------------------------------------------------------------------------------------
Public Sub ShowInfo(ByVal psPrompt As String, _
                    Optional ByVal psTitulo As String, _
                    Optional ByVal XLeft As Long = -1, _
                    Optional ByVal YTop As Long = -1, _
                    Optional ByVal pImagem As eImagem = imInformation)

  Call SetarValores(psPrompt, psTitulo, XLeft, YTop, pImagem)

  Call ResizeCaixaMensagem(pImagem = imEspacoTotal)
  
  CampoInput.Visible = False

  Call CriarBotoes("&OK")

  mbShowInput = False
  Me.Show

  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Property Get IconeMsg() As Image
  Set IconeMsg = imgIcone
End Property

Public Property Get ResultInput() As enuResultInput
  ResultInput = meResultInput
End Property

Public Property Get TextoMensagem() As TextBox
  Set TextoMensagem = txtTexto(1)
End Property

Private Property Get CampoInput() As TextBox
  Set CampoInput = txtTexto(0)
End Property

Public Property Let AtualizarMensagem(ByVal pMsg As String)
  TextoMensagem.Text = pMsg
End Property

Private Sub cboOpcoes_Click()
  CampoInput.Text = cboOpcoes.Text
End Sub

Private Sub chkOpcao_Click()

End Sub

Private Sub cmdButtons_Click(Index As Integer)
  If mbShowInput Then
    meResultInput = IIf(Index = 0, inputOK, inputCancel)
    If meResultInput = inputCancel Then
      CampoInput.Text = ""
    End If
  Else
    CampoInput.Text = cmdButtons(Index).Caption
  End If

  Me.Hide
End Sub

Private Sub CriarBotoes(ByVal psCapButtonsArray As String)
  Dim aCap() As String
  Dim i As Integer
  Dim nLeftPadrao As Integer
  Dim nWidth As Long
  Dim nCount As Integer

  nLeftPadrao = cmdButtons(0).Left
  Select Case True
  Case InStr(1, psCapButtonsArray, ",") > 0
    psCapButtonsArray = Replace(psCapButtonsArray, ",", "|")

  Case InStr(1, psCapButtonsArray, ";") > 0
    psCapButtonsArray = Replace(psCapButtonsArray, ";", "|")

  End Select

  aCap = Split(psCapButtonsArray, "|")

  nCount = UBound(aCap)
  'Numero de Botoes que pode ser incluido,
  'os demais sao ignorados
  If nCount > 2 Then nCount = 2
  
  For i = nCount To 0 Step -1
    If i > 0 Then Load cmdButtons(i)
    With cmdButtons(i)
      .Caption = aCap(i)
      If TextWidth(aCap(i)) + 200 > .Width Then
        .Width = TextWidth(aCap(i)) + 200
      End If
      
'      .Left = nLeftPadrao - (.Width * (nCount - i))
      
      If i = nCount Then
        .Left = Me.ScaleWidth - (.Width + 60)
      Else
        .Left = cmdButtons(i + 1).Left - .Width - 15
      End If

      .Visible = True
    End With
  Next
End Sub

Private Sub DragForm()
  On Local Error Resume Next
  'Move the borderless form...
  Call ReleaseCapture
  Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Form_Click()
'  MsgBox TextWidth(TextoMensagem.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    cmdButtons_Click 0
    KeyCode = 0
  ElseIf KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Load()
  Dim sCap As String

  sCap = App.CompanyName
  lblRotulo.Caption = gtApp.URL

  TextoMensagem.Text = ""
  imgIcone.BorderStyle = 0
  CampoInput.Visible = True
  cboOpcoes.Visible = False

  'lblRotulo.BackColor = RGB(210, 198, 108)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set FormMessage = Nothing
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  Call DragForm
End Sub

Private Sub SetarValores(ByVal pPrompt As String, ByVal pTitulo As String, _
                         ByVal XLeft As Long, ByVal YTop As Long, _
                         ByVal pImagem As eImagem)

  If pTitulo = "" Then pTitulo = App.ProductName
  lblCaption.Caption = pTitulo

  Me.Top = IIf(YTop = -1, (Screen.Height - Me.Height) / 2, YTop)
  Me.Left = IIf(XLeft = -1, (Screen.Width - Me.Width) / 2, XLeft)

  TextoMensagem.Text = pPrompt

  On Local Error Resume Next
  If pImagem = imSemImagem Or pImagem = imEspacoTotal Then
    imgIcone.Visible = False
  Else
    imgIcone.Visible = True
    imgIcone.Picture = ImageList1.ListImages(pImagem).Picture
  End If
  On Error GoTo 0
End Sub

Private Sub ResizeCaixaMensagem(ByVal pAmpliar As Boolean)
  If pAmpliar Then
    picMensagem.Left = 100
    picMensagem.Width = Me.ScaleWidth - 200
  End If
  picMensagem.Height = (CampoInput.Top + CampoInput.Height) - picMensagem.Top
  
  TextoMensagem.Left = picMensagem.Left + 100
  TextoMensagem.Width = picMensagem.Width - 150
  TextoMensagem.Height = picMensagem.Height - 120
End Sub
