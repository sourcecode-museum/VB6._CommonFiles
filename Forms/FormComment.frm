VERSION 5.00
Object = "{AB4C3C68-3091-48D0-BB3D-8F92CD2CB684}#1.0#0"; "AButtons.ocx"
Begin VB.Form FormComment 
   BackColor       =   &H00E9FEFE&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
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
   ScaleHeight     =   4035
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTexto 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00212121&
      Height          =   1770
      Left            =   975
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "FormComment.frx":0000
      Top             =   1200
      Width           =   7950
   End
   Begin VB.Frame fraRodape 
      BackColor       =   &H00FCFCFC&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3735
      Width           =   9210
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
         TabIndex        =   1
         Top             =   45
         Width           =   735
      End
   End
   Begin VB.PictureBox picMensagem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   870
      ScaleHeight     =   1830
      ScaleWidth      =   8040
      TabIndex        =   4
      Top             =   1140
      Width           =   8100
   End
   Begin AButtons.AButton cmdButtons 
      Height          =   420
      Index           =   0
      Left            =   7665
      TabIndex        =   6
      Top             =   3180
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   741
      BTYPE           =   4
      TX              =   "Gravar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8454016
      FCOL            =   0
   End
   Begin AButtons.AButton cmdButtons 
      Height          =   420
      Index           =   1
      Left            =   6240
      TabIndex        =   7
      Top             =   3180
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   741
      BTYPE           =   4
      TX              =   "Fechar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      FCOL            =   0
   End
   Begin VB.Label lblLabel1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informe aqui, as observações que deseja imprimir no rodapé dos relatórios de compra e venda."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   870
      TabIndex        =   5
      Top             =   480
      Width           =   8100
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
      TabIndex        =   2
      Top             =   0
      Width           =   9225
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
      Width           =   9240
   End
   Begin VB.Image imgIcone 
      Height          =   480
      Left            =   75
      Picture         =   "FormComment.frx":001A
      Top             =   360
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B8E4EB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   2670
      Index           =   0
      Left            =   180
      Top             =   120
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0049CAE0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   2310
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
      Height          =   1950
      Index           =   1
      Left            =   360
      Top             =   120
      Width           =   105
   End
End
Attribute VB_Name = "FormComment"
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
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'=================================
'NA FRENTE
'   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'ATRAZ
'   SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'=================================

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub DragForm()
  On Local Error Resume Next
  'Move the borderless form...
  Call ReleaseCapture
  Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Public Function ShowModal() As String
  Me.Show vbModal
  ShowModal = txtTexto.Text
  Unload Me
End Function

Private Sub cmdButtons_Click(Index As Integer)
  Select Case Index
  Case Is = 0
    'Gravar
    Unload Me
  Case Is = 1
    If cmdButtons(1).Caption = "Cancelar" Then
      cmdButtons(1).Caption = "Fechar"
      'txtTexto.Text = textoOriginal
    Else
      Unload Me
    End If
  End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    cmdButtons_Click 0
    KeyCode = 0
  ElseIf KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  lblCaption.Caption = App.EXEName
  lblRotulo.Caption = App.LegalCopyright
  
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set FormComment = Nothing
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call DragForm
End Sub

Private Sub txtTexto_KeyDown(KeyCode As Integer, Shift As Integer)
  cmdButtons(1).Caption = "Cancelar"
End Sub
