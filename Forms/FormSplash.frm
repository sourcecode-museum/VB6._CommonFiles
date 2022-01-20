VERSION 5.00
Begin VB.Form FormSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versão: 1.4.8888"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   3
      Top             =   2085
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "App.ProductName"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1230
      Index           =   2
      Left            =   345
      TabIndex        =   2
      Top             =   600
      Width           =   3150
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgIcon 
      Height          =   435
      Left            =   3855
      Top             =   600
      Width           =   435
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C48902&
      BorderWidth     =   15
      Height          =   795
      Index           =   0
      Left            =   3675
      Top             =   420
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFCA88&
      BorderWidth     =   15
      FillColor       =   &H008D550A&
      Height          =   795
      Index           =   1
      Left            =   3210
      Top             =   810
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0025B1DA&
      BorderWidth     =   15
      Height          =   795
      Index           =   2
      Left            =   2790
      Top             =   1350
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Top             =   30
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FormSplash.frx":0000
      ForeColor       =   &H80000008&
      Height          =   660
      Index           =   5
      Left            =   105
      TabIndex        =   5
      Top             =   2340
      Width           =   4560
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By: Heliomar P. Marques"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0025B1DA&
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1290
      Index           =   6
      Left            =   360
      Top             =   255
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1755
      Index           =   5
      Left            =   240
      Top             =   255
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2055
      Index           =   3
      Left            =   120
      Top             =   255
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contato: codeuapp@gmail.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2325
      TabIndex        =   4
      Top             =   3015
      Width           =   2250
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2670
      Index           =   4
      Left            =   105
      Top             =   240
      Width           =   4560
   End
End
Attribute VB_Name = "FormSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
'Sendo usada para fechar as janelas dos Aplicativos externos
'E para Drag em Forms
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Const WM_CLOSE = &H10
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
  Label1(2).Caption = App.ProductName
  Label1(3).Caption = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call DragForm
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Label1(0).ForeColor = vbWhite
  Label1(4).ForeColor = vbWhite
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set FormSplash = Nothing
End Sub

Private Sub Label1_Click(Index As Integer)
    On Local Error Resume Next
    
    Dim email As String
    
    Select Case Index
    Case 0
        Unload Me
    Case 4
        email = Replace(Label1(4).Caption, "Contato: ", "mailto:")
        Call ShellExecute(0&, vbNullString, email, vbNullString, "C:\", SW_SHOWNORMAL)
    End Select
End Sub

Public Sub DragForm()
  On Local Error Resume Next
  'Move the borderless form...
  Call ReleaseCapture
  Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  If Index = 1 Then Call DragForm
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  If Index = 0 Then Label1(Index).ForeColor = &H25B1DA
  If Index = 4 Then Label1(Index).ForeColor = vbBlue
End Sub

Public Sub CloseIn(ByVal dwMilliseconds As Long)
  Me.Visible = True
  Me.ZOrder 0
  Sleep dwMilliseconds
  Unload Me
End Sub

