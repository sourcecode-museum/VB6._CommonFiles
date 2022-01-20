VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{7493D2DD-8190-4122-AEA8-67726C4A96F5}#4.0#0"; "ideFrame.ocx"
Object = "{AB4C3C68-3091-48D0-BB3D-8F92CD2CB684}#1.0#0"; "AButtons.ocx"
Object = "{C6FEE5AC-DF5F-47A6-BE77-6DCE10AA8AB9}#4.1#0"; "ideDSControl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FormCadListaSimples 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6225
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Century Gothic"
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
   ScaleHeight     =   4320
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Insignia_Frame.ideFrame Panel 
      Align           =   1  'Align Top
      Height          =   1260
      Index           =   0
      Left            =   0
      Top             =   960
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   2223
      BackColor       =   16579836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin rdActiveText.ActiveText txtCampo 
         DataField       =   "NOME"
         Height          =   300
         Index           =   1
         Left            =   1185
         TabIndex        =   4
         Tag             =   "Obrigatorio"
         Top             =   690
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   529
         Appearance      =   0
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   100
         TextCase        =   1
         RawText         =   0
         FontName        =   "Century Gothic"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtCampo 
         DataField       =   "ID"
         Height          =   300
         Index           =   0
         Left            =   1185
         TabIndex        =   2
         Tag             =   "Obrigatorio"
         Top             =   240
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   529
         Alignment       =   1
         Appearance      =   0
         BackColor       =   14737632
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         TextMask        =   3
         RawText         =   3
         FloatFormat     =   2
         FontName        =   "Century Gothic"
         FontSize        =   8,25
         Locked          =   -1  'True
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         ForeColor       =   &H008D550A&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código ID:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008D550A&
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   285
         Width           =   825
      End
   End
   Begin Insignia_Frame.ideFrame PanelTitle 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   979
      BorderExt       =   6
      BorderPaint     =   8
      BorderWidth     =   40
      BackColor       =   10711090
      BackColorB      =   16763528
      GradientStyle   =   4
      Caption         =   "Cadastro Lista Simples"
      ForeColor       =   16777215
      CaptionAlign    =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Timer tmrResize 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4950
         Top             =   75
      End
      Begin AButtons.AButton cmdFechar 
         Height          =   270
         Left            =   5865
         TabIndex        =   5
         Top             =   150
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   476
         BTYPE           =   5
         TX              =   "r"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14133058
         FCOL            =   0
      End
      Begin VB.Image imgLogoForm 
         Height          =   480
         Left            =   135
         Picture         =   "FormCadListaSimples.frx":0000
         Top             =   45
         Width           =   480
      End
   End
   Begin Insignia_DSControl.ideDSControl dscMaster 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   555
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   714
      Modelo          =   1
      ButtonsExtras   =   6
      ButtonColor     =   15987699
      BackColor       =   15987699
   End
   Begin Insignia_Frame.ideFrame Panel 
      Align           =   2  'Align Bottom
      Height          =   300
      Index           =   2
      Left            =   0
      Top             =   4020
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   529
      BorderExt       =   6
      BorderWidth     =   5
      BackColor       =   15987699
      Caption         =   "contato: codeuapp@gmail.com"
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
   Begin MSDataGridLib.DataGrid dtgMaster 
      Align           =   1  'Align Top
      Height          =   1770
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2220
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   3122
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   3
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "Cód. ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NOME"
         Caption         =   "Nome"
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4004,788
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormCadListaSimples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=-=-=-=-=-=-= Variaveis de controle padrao
Private msLastValue     As String     'Guarda o valor do campo qdo recebe o Focu
Private mnCorLostFocus  As Long       'Guarda o Background do campo qdo recebe o Focu

'=-=-=-=-=-=-= Definição Especifica da Janela
Private Const TBL_MASTER  As String = "NOME_TABELA"

' --- Index dos Campos
Private Const cID       As Byte = 0
Private Const cNome     As Byte = 1

Public Sub ShowModal()
  Dim sSQL As String
  
  Me.MousePointer = vbHourglass

  sSQL = "SELECT * FROM " & TBL_MASTER

  If dscMaster.Conectar(sSQL, gOConn) = cnErroProcesso Then
    Me.MousePointer = vbDefault
    Unload Me
    
  Else
    Load Me
    Call ConfigurarDados
    
    Me.MousePointer = vbDefault
    Me.Show vbModal
  End If
End Sub

Private Sub ConfigurarDados()
  Dim oT As ActiveText
  
  Static configOK As Boolean
  Dim sPesq As String, sMask As String

  For Each oT In txtCampo

    Set oT.DataSource = dscMaster.DataSource.RS
    
    If Not configOK Then
      Select Case oT.Index
      Case Is = cID, cNome
        sMask = IIf(oT.TextMask = [Integer Mask], "############", oT.Mask)
        sPesq = sPesq & _
                Replace(lblLabel(oT.Index).Caption, ":", "") & "," & _
                oT.DataField & "," & sMask & "|"
                
      End Select
    End If
    
    Set oT = Nothing
    DoEvents
  Next

  Set dtgMaster.DataSource = dscMaster.DataSource.RS
  
  If Not configOK Then
    configOK = True
    sPesq = Mid$(sPesq, 1, Len(sPesq) - 1)
    dscMaster.MontarPesquisa sPesq
  End If

End Sub

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub Form_Initialize()
  Panel(2).Caption = App.LegalCopyright
  Call mdlGeral.FlatControles(txtCampo)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call MFormsKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If dscMaster.DataSource.OperacaoPendente Then
    Cancel = True
  Else
    Set FormCadListaSimples = Nothing
  End If
End Sub

Private Sub Form_Resize()
  tmrResize.Enabled = True
End Sub

Private Sub PanelTitle_DblClick()
  Me.WindowState = IIf(Me.WindowState = vbNormal, vbMaximized, vbNormal)
End Sub

Private Sub PanelTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call mdlGeral.DragForm(Me.hwnd)
End Sub

Private Sub tmrResize_Timer()
  If Me.WindowState <> vbMinimized Then
    With cmdFechar
      .Top = .Container.Height - CInt(.Height * 1.5)
      .Left = Me.ScaleWidth - .Width - 150
    End With
    dtgMaster.Height = Panel(2).Top - dtgMaster.Top
  End If
  
  tmrResize.Enabled = False
End Sub

Private Sub txtCampo_GotFocus(Index As Integer)
  msLastValue = txtCampo(Index).Text
  mnCorLostFocus = txtCampo(Index).BackColor
  txtCampo(Index).BackColor = modConstantes.gcCorFocus
End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case Is = cNome
      Call mdlADO.AutoComplete(txtCampo(Index), txtCampo(Index).DataField, TBL_MASTER, KeyAscii)
  End Select
End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
  txtCampo(Index).BackColor = mnCorLostFocus
  If msLastValue = txtCampo(Index).Text Then Exit Sub
End Sub

Private Sub dtgMaster_DblClick()
  dscMaster.Edit
End Sub

Private Sub dscMaster_AntesUpdate(Cancel As Boolean, eOperacao As Insignia_DSControl.eDSOperacao)
  Cancel = Not modForms.MFormsValidateRequiredFields(Me.txtCampo)
  
  If Not Cancel Then
    'Verificar informações duplicadas
    Dim nID     As Long
    Dim sWhere  As String
    Dim sMsg    As String
    
    sWhere = "NOME = '" & Replace(txtCampo(1).Text, "'", "''") & "'"
    If eOperacao = opAlteracao Then
      sWhere = sWhere & " AND ID <> " & txtCampo(0).Text
    End If
    If modQuerys.RegistroExiste(TBL_MASTER, sWhere, nID) Then
      sMsg = " já contém este NOME"
    End If
    
    If sMsg <> "" Then
      sMsg = "O registro [ " & nID & " ]" & sMsg & "!"
      sMsg = sMsg & vbCrLf & vbCrLf & "Deseja continuar a Gravação?"
      
      Dim msgRet As String
      msgRet = MGShowQuest(sMsg, "Confirmação de Gravação", "&Sim|&Não")
      Cancel = (msgRet = "&Não")
    End If
  End If
End Sub

Private Sub dscMaster_Operacao(ByVal eOperacao As Insignia_DSControl.eDSOperacao, ByVal eOperacaoAnterior As Insignia_DSControl.eDSOperacao)
  Dim iFocu As Byte
  Dim bEdit As Boolean
  
  bEdit = (eOperacao <> opVisualizacao)
  dtgMaster.Enabled = Not bEdit
  mdlGeral.HabilitarEdicao txtCampo, bEdit
  
  If bEdit Then iFocu = cNome
  If txtCampo(iFocu).Enabled And txtCampo(iFocu).Visible Then txtCampo(iFocu).SetFocus
End Sub
