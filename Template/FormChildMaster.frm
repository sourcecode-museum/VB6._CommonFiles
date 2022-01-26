VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ActiveText.ocx"
Object = "{7493D2DD-8190-4122-AEA8-67726C4A96F5}#4.0#0"; "ideFrame.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{AB4C3C68-3091-48D0-BB3D-8F92CD2CB684}#1.0#0"; "AButtons.ocx"
Object = "{C6FEE5AC-DF5F-47A6-BE77-6DCE10AA8AB9}#4.1#0"; "ideDSControl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FormChildMaster 
   BackColor       =   &H00FCFCFC&
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10560
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   10560
   WindowState     =   2  'Maximized
   Begin Insignia_Frame.ideFrame Panel 
      Align           =   1  'Align Top
      Height          =   1995
      Index           =   0
      Left            =   0
      Top             =   1365
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   3519
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
      Begin VB.CommandButton cmdPesquisa 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2190
         Picture         =   "FormChildMaster.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1140
         Width           =   300
      End
      Begin rdActiveText.ActiveText txtCampo 
         DataField       =   "NOME"
         Height          =   300
         Index           =   1
         Left            =   1185
         TabIndex        =   3
         Tag             =   "Obrigatorio"
         Top             =   690
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   529
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
         TabIndex        =   1
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
      Begin rdActiveText.ActiveText txtCampo 
         DataField       =   "ID_CIDADE"
         Height          =   300
         Index           =   2
         Left            =   1185
         TabIndex        =   9
         Top             =   1140
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   529
         Alignment       =   1
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
         FontName        =   "Century Gothic"
         FontSize        =   8,25
      End
      Begin rdActiveText.ActiveText txtCampoFK 
         DataField       =   "NOME"
         Height          =   300
         Index           =   2
         Left            =   2520
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1140
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   529
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
         MaxLength       =   30
         TextCase        =   1
         RawText         =   0
         FloatFormat     =   2
         FontName        =   "Century Gothic"
         FontSize        =   8,25
      End
      Begin VB.Label lblRotulos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   11
         Top             =   1170
         Width           =   675
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         ForeColor       =   &H008D550A&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   2
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
         TabIndex        =   0
         Top             =   285
         Width           =   825
      End
   End
   Begin Insignia_Frame.ideFrame PanelTitle 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   979
      BorderExt       =   6
      BorderPaint     =   8
      BorderWidth     =   40
      BackColor       =   10711090
      BackColorB      =   16763528
      GradientStyle   =   4
      Caption         =   "Cadastro"
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
      Begin AButtons.AButton cmdFechar 
         Height          =   270
         Left            =   10185
         TabIndex        =   4
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
         Picture         =   "FormChildMaster.frx":058A
         Top             =   45
         Width           =   480
      End
   End
   Begin MSDataGridLib.DataGrid dtgMaster 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3360
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   2143
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
   Begin Insignia_DSControl.ideDSControl dscMaster 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   6
      Top             =   555
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1429
      ButtonsExtras   =   3
      ButtonColor     =   15987699
      BackColor       =   15987699
   End
   Begin MSComctlLib.StatusBar stbFooter 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   5730
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   18
            MinWidth        =   18
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   512
            MinWidth        =   512
            Picture         =   "FormChildMaster.frx":1254
            Key             =   "VerGrid"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Key             =   "Desc1"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Key             =   "Desc2"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   12012
            Key             =   "Desc3"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FormChildMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=-=-=-=-=-=-= Variaveis de controle padrao
Private msLastValue     As String     'Guarda o valor do campo qdo recebe o Focu
Private mnCorLostFocus  As Long       'Guarda o Background do campo qdo recebe o Focu
Private mbFormCarregado As Boolean    'Guarda um flag de que o form está carregado

'=-=-=-=-=-=-= Definição Especifica da Janela
Private Const TBL_MASTER  As String = "NOME_TABELA"
Private Const TAB_TITLE   As String = "TAB_TITLE"

' --- Index dos Campos
Private Const cID       As Byte = 0
Private Const cNome     As Byte = 1
Private Const cCidade   As Byte = 2   'EXEMPLO DE CAMPO DE PESQUISA

Private Type tRSPosition
  AbsolutePosition  As Long
  Bookmark          As Variant
End Type
Private mtRSPosition As tRSPosition

Public Sub ShowForm()
  Dim sSQL As String
  
  Me.MousePointer = vbHourglass

  sSQL = "SELECT * FROM " & TBL_MASTER

  If dscMaster.Conectar(sSQL, gOConn) = cnErroProcesso Then
    Me.MousePointer = vbDefault
    Unload Me
    
  Else
    Load Me
    Call ConfigurarDados
    
    If dscMaster.DataSource.RS.RecordCount > 1 Then dscMaster.DataSource.MoveLast
    If Not mbFormCarregado Then Me.Show
      
    Me.MousePointer = vbDefault
  End If
End Sub

Private Sub ConfigurarDados()
  Dim oT As ActiveText
  Dim sPesq As String, sMask As String

  For Each oT In txtCampo
    Set oT.DataSource = dscMaster.DataSource.RS
    
    Select Case oT.Index
    Case Is = cID, cNome
      Select Case oT.TextMask
      Case Is = [Integer Mask]:   sMask = String(5, "#")
      Case Is = [Float Mask]:     sMask = "ñP"
      Case Else
        sMask = oT.Mask
      End Select
            
      sPesq = sPesq & _
              Replace(lblLabel(oT.Index).Caption, ":", "") & "," & _
              oT.DataField & "," & sMask & "|"
            
    End Select
    
    Set oT = Nothing
    DoEvents
  Next

  Set dtgMaster.DataSource = dscMaster.DataSource.RS
  
  sPesq = Mid$(sPesq, 1, Len(sPesq) - 1)
  dscMaster.MontarPesquisa sPesq
End Sub

Public Property Get TabTitle() As String
  TabTitle = TAB_TITLE
End Property

Public Property Get FormCarregado() As Boolean
   FormCarregado = mbFormCarregado
End Property

Public Property Get OperacaoPendente() As Boolean
  If mbFormCarregado Then
    OperacaoPendente = dscMaster.DataSource.OperacaoPendente
  End If
End Property

Private Sub cmdFechar_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  Call MDIMain.CheckedTBarForms(TypeName(Me))
End Sub

Private Sub Form_Initialize()
  Call mdlGeral.FlatControles(txtCampo)
  Call mdlGeral.FlatControles(txtCampoFK)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF12 Then
    'stbFooter_PanelClick stbFooter.Panels("Desc1")
  Else
    Call modForms.MFormsKeyDown(Me, KeyCode, Shift)
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If dscMaster.DataSource.OperacaoPendente Then
    Cancel = True
    Exit Sub
  End If
  dscMaster.DesConectar  'DESCONECTA OS DADOS
  
  Call MDIMain.TBarDeleteButton(TypeName(Me)) 'REMOVE A ABA
    
  Set FormChildMaster = Nothing
End Sub

Private Sub Form_Resize()
  If Not mbFormCarregado Then
    Call modForms.MFormsResize(Me)
    mbFormCarregado = True
  End If
End Sub

Private Sub PanelTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call mdlGeral.DragForm(Me.hwnd)
End Sub

Private Sub stbFooter_PanelClick(ByVal Panel As MSComctlLib.Panel)
  'If dscMaster.Operacao = opVisualizacao Then
  
    Select Case UCase(Panel.Key)
    Case "VERGRID"
        Call modForms.MFormsShowGrid(Me, Not dtgMaster.Visible)
    
    Case Is = "DESC1"
      'Abre o menu popup para mais opções
'        Panel.Bevel = sbrInset
'        PopupMenu menuPopMenu, , Panel.Left, stbFooter.Top - (mnuPop.Count * 230)
'        Panel.Bevel = sbrRaised
    End Select
    
  'End If
End Sub

Private Sub dtgMaster_DblClick()
  Call MFormsShowGrid(Me, Not dtgMaster.Visible)
End Sub

Private Sub dscMaster_AntesUpdate(Cancel As Boolean, eOperacao As Insignia_DSControl.eDSOperacao)
  Cancel = Not modForms.MFormsValidateRequiredFields(Me.txtCampo)
  If Cancel Then Exit Sub
  
  'Checando informações duplicadas
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
  
End Sub

Private Sub dscMaster_DepoisUpdate(eOperacao As Insignia_DSControl.eDSOperacao)
  If eOperacao = -1 Then
    'Se não existe mais registros atualizar todo o controle
'    dscMaster.DataSource.Requery
  End If
End Sub

Private Sub dscMaster_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  With pRecordset
    If (.BOF Or .EOF) Then
      'Ocorre quando estiver excluindo
      If .AbsolutePosition <> mtRSPosition.AbsolutePosition Or mtRSPosition.Bookmark <> -1 Then
        mtRSPosition.AbsolutePosition = .AbsolutePosition
        mtRSPosition.Bookmark = -1
        
        Call LimparTela(False)
      End If
      
    ElseIf .AbsolutePosition <> mtRSPosition.AbsolutePosition Or .Bookmark <> mtRSPosition.Bookmark Then
      If .EditMode <> adEditAdd Then
        Call ExibirDados(pRecordset)
      End If
      
      mtRSPosition.AbsolutePosition = .AbsolutePosition
      mtRSPosition.Bookmark = .Bookmark
    End If
    
  End With
End Sub

Private Sub dscMaster_Operacao(ByVal eOperacao As Insignia_DSControl.eDSOperacao, ByVal eOperacaoAnterior As Insignia_DSControl.eDSOperacao)
  Dim bEdit As Boolean
  
  bEdit = (eOperacao <> opVisualizacao)
  
  dtgMaster.Enabled = Not bEdit
  mdlGeral.HabilitarEdicao txtCampo, bEdit
  mdlGeral.HabilitarEdicao txtCampoFK, bEdit
  mdlGeral.HabilitarEdicao cmdPesquisa, bEdit
  
  If bEdit Then
    Call MFormsShowGrid(Me, False)
    
    If eOperacao = opInclusao Then
      Call LimparTela(True)
    End If
    
    mdlGeral.SetFocus txtCampo(cNome)
    
  Else
    ' Ocorre sempre que CONFIRMAR OU CANCELAR operacao de inclusao e alteracao
    Call ExibirDados(dscMaster.DataSource.RS)
  End If
  
End Sub

Private Sub ExibirDados(ByRef pRecordset As ADODB.Recordset)
  With pRecordset
    If Not (.EOF Or .BOF) Then
      txtCampoFK(cCidade).Text = modQuerys.NomeCidadeUF(mdlGeral.IfNull(!ID_CIDADE, 0))
    Else
      ' Ocorre em exclusoes ou cancelamento de inclusao
      Call LimparTela(False)
    End If
    
  End With
End Sub

Private Sub LimparTela(ByVal pParaIncluir As Boolean)
  Dim oC As Control
  
  For Each oC In txtCampoFK
'    oC.Text = IIf(txtCampo(oC.Index).Tag = "Obrigatorio", "<OBRIGATÓRIO>", "")
    oC.Text = ""
  Next
  
  If pParaIncluir Then
    ' PREENCHE VALORES DEFAULT
'    txtCampo(cDTCad).Text = mdlGeral.ValorData([DT Atual])
  End If
  
  dscMaster.Informe = ""
End Sub

Private Sub txtCampo_GotFocus(Index As Integer)
  msLastValue = txtCampo(Index).Text
  mnCorLostFocus = txtCampo(Index).BackColor
  txtCampo(Index).BackColor = modConstantes.gcCorFocus
End Sub

Private Sub txtCampo_LostFocus(Index As Integer)
  txtCampo(Index).BackColor = mnCorLostFocus
  If msLastValue = txtCampo(Index).Text Then Exit Sub
  
  Dim bShowMsg As Boolean
   
  Select Case Index
    Case Is = cCidade
      bShowMsg = dscMaster.Operacao <> opVisualizacao
      txtCampoFK(Index).Text = modQuerys.NomeCidadeUF(CLng(txtCampo(Index).Text), bShowMsg)
  End Select
End Sub

Private Sub txtCampo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case Index
    Case Is = cCidade
      If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then cmdPesquisa_Click Index
  End Select
End Sub

Private Sub txtCampo_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case Is = cNome
      Call mdlADO.AutoComplete(txtCampo(Index), txtCampo(Index).DataField, TBL_MASTER, KeyAscii)
  End Select
End Sub

Private Sub txtCampoFK_GotFocus(Index As Integer)
  msLastValue = txtCampoFK(Index).Text
  mnCorLostFocus = txtCampoFK(Index).BackColor
  txtCampoFK(Index).BackColor = modConstantes.gcCorFocus
End Sub

Private Sub txtCampoFK_LostFocus(Index As Integer)
  txtCampoFK(Index).BackColor = mnCorLostFocus
  If msLastValue = txtCampoFK(Index).Text Then Exit Sub
  
  Select Case Index
  Case Is = cCidade
    txtCampo(Index).Text = modQuerys.BuscarCodigoID(tbCidadeUF, "NOME", txtCampoFK(Index).Text)
  End Select
    
  If txtCampo(Index).Text = 0 And Trim(txtCampoFK(Index).Text) <> "" Then
    mdlGeral.MGShowInfo "Registro não encontrado!"
    txtCampoFK(Index).SetFocus
  End If
End Sub

Private Sub txtCampoFK_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim sTBL As String

  Select Case Index
  Case Is = cCidade:  sTBL = "TB_CIDADES"
  End Select
  
  If (sTBL <> "") Then
    Call mdlADO.AutoComplete(txtCampoFK(Index), txtCampoFK(Index).DataField, sTBL, KeyAscii)
  End If
End Sub

Private Sub cmdPesquisa_Click(Index As Integer)
  Dim sResult As String
  
  Select Case Index
  Case Is = cCidade
    sResult = modFormPesquisa.Cidade
  End Select
  
  If sResult <> "" Then
    txtCampo(Index).Text = sResult
    txtCampo(Index).SetFocus
  End If
End Sub
