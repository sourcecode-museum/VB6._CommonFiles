VERSION 5.00
Begin VB.Form FormCadastroPadrao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBotoes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   480
      TabIndex        =   11
      Top             =   5955
      Width           =   8025
      Begin VB.CommandButton cmdBotoes 
         Caption         =   "Imprimir"
         Height          =   800
         Index           =   5
         Left            =   5310
         Picture         =   "FormCadastroPadrao.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdBotoes 
         Caption         =   "Cancelar"
         Height          =   800
         Index           =   1
         Left            =   6240
         Picture         =   "FormCadastroPadrao.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdBotoes 
         Caption         =   "Excluir"
         Height          =   800
         Index           =   8
         Left            =   1710
         Picture         =   "FormCadastroPadrao.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdBotoes 
         Caption         =   "Salvar"
         Height          =   800
         Index           =   7
         Left            =   855
         Picture         =   "FormCadastroPadrao.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdBotoes 
         Caption         =   "Novo"
         Height          =   800
         Index           =   6
         Left            =   0
         Picture         =   "FormCadastroPadrao.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdBotoes 
         Caption         =   "Próximo"
         Height          =   800
         Index           =   4
         Left            =   4365
         Picture         =   "FormCadastroPadrao.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdBotoes 
         Caption         =   "Anterior"
         Height          =   800
         Index           =   3
         Left            =   3510
         Picture         =   "FormCadastroPadrao.frx":34BC
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdBotoes 
         Caption         =   "Pesquisar"
         Height          =   800
         Index           =   2
         Left            =   2655
         Picture         =   "FormCadastroPadrao.frx":3D86
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdBotoes 
         Cancel          =   -1  'True
         Caption         =   "Fechar"
         Height          =   800
         Index           =   0
         Left            =   7110
         Picture         =   "FormCadastroPadrao.frx":4650
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame fraForm 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5880
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8505
      Begin VB.TextBox txtCod 
         Height          =   315
         Index           =   0
         Left            =   1305
         TabIndex        =   9
         Top             =   630
         Width           =   1000
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   2685
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   630
         Width           =   5490
      End
      Begin VB.CommandButton cmdPes 
         Height          =   315
         Index           =   0
         Left            =   2340
         Picture         =   "FormCadastroPadrao.frx":4F1A
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   630
         Width           =   315
      End
      Begin VB.TextBox txtCodSEQ 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1305
         MaxLength       =   38
         TabIndex        =   4
         ToolTipText     =   "Código do registro"
         Top             =   180
         Width           =   1350
      End
      Begin VB.CommandButton cmdPes 
         Height          =   315
         Index           =   1
         Left            =   2340
         Picture         =   "FormCadastroPadrao.frx":54A4
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1275
         Width           =   315
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   2685
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1275
         Width           =   5490
      End
      Begin VB.TextBox txtCod 
         Height          =   315
         Index           =   1
         Left            =   1305
         TabIndex        =   1
         Top             =   1275
         Width           =   1000
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         Height          =   195
         Index           =   6
         Left            =   330
         TabIndex        =   10
         Top             =   690
         Width           =   495
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   7
         Left            =   330
         TabIndex        =   5
         Top             =   1335
         Width           =   555
      End
   End
End
Attribute VB_Name = "FormCadastroPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'** Objetos de Controle de Tela (Padrão) ***************

Private WithEvents moInterface      As CFormCadastro
Attribute moInterface.VB_VarHelpID = -1
Private moControles                 As cControles
Private moRSTela                    As ADODB.Recordset
'---------------------------------------------------------------------------------------

Private Const cCampoBanco = 0
Private Const cCampoCliente = 1


Private Sub cmdBotoes_Click(Index As Integer)
    Call moInterface.ButtonClick(Index)
End Sub

Private Sub cmdPes_Click(Index As Integer)
    Select Case Index
    Case 0, 1 'CLIENTE e BANCO
        modPesquisa.PesqCliente txtCod(Index), txtDesc(Index), IIf(Index = 0, "Consulta de Bancos", "")
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    moInterface.KeyDown KeyCode, Shift
End Sub

Private Sub Form_Load()

 On Error GoTo Err_Trat

    gs_IniciaFormulario Me
    Screen.MousePointer = vbHourglass

    Me.Icon = SCO_Mdi.Icon
    Senha = Senha_Global(mcNUMMOD)

    fraBotoes.BackColor = Me.BackColor
    
    Set moInterface = New clsInterfaceTela
    With moInterface
        .PermitirApenasConsulta = Not Senha

        Call .SetarBotoes(cmdBotoes)
        Call .EnabledButtons(opNenhum)

        moInterface_LimparTela
    End With

    Set moControles = New cControles
    With moControles
        .AutoSelecionar = True
        .CorSelecao = tPrefUsuario.tConfigGerais.CorDoFoco
        .AddControls Me
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
Err_Trat:

    Err_SCO Err, False, "Form_Load", Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With moInterface
        Cancel = Not (.Status = opNenhum Or .Status = opVisualizar)
    End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With fraBotoes
        .Move Me.Width - (.Width + 100), Me.Height - (.Height * 1.6)
    End With
     With fraForm
        .Move 15, 0, Me.Width - 130, fraBotoes.Top - 80
    End With
On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set moInterface = Nothing
    Set moControles = Nothing
    Set moRSTela = Nothing

    Set FormCadastroPadrao = Nothing
End Sub

Private Sub moInterface_Cancelar()
On Local Error Resume Next
    moRSTela.CancelBatch adAffectAllChapters
End Sub

Private Sub moInterface_Carregar()
On Error GoTo TrataErro
    If Not RecordSetOK(moRSTela) Then Exit Sub

    lblCaption(8).Visible = False
    TDBDDatMovfFim.Visible = False

    Screen.MousePointer = vbHourglass

    With moRSTela
        If .RecordCount > 0 Then
            txtCodSEQ.Text = .Fields("SEQARQ")
            txtDATMOV.value = .Fields("DATMOV")

            txtCod(cCampoBanco).Text = .Fields("CODCLIPRO") & ""
            txtDesc(cCampoBanco).Text = .Fields("NOMBCO") & ""
            txtCod(cCampoCliente).Text = .Fields("CODCLI") & ""
            txtDesc(cCampoCliente).Text = .Fields("NOMCLI") & ""

            txtNome.Text = .Fields("NOMARQ")

            lblUsuarioINC.Caption = .Fields("NOMUSUINC") & ""
            lblDataINC.Caption = .Fields("DATINC") & ""
            lblUsuarioALT.Caption = .Fields("NOMUSUALT") & ""
            lblDataALT.Caption = .Fields("DATALT") & ""

            lblPosicaoRS.Caption = modFuncoes.TraducaoMsg("@0 de @1", .AbsolutePosition, .RecordCount)
        End If
    End With

Sair:
    Screen.MousePointer = vbDefault
    Exit Sub

TrataErro:
    Err_SCO Err, , "moInterface_Carregar", Me
    GoTo Sair
End Sub

Private Sub moInterface_DepoisDoClick(ByVal Botao As eButtonClick)
    If Botao = bcMoveANT Or Botao = bcMovePRO Then
        modFuncoes.SelecionarLinhaGrid TDBGrid
    End If
End Sub

Private Sub moInterface_Excluir(TemRegistro As Boolean, Cancelar As Boolean)
On Error GoTo TrataErro
    Dim oFC As Object

    Screen.MousePointer = vbHourglass

    Cancelar = MsgBox(modFuncoes.Traduz("Confirma [ EXCLUSÃO ] do registro?"), vbQuestion + vbYesNo + vbDefaultButton2, modFuncoes.Traduz("Excluir")) = vbNo
    If Not Cancelar Then
        Set oFC = CreateObjectFC(mcObjetoFC)

        Call oFC.RemoverPorChave(Val(txtCodSEQ.Text), Codigo_Filial)
        moRSTela.Delete

        TemRegistro = moRSTela.RecordCount > 0
    End If

Sair:
    Set oFC = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

TrataErro:
    Cancelar = True
    Err_SCO Err, , "moInterface_Excluir", Me
    GoTo Sair
End Sub

Private Sub moInterface_Fechar()
    Unload Me
End Sub

Private Sub moInterface_LimparTela()
    Call moInterface.LimparCampos(Me)

    lblCaption(8).Visible = True
    TDBDDatMovfFim.Visible = True
    lblPosicaoRS.Caption = modFuncoes.TraducaoMsg("@0 de @1", 0, 0)

    lblUsuarioINC.Caption = ""
    lblDataINC.Caption = ""
    lblUsuarioALT.Caption = ""
    lblDataALT.Caption = ""
End Sub

Private Sub moInterface_Navegacao(adoRS As ADODB.Recordset)
    Set adoRS = moRSTela
End Sub

Private Sub moInterface_Pesquisar(Cancelar As Boolean)
On Error GoTo TrataErro
    Dim oFC     As Object
    Dim sSQL    As String

    If Not pf_ValidarPesquisa Then
        Cancelar = True
        GoTo Sair
    End If

    Screen.MousePointer = vbHourglass

    Set oFC = CreateObjectFC(mcObjetoFC)
    Set moRSTela = oFC.BuscarPorFiltroRelacionado(Codigo_Filial, WhereStr)

    Cancelar = (moRSTela.EOF)
    If Cancelar Then
        MsgBox modFuncoes.Traduz("Nenhum registro encontrado."), vbExclamation, modFuncoes.Traduz("Pesquisa")
        SetarFocu txtCodSEQ
    End If

Sair:
    Set oFC = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

TrataErro:
    Cancelar = True
    Err_SCO Err, , "moInterface_Pesquisar", Me
    GoTo Sair
End Sub

Private Sub moInterface_Salvar(ByVal pAddNew As Boolean, Cancelar As Boolean)
On Error GoTo TrataErro
    Dim oFC     As Object
    Dim oRS     As ADODB.Recordset
    Dim sArq    As String

    sArq = txtArquivo.Text
    If sArq = "" Or Dir(sArq, vbArchive) = "" Then sArq = ""

    If pAddNew Then
        If sArq = "" Then
            Cancelar = True

            MsgBox modFuncoes.Traduz("Caminho do arquivo de origem deve ser informado!"), vbExclamation, modFuncoes.Traduz("Salvar")
            txtArquivo.SetFocus
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbHourglass

    Set oFC = CreateObjectFC(mcObjetoFC)
    If pAddNew Then
        Set oRS = oFC.BuscarNovo(Codigo_Filial)

        oRS.AddNew

        oRS("CODFIL") = Codigo_Filial
        oRS("NOMUSUINC") = Nome_Usuario
        oRS("DATINC") = Now
    Else
        Set oRS = oFC.BuscarPorChave(Val(txtCodSEQ.Text), Codigo_Filial)

        oRS("NOMUSUALT") = Nome_Usuario
        oRS("DATALT") = Now
    End If

    'Atualizando para mostrar
    With oRS
        .Fields("NOMARQ") = txtNome.Text
        .Fields("DATMOV") = CDate(Format(txtDATMOV.value, "DD/MM/YYYY"))
        .Fields("CODCLIPRO") = txtCod(cCampoBanco).Text
        .Fields("CODCLI") = txtCod(cCampoCliente).Text
        .Fields("VISPAF") = 1

        If sArq <> "" Or pAddNew Then
            .Fields("ARQBIN").AppendChunk modFuncoes.ArqBinario2Array(sArq)
        End If
    End With

    ' Monta email de inclusão de Arquivo GCD
    Dim rsEmail As ADODB.Recordset
    Set rsEmail = EnviarEmailCliente(txtCod(cCampoCliente).Text)

    If RecordSetOK(rsEmail) Then
        Call oFC.Salvar(oRS, Codigo_Filial, rsEmail)
    Else
        Call oFC.Salvar(oRS, Codigo_Filial)
    End If

    MsgBox modFuncoes.TraducaoMsg("Registro [ @0 ] com sucesso!", IIf(pAddNew, UCase$(modFuncoes.Traduz("Incluso")), UCase$(modFuncoes.Traduz("Alterado")))), vbInformation, modFuncoes.Traduz("Salvar")

    'Sicronizando os valores alterados no RecordSet da Tela
    If moInterface.Status = opEditar Then
        On Local Error GoTo Sair
        Call moInterface.SincronizarDados(oRS.Fields, moRSTela.Fields)

        'Atualizando os valores do campos secundarios
        moRSTela.Fields("NOMBCO") = txtDesc(cCampoBanco).Text
        moRSTela.Fields("NOMCLI") = txtDesc(cCampoCliente).Text
        moRSTela.UpdateBatch adAffectAllChapters
    End If

Sair:
    FechaRegistros rsEmail
    Set oFC = Nothing
    Set oRS = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    Resume
TrataErro:
    Cancelar = True
    Err_SCO Err, , "moInterface_Salvar", Me
    GoTo Sair
End Sub

Private Sub moInterface_StatusOperacao(ByVal NewStatus As eOperacao, ByVal OldStatus As eOperacao)
    TDBGrid.Enabled = Not (NewStatus = opIncluir Or NewStatus = opEditar Or NewStatus = opNenhum)
    txtCodSEQ.Enabled = (NewStatus = opNenhum)
    cmdDialog(0).Enabled = Not (NewStatus = opNenhum)
    cmdDialog(1).Enabled = Not (NewStatus = opNenhum Or NewStatus = opIncluir)

    Select Case NewStatus
    Case opNenhum
        txtCodSEQ.SetFocus
        Set moRSTela = Nothing
        Set TDBGrid.DataSource = Nothing

    Case opVisualizar
        txtDATMOV.SetFocus
        Set TDBGrid.DataSource = moRSTela

    Case opIncluir
        txtCodSEQ.Text = "{" & modFuncoes.Traduz("Auto") & "}"
        lblCaption(8).Visible = False
        TDBDDatMovfFim.Visible = False
        txtDATMOV.value = Data_Default
        txtDATMOV.SetFocus
    End Select
End Sub

Private Sub moInterface_ValidarSalvar(Cancelar As Boolean)

    Dim cCap As String
    cCap = modFuncoes.Traduz("Validando informações...")

    Cancelar = True

    Select Case True
    Case Not IsDate(txtDATMOV.value)
        MsgBox modFuncoes.Traduz("Campo DATA MOVIMENTAÇÃO deve ser informado!"), vbExclamation, cCap
        txtDATMOV.SetFocus

    Case Not IsNumeric(txtCod(cCampoBanco).Text)
        MsgBox modFuncoes.Traduz("Campo BANCO deve ser informado!"), vbExclamation, cCap
        txtCod(cCampoBanco).SetFocus

    Case Not IsNumeric(txtCod(cCampoCliente).Text)
        MsgBox modFuncoes.Traduz("Campo CLIENTE deve ser informado!"), vbExclamation, cCap
        txtCod(cCampoCliente).SetFocus

    Case Trim$(txtNome.Text) = ""
        MsgBox modFuncoes.Traduz("Campo NOME deve ser informado!"), vbExclamation, cCap
        txtNome.SetFocus
    Case Else
        Cancelar = False
    End Select
End Sub

Private Sub TDBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'Atualiza os dados dos campos quando movimenta atravez do grid
    If Not (moInterface.Status = opNenhum) Then
        If TDBGrid.Bookmark <> "" Then
            moRSTela.Bookmark = TDBGrid.Bookmark
            Call moInterface.Carregar
        End If
    End If
End Sub

Private Sub TDBGrid_HeadClick(ByVal ColIndex As Integer)
    Call modFuncoes.OrdenaTDBGridDS(TDBGrid, TDBGrid.DataSource, TDBGrid.Columns(ColIndex).DataField)
End Sub

Private Sub txtCod_Change(Index As Integer)
    moInterface.Editar
End Sub

Private Sub txtCod_GotFocus(Index As Integer)
    txtCod(Index).Tag = txtCod(Index).Text
End Sub

Private Sub txtCod_KeyPress(Index As Integer, KeyAscii As Integer)
    modFuncoes.ValidaNumeros (KeyAscii)
End Sub

Private Sub txtCod_LostFocus(Index As Integer)
    With txtCod(Index)
        If Trim$(.Text) = "" Then
            txtDesc(Index).Text = ""
        ElseIf .Text <> .Tag Then
            cmdPes_Click Index
        End If
    End With
End Sub

Private Sub txtCodSEQ_KeyPress(KeyAscii As Integer)
    Call modFuncoes.ValidaNumeros(KeyAscii)
End Sub

Private Sub txtDATMOV_Change()
    moInterface.Editar
End Sub

Private Sub txtArquivo_Change()
    moInterface.Editar
End Sub

Private Sub txtDesc_GotFocus(Index As Integer)
    txtDesc(Index).Tag = txtDesc(Index).Text
End Sub

Private Sub txtDesc_KeyPress(Index As Integer, KeyAscii As Integer)
    Call modFuncoes.CampoEmMaiusculo(KeyAscii)
End Sub

Private Sub txtDesc_LostFocus(Index As Integer)
    With txtDesc(Index)
        If Trim$(.Text) = "" Then
            txtCod(Index).Text = ""
        ElseIf .Text <> .Tag Then
            txtCod(Index).Text = ""
            cmdPes_Click Index
        End If
    End With
End Sub

Private Sub txtNome_Change()
    Call moInterface.Editar
End Sub

Private Function pf_ValidarPesquisa() As Boolean

    Dim cCap As String
    cCap = modFuncoes.Traduz("Validando informações...")

    pf_ValidarPesquisa = False

    Select Case True
    Case txtDATMOV.ValueIsNull
        MsgBox modFuncoes.Traduz("Campo DATA MOVIMENTAÇÃO deve ser informado!"), vbExclamation, cCap
        txtDATMOV.SetFocus
    Case Else
        pf_ValidarPesquisa = True
    End Select
End Function

Public Function EnviarEmailCliente(ByVal pCODCLI As String) As Object
On Error GoTo Err_Trat

    Dim rsParam     As New ADODB.Recordset
    Dim rsEmail     As New ADODB.Recordset
    Dim objFilEma   As Object
    Dim sMsg        As String

    ' Recupera os parâmetros nacessários para o envio do Email
    Set rsParam = gF_CarregaParametroExternoClienteGenerico(Codigo_Filial, txtCod(cCampoCliente).Text, "146")

    ' Valida a quantidade de parâmetros encontrados
    If RecordSetOK(rsParam) Then
        If rsParam.RecordCount = 6 Then

            'Busca a estrutura da tabela de cadastro de e-mail
            Set objFilEma = CreateObjectFC(cstCompFachTSA & ".FACH_SCO_TFILEMA")
            Set rsEmail = objFilEma.BuscarNovo(Codigo_Filial)

            'Monta Recordset para enviar os email's
            rsParam.MoveFirst   ' 1º Registro
            With rsEmail
                .AddNew
                .Fields("DATINC").value = CDate(Now)
                .Fields("DE").value = "aviso.prosegur@br.prosegur.com"
                rsParam.MoveNext    ' 2º Registro
                .Fields("ASSUNTO").value = modFuncoes.Traduz(rsParam(1))
                rsParam.MoveNext    ' 3º Registro
                .Fields("PARA").value = rsParam(1)
                rsParam.MoveNext    ' 4º Registro
                sMsg = modFuncoes.Traduz(rsParam(1))
                sMsg = sMsg & vbCrLf & modFuncoes.TraducaoMsg("Nome do Arquivo: @0", txtNome.Text)
                sMsg = sMsg & vbCrLf & modFuncoes.TraducaoMsg("Data carregamento: @0", Format(Now, "dd/mm/yyyy hh:mm:ss"))
                rsParam.MoveNext    ' 5º Registro
                sMsg = sMsg & vbCrLf & vbCrLf & rsParam(1)
                .Fields("MSG").value = sMsg
                .Fields("STATUS").value = "A"
                .Fields("ORIGEM").value = "SCO"
                rsParam.MoveNext    ' 6º Registro
                .Fields("EMAILHTML").value = rsParam(1)
                .Update
            End With
        End If
    End If

    Set EnviarEmailCliente = rsEmail

Fim:
    FechaRegistros rsParam
    Set objFilEma = Nothing

    Exit Function
    Resume
Err_Trat:
    Err_SCO Err, False, "modFuncoes.EnviarEmailCliente", Me
    GoTo Fim
End Function
