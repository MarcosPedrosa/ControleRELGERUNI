VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.MDIForm MDI_Ger_Info 
   BackColor       =   &H8000000C&
   Caption         =   "Controle de Distribuição"
   ClientHeight    =   7530
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11880
   Icon            =   "MDI_Ger_Info.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   1260
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   7185
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Versão"
            TextSave        =   "Versão"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "usuário"
            TextSave        =   "usuário"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Banco"
            TextSave        =   "Banco"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   150
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDI_Ger_Info.frx":1D2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnu_principal 
      Caption         =   "Tabelas"
      Index           =   0
      Begin VB.Menu Mnu_Tabelas 
         Caption         =   "Usuário..."
         Index           =   1
      End
   End
   Begin VB.Menu Mnu_principal 
      Caption         =   "Movimentação"
      Index           =   1
      Begin VB.Menu Mnu_MovDisribuicao 
         Caption         =   "Cancelamento Nota Fiscal..."
         Index           =   1
      End
      Begin VB.Menu Mnu_MovDisribuicao 
         Caption         =   "Troca de caixa..."
         Index           =   2
      End
      Begin VB.Menu Mnu_MovDisribuicao 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Mnu_MovDisribuicao 
         Caption         =   "Abertura de inventário..."
         Index           =   4
      End
      Begin VB.Menu Mnu_MovDisribuicao 
         Caption         =   "Importação mov. inventário..."
         Index           =   5
      End
      Begin VB.Menu Mnu_MovDisribuicao 
         Caption         =   "Gerar Pallet's para R3..."
         Index           =   6
      End
   End
   Begin VB.Menu Mnu_principal 
      Caption         =   "Relatórios"
      Index           =   2
      Begin VB.Menu Mnu_RelDisribuicao 
         Caption         =   "Composição do Pallet's..."
         Index           =   1
      End
      Begin VB.Menu Mnu_RelDisribuicao 
         Caption         =   "Localização de peças..."
         Index           =   2
      End
      Begin VB.Menu Mnu_RelDisribuicao 
         Caption         =   "Informação da caixa..."
         Index           =   3
      End
      Begin VB.Menu Mnu_RelDisribuicao 
         Caption         =   "Informação da sequência do caminhão..."
         Index           =   4
      End
      Begin VB.Menu Mnu_RelDisribuicao 
         Caption         =   "Composição da Nota Fiscal..."
         Index           =   5
      End
   End
   Begin VB.Menu Mnu_principal 
      Caption         =   "Integração Unimed"
      Index           =   3
      Begin VB.Menu Mnu_Unimed 
         Caption         =   "Importação Mov. Unimed..."
         Index           =   1
      End
      Begin VB.Menu Mnu_Unimed 
         Caption         =   "Digitação desconto extra no Mov..."
         Index           =   2
      End
      Begin VB.Menu Mnu_Unimed 
         Caption         =   "Calculo Mensal..."
         Index           =   3
      End
      Begin VB.Menu Mnu_Unimed 
         Caption         =   "Exportar Mov. P/Folha de Pagto..."
         Index           =   4
      End
      Begin VB.Menu Mnu_Unimed 
         Caption         =   "Abertura/Atualização parametros..."
         Index           =   5
      End
      Begin VB.Menu Mnu_Unimed 
         Caption         =   "Consulta Saldos..."
         Index           =   6
      End
   End
   Begin VB.Menu Mnu_principal 
      Caption         =   "Utilitários"
      Index           =   4
      Begin VB.Menu Mnu_Utilitario 
         Caption         =   "Troca de Usuário..."
         Index           =   0
      End
      Begin VB.Menu Mnu_Utilitario 
         Caption         =   "Manutenção Menu sistema..."
         Index           =   1
      End
      Begin VB.Menu Mnu_Utilitario 
         Caption         =   "Log do Sistema..."
         Index           =   2
      End
   End
   Begin VB.Menu Mnu_principal 
      Caption         =   "Saida"
      Index           =   5
   End
   Begin VB.Menu janela_mdi 
      Caption         =   "Janela"
      Index           =   0
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDI_Ger_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'192.168.0.1
Private m_BGnd As CMdiBackground
Private bMenuAtivo As Boolean
Private Sub MDIForm_Load()
'Dim formclass As Object
'
'On Error GoTo Erro
'
'    Set formclass = New frmGIFIUniImportacaoUnimed
'    formclass.Show 0
'
'Set formclass = Nothing
'
'Exit Sub
'
'
'
'
'
'
'
'
'
'
Dim dDate As Date

On Error GoTo Erro

'App.HelpFile = App.Path & "\help\Aeshelp.hlp"
    
If App.PrevInstance Then
   MsgBox "Este Programa JÁ esta sendo processado neste computador", 16, "<ENTER>=Para Finalizar"
   Close: End
End If

Me.MousePointer = vbHourglass

Rem teste da data do computador para o formato dd/mm/yyyy

If Len(Now()) < 19 Then
   MsgBox "O seu computador está com o formato da DATA DIFERENTE DO PADRÃO dd/mm/yyyy. Altere as Configurações Regionais , no Painel de Controle."
   End
End If

dDate = Mid(Now(), 1, 10)
If Len(Trim(dDate)) <> 10 Then
   MsgBox "O seu computador está com o formato da DATA DIFERENTE DO PADRÃO dd/mm/yyyy. Altere as Configurações Regionais , no Painel de Controle."
   End
End If

'' Digitacao do Banco e usuario
'
If Not Entrada_Sistema Then End

Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
Unload Me
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    Dim Form As Form
    For Each Form In Forms
        If Form Is Me Then
            Set Form = Nothing
            Exit For
        End If
    Next Form
    End
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If 7 = MsgBox("Deseja realmente sair do sistema?", 32 + 4) Then
            'respondeu não
            Cancel = True
        End If
    End If

End Sub

Private Sub MDIForm_Activate()
Dim i As Long
Dim cRec As ADODB.Recordset
Dim cFields As Collection
Dim sCaminhoLogo As String

On Error GoTo Erro

Call Atualizar_Menu_Usuario

Me.StatusBar1.Panels(1).Text = "Versão " & App.Major & "." & Format(App.Revision, "000")

If Not bMenuAtivo Then
   bMenuAtivo = True
   Set m_BGnd = New CMdiBackground
   With m_BGnd
      Set .Client = Me
      .AutoRefresh = True
   End With
   
   Rem leitura do cadastro de etiquetas para pegar os clientes
   
   Set rRec_cliente = New ADODB.Recordset
   
   Me.MousePointer = vbHourglass
   
   Set rRec_cliente = CCTemp.EXPEDICAO_Consultar_Cliente(sNomeBanco)
   
   Me.MousePointer = vbDefault
   
'   sCaminhoLogo = App.Path & "\logo.bmp"
   
'   Me.Image1.Picture = LoadPicture(sCaminhoLogo)
      
'   Set m_BGnd.Graphic = Me.Image1
'   m_BGnd.GraphicPosition = mdiStretched
'
'   Me.MousePointer = vbHourglass
'   Set cRec = New ADODB.Recordset
   
'   Set crec = CCTempneMovManPreventiva.MovManPreventiva_Cons_Vencida()
   
'   If cRec.RecordCount > 0 Then
'      Me.TlbMenu.Buttons(7).Visible = True
'   End If

' Digitacao do Banco e usuario
'If Not Entrada_Sistema Then End
   
   Me.MousePointer = vbDefault
'   Set crec = Nothing
   Rem habilitar pendencias de prevencao do controle de frotas

End If

Exit Sub

Erro:
Set cRec = Nothing
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub Mnu_MovDisribuicao_Click(Index As Integer)
Dim formclass As Object

If Index = 1 Then
    Set formclass = New frmGIFMntCancelaNF 'Cancelamento Nota Fiscal
    formclass.Show 0
ElseIf Index = 2 Then
    Set formclass = New frmGIFMntTrocaDeCaixa
    formclass.Show 0
ElseIf Index = 4 Then
    Set formclass = New frmGIFMntAberturaInventário
    formclass.Show 0
ElseIf Index = 5 Then
    Set formclass = New frmGIFMntImportaMovInventário
    formclass.Show 0
ElseIf Index = 6 Then
    Set formclass = New frmGIFMntGerarPalletR3
    formclass.Show 0
End If

Set formclass = Nothing

End Sub

Private Sub Mnu_principal_Click(Index As Integer)
Dim formclass As Object

If Index = 5 Then
    End: Close
End If

Set formclass = Nothing
End Sub

Private Sub Mnu_RelDisribuicao_Click(Index As Integer)
Dim formclass As Object

If Index = 1 Then
    Set formclass = New frmGIFDisComposicaopallet
    formclass.Show 0
ElseIf Index = 2 Then
    Set formclass = New frmGIFDisLocalizacaoPecas
    formclass.Show 0
ElseIf Index = 3 Then
    Set formclass = New frmGIFDisLocalizacaCaixa
    formclass.Show 0
ElseIf Index = 4 Then
    Set formclass = New frmGIFDisInfoSeqCaminhao
    formclass.Show 0
ElseIf Index = 5 Then
    Set formclass = New frmGIFDisComposicaoNotaFiscal
    formclass.Show 0
End If

Set formclass = Nothing
End Sub

Private Function Entrada_Sistema() As Boolean

'Dim cRecAux As ADODB.Recordset
Dim nx As Integer
Dim otemp As ADODB.Recordset
Dim cCodigo As String
Dim oTela As frmGIFDisUsuarioBanco

Me.MousePointer = vbHourglass
On Error GoTo Erro
Set oTela = New frmGIFDisUsuarioBanco

oTela.sChave = "SEN_SISTEMA"

oTela.Show 1
If oTela.TXTSENHA = "" Then
    Set oTela = Nothing
    Me.MousePointer = vbDefault
    Entrada_Sistema = False
    Exit Function
Else
    If Not oTela.bAchou Then
       Set oTela = Nothing
       Me.MousePointer = vbDefault
       Unload Me
       Entrada_Sistema = False
       Exit Function
    Else
'       Unload oTela
       Set oTela = Nothing
       Me.MousePointer = vbDefault
       Entrada_Sistema = True
    End If
End If

Entrada_Sistema = True
Me.StatusBar1.Panels(2).Text = sNome_Usuario
Me.StatusBar1.Panels(3).Text = sNomeEmpresa

Rem CRIAR TABELAS
''''Call CCTempneCriarTabelas.Criar_Tabelas_dados(sNomeBanco)

Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
Unload Me
End Function

Private Sub Mnu_Tabelas_Click(Index As Integer)
Dim formclass As Object

If Index = 1 Then
    Set formclass = New frmGIFDisUsuario
    formclass.Show 0
'ElseIf Index = 1 Then
'    Set formclass = New frmGIFDisComposicaopallet
'    formclass.Show 0
End If

Set formclass = Nothing

End Sub

Private Sub Mnu_Unimed_Click(Index As Integer)

Dim formclass As Object

On Error GoTo Erro

If Index = 1 Then
    Set formclass = New frmGIFIUniImportacaoUnimed
    formclass.Show 0
ElseIf Index = 2 Then
    Set formclass = New frmGIFIUniDigitacaoDescontoExtra
    formclass.Show 0
ElseIf Index = 3 Then
    Set formclass = New frmGIFIUniCalculoUnimed
    formclass.Show 0
ElseIf Index = 4 Then
    Set formclass = New frmGIFIUniExportacaoUnimed
    formclass.Show 0
ElseIf Index = 5 Then
    Set formclass = New frmGIFIUniAberturaUnimed
    formclass.Show 0
ElseIf Index = 6 Then
    Set formclass = New frmGIFIUniConsultarSaldo
    formclass.Show 0
End If

Set formclass = Nothing


Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
Unload Me
End Sub

Private Sub Mnu_Utilitario_Click(Index As Integer)
Dim formclass As Object

On Error GoTo Erro

If Index = 0 Then
    Unload Me
    bMenuAtivo = True
    Me.Show
    Me.SetFocus
ElseIf Index = 1 Then
    Set formclass = New frmGIFMntFormularios
    formclass.Show 0
ElseIf Index = 2 Then
    Set formclass = New frmGIFMntLogSistema
    formclass.Show 0
End If

Set formclass = Nothing


Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
Unload Me
End Sub

Private Function Atualizar_Menu_Usuario()
Dim cRec As ADODB.Recordset
Dim nx As Integer
Dim sGrupo As String

'On Error GoTo Erro
On Error Resume Next

Me.MousePointer = vbHourglass

Set cRec = New ADODB.Recordset

Set cRec = CCTempneUsuario.USUARIO_Consultar_Acesso(sNomeBanco, sUsuario)

While Not cRec.EOF
      
''**********************************************************************************************
      If cRec!FOR_NOME_EDITOR = "Mnu_Principal" Then
         Me.Mnu_principal(cRec!FRM_ACESSO).Visible = cRec!FRM_VISUALIZA
         If cRec!FRM_VISUALIZA = 1 Then
            Me.Mnu_principal(cRec!FRM_ACESSO).Enabled = cRec!FRM_HABILITAR
         End If
      End If
''**********************************************************************************************
      If cRec!FOR_NOME_EDITOR = "Mnu_Tabelas" Then
         Me.Mnu_Tabelas(cRec!FRM_ACESSO).Visible = cRec!FRM_VISUALIZA
         If cRec!FRM_VISUALIZA = 1 Then
            Me.Mnu_Tabelas(cRec!FRM_ACESSO).Enabled = cRec!FRM_HABILITAR
         End If
      End If
''**********************************************************************************************
      If cRec!FOR_NOME_EDITOR = "Mnu_MovDisribuicao" Then
'         Me.mnuArmRel(cRec!FRM_ACESSO).Visible = 1
'         Me.mnuArmRel(cRec!FRM_ACESSO).Enabled = 1
         Me.Mnu_MovDisribuicao(cRec!FRM_ACESSO).Visible = cRec!FRM_VISUALIZA
         If cRec!FRM_VISUALIZA = 1 Then
            Me.Mnu_MovDisribuicao(cRec!FRM_ACESSO).Enabled = cRec!FRM_HABILITAR
         End If
      End If
''**********************************************************************************************
      If cRec!FOR_NOME_EDITOR = "Mnu_RelDisribuicao" Then
         Me.Mnu_RelDisribuicao(cRec!FRM_ACESSO).Visible = cRec!FRM_VISUALIZA
         If cRec!FRM_VISUALIZA = 1 Then
            Me.Mnu_RelDisribuicao(cRec!FRM_ACESSO).Enabled = cRec!FRM_HABILITAR
         End If
      End If
''**********************************************************************************************
      If cRec!FOR_NOME_EDITOR = "Mnu_Utilitario" Then
         Me.Mnu_Utilitario(cRec!FRM_ACESSO).Visible = cRec!FRM_VISUALIZA
         If cRec!FRM_VISUALIZA = 1 Then
            Me.Mnu_Utilitario(cRec!FRM_ACESSO).Enabled = cRec!FRM_HABILITAR
         End If
      End If
''**********************************************************************************************
      If cRec!FOR_NOME_EDITOR = "Mnu_Unimed" Then
         Me.Mnu_Unimed(cRec!FRM_ACESSO).Visible = cRec!FRM_VISUALIZA
         If cRec!FRM_VISUALIZA = 1 Then
            Me.Mnu_Unimed(cRec!FRM_ACESSO).Enabled = cRec!FRM_HABILITAR
         End If
      End If
''**********************************************************************************************
'      If crec!FOR_NOME_EDITOR = "Mnusub" Then
'         Me.MnuSub(crec!FRM_ACESSO).Visible = crec!FRM_VISUALIZA
'         If crec!FRM_VISUALIZA = 1 Then
'            Me.MnuSub(crec!FRM_ACESSO).Enabled = crec!FRM_HABILITAR
'         End If
'
'         If crec!FRM_ACESSO = 1 Then
'            Me.TlbMenu.Buttons.Item(5).Enabled = crec!FRM_HABILITAR
'            If crec!FRM_VISUALIZA = 1 Then
'               Me.TlbMenu.Buttons.Item(5).Visible = crec!FRM_VISUALIZA
'            End If
'         End If
'         If crec!FRM_ACESSO = 2 Then
'            Me.TlbMenu.Buttons.Item(6).Enabled = crec!FRM_HABILITAR
'            If crec!FRM_VISUALIZA = 1 Then
'               Me.TlbMenu.Buttons.Item(6).Visible = crec!FRM_VISUALIZA
'            End If
'         End If
'         If crec!FRM_ACESSO = 3 Then
'            Me.TlbMenu.Buttons.Item(3).Enabled = crec!FRM_HABILITAR
'            If crec!FRM_VISUALIZA = 1 Then
'               Me.TlbMenu.Buttons.Item(3).Visible = crec!FRM_VISUALIZA
'            End If
'         End If
'         If crec!FRM_ACESSO = 4 Then
'            Me.TlbMenu.Buttons.Item(4).Enabled = crec!FRM_HABILITAR
'            If crec!FRM_VISUALIZA = 1 Then
'               Me.TlbMenu.Buttons.Item(4).Visible = crec!FRM_VISUALIZA
'            End If
'         End If
'
'      End If
''**********************************************************************************************
LEITURA_PROXIMO:
      cRec.MoveNext
      
Wend
'Me.TlbMenu.Buttons(1).Enabled = False

Me.MousePointer = vbDefault

Set cRec = Nothing

Exit Function

Erro:

If (Err.Number = 387 And Not cRec.EOF) Or (Err.Number = 340 And Not cRec.EOF) Then
   GoTo LEITURA_PROXIMO
End If
Set cRec = Nothing
MsgBox Err.Description
Me.MousePointer = vbDefault

End Function
Public Function CCTemp() As neExpedicao
     Set CCTemp = New neExpedicao
End Function

'Private Sub Timer1_Timer()
'VBA.DoEvents
'End Sub
