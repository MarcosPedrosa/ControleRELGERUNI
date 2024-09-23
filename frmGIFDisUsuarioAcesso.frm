VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmGIFDisUsuarioAcesso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso Usuário ao Sistema"
   ClientHeight    =   5130
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3765
      Left            =   90
      TabIndex        =   5
      Top             =   5610
      Width           =   7095
      Begin GridEX16.GridEX GEXPesquisa 
         Height          =   3405
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   6006
         CursorLocation  =   3
         HideSelection   =   1
         UseEvenOddColor =   -1  'True
         MethodHoldFields=   -1  'True
         ContScroll      =   -1  'True
         SelectionStyle  =   1
         Options         =   2
         RecordsetType   =   1
         GroupByBoxInfoText=   "Arraste a coluna para agrupar"
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         HeaderFontBold  =   -1  'True
         HeaderFontWeight=   700
         ColumnHeaderHeight=   285
         IntProp2        =   0
      End
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "&Inicio"
      Height          =   330
      Left            =   1410
      TabIndex        =   4
      Top             =   5190
      Width           =   1275
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   5910
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4350
      Width           =   1275
   End
   Begin VB.TextBox txtNome 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   2220
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   90
      Width           =   3555
   End
   Begin VB.CommandButton cmd_Atualizar 
      Caption         =   "&Atualizar"
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   5190
      Width           =   1275
   End
   Begin ComctlLib.TreeView Trv_Acesso 
      Height          =   3765
      Left            =   90
      TabIndex        =   0
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6641
      _Version        =   327682
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList2"
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Usuário :"
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   150
      Width           =   615
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   6480
      Top             =   5130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGIFDisUsuarioAcesso.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGIFDisUsuarioAcesso.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGIFDisUsuarioAcesso.frx":0634
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   300
      Picture         =   "frmGIFDisUsuarioAcesso.frx":094E
      Top             =   4320
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bloqueado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   9
      Top             =   4830
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2175
      Picture         =   "frmGIFDisUsuarioAcesso.frx":0C58
      Top             =   4320
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Desabilitado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1732
      TabIndex        =   8
      Top             =   4830
      Width           =   1350
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Habilitado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3780
      TabIndex        =   7
      Top             =   4830
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4050
      Picture         =   "frmGIFDisUsuarioAcesso.frx":0F62
      Top             =   4320
      Width           =   480
   End
End
Attribute VB_Name = "frmGIFDisUsuarioAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ccodigo_pesquisa As String 'Codigo escolhido pelo usuário.
Public cnome As String 'Nome do escolhido pelo usuário.
Private i As Integer, j As Integer
Private PaiTre As Node
Private Parent As Node
Private Parent3 As Node
Private Parent5 As Node
Private Parent7 As Node
Private Parent9 As Node
Private rs As ADODB.Recordset
Private sGrupo As String

Private Sub cmd_Atualizar_Click()
Dim nRet As Integer
Dim nx As Integer
Dim nDesconto As Double
Dim nValor As String * 1
Dim nvalor1 As String * 1

On Error GoTo Erro

nRet = MsgBox("Confirma Acesso deste usuario? ", vbQuestion & vbYesNo, Me.Caption)
'Se confirmou a pagamento:
If nRet = 6 Then
   With GEXPesquisa
   .MoveFirst
   Me.MousePointer = vbHourglass
   For nx = 1 To .RowCount
       GEXPesquisa.Row = nx
       nValor = "1"
       nvalor1 = "1"
       If Me.GEXPesquisa.Value(1) <> 1 And Me.GEXPesquisa.Value(1) <> -1 Then nValor = "0"
       If Me.GEXPesquisa.Value(2) <> 1 And Me.GEXPesquisa.Value(2) <> -1 Then nvalor1 = "0"
       Call CCTempneUsuario.USUARIO_Permissao_Alterar(sNomeBanco, _
                                                      Me.GEXPesquisa.Value(12), _
                                                      nValor, _
                                                      nvalor1, _
                                                      ccodigo_pesquisa, _
                                                      Format(Me.GEXPesquisa.Value(9), "yyyy-mm-dd hh:mm:ss"), _
                                                      Me.GEXPesquisa.Value(6))
   Next nx
   End With
   cmd_atualizar.Enabled = False
   Me.MousePointer = vbDefault
   Call Carrega_januspesquisa
End If

Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub cmdfechar_Click()
    Unload Me
End Sub

Private Sub cmdSelecionar_Click()
Call Carrega_januspesquisa
Call Carrega_TreeViewpesquisa
End Sub

Private Sub Form_Activate()
   txtNome.SetFocus
End Sub

Private Sub Form_Load()
   Call Carrega_januspesquisa
   Call Carrega_TreeViewpesquisa
End Sub

Private Sub GEXPesquisa_Change()
cmd_atualizar.Enabled = True
End Sub

Private Sub GEXPesquisa_DblClick()
Dim cRec As ADODB.Recordset
Dim nx As Integer
Dim Nz As Integer
Dim sGrupo As String
Dim nGrupo As Integer

On Error GoTo Erro

If GEXPesquisa.Value(2) <> 1 Then
   MsgBox "Marque o item para abrir os itens deste menu."
   Exit Sub
End If

'If GEXPesquisa.Value(3) <> "MnuMain" Then
'   Exit Sub
'End If
Rem saber a quantidade de grupos existentes
nGrupo = 0
Nz = 1

For nx = 1 To 4
    If Val(Mid$(GEXPesquisa.Value(6), Nz, 2)) > 0 Then nGrupo = nGrupo + 1
    Nz = Nz + 2
Next

sGrupo = GEXPesquisa.Value(6)

Set rs = New ADODB.Recordset

Me.MousePointer = vbHourglass

Set rs = CCTempneUsuario.USUARIO_Consultar_Acesso(ccodigo_pesquisa)

Set cRec = New ADODB.Recordset

cRec.Fields.Append "FRM_VISUALIZA", adBSTR '1
cRec.Fields.Append "FRM_HABILITAR", adBSTR '2
cRec.Fields.Append "FOR_NOME_EDITOR", adBSTR '3
cRec.Fields.Append "FOR_NOME_FORM", adBSTR '4
cRec.Fields.Append "FOR_DESCRICAO", adBSTR '5
cRec.Fields.Append "FOR_GRUPO", adBSTR '6
cRec.Fields.Append "FRM_BOTAO", adBSTR '7
cRec.Fields.Append "FRM_ACESSO", adBSTR '8
cRec.Fields.Append "FRM_DTA", adBSTR '9
cRec.Fields.Append "FRM_DTI", adBSTR '10
cRec.Fields.Append "FRM_USU", adBSTR '11
cRec.Fields.Append "FRM_FOR_CODIGO", adBSTR '12
cRec.Open

While Not rs.EOF
      Nz = 0
      If Mid$(rs!FOR_GRUPO, 1, 2) = Mid$(sGrupo, 1, 2) Then
         
         If nGrupo = 0 And _
            Val(Mid$(rs!FOR_GRUPO, 1, 2)) = 0 And Val(Mid$(rs!FOR_GRUPO, 3, 10)) > 0 Then
            Nz = 1
         End If
         
         If nGrupo = 1 And _
            Val(Mid$(rs!FOR_GRUPO, 3, 2)) > 0 And Val(Mid$(rs!FOR_GRUPO, 5, 10)) = 0 Then
            Nz = 1
         End If
         
         If nGrupo = 2 And _
            Mid$(rs!FOR_GRUPO, 3, 2) = Mid$(sGrupo, 3, 2) And _
            Val(Mid$(rs!FOR_GRUPO, 5, 2)) > 0 And Val(Mid$(rs!FOR_GRUPO, 7, 10)) = 0 Then
            Nz = 1
         End If
         
         If nGrupo = 3 And _
            Mid$(rs!FOR_GRUPO, 5, 2) = Mid$(sGrupo, 5, 2) And _
            Val(Mid$(rs!FOR_GRUPO, 7, 2)) > 0 And Val(Mid$(rs!FOR_GRUPO, 9, 10)) = 0 Then
            Nz = 1
         End If
         
         If nGrupo = 4 And _
            Mid$(rs!FOR_GRUPO, 7, 2) = Mid$(sGrupo, 7, 2) And _
            Val(Mid$(rs!FOR_GRUPO, 9, 2)) > 0 Then
            Nz = 1
         End If
            
         If Nz = 1 Then
            cRec.AddNew
            cRec.Fields.Item("FRM_VISUALIZA").Value = rs!FRM_VISUALIZA
            cRec.Fields.Item("FRM_HABILITAR").Value = rs!FRM_HABILITAR
            cRec.Fields.Item("FOR_NOME_EDITOR").Value = IIf(IsNull(rs!FOR_NOME_EDITOR), " ", rs!FOR_NOME_EDITOR)
            cRec.Fields.Item("FOR_NOME_FORM").Value = IIf(IsNull(rs!FOR_NOME_FORM), " ", rs!FOR_NOME_FORM)
            cRec.Fields.Item("FOR_DESCRICAO").Value = rs!FOR_DESCRICAO
            cRec.Fields.Item("FOR_GRUPO").Value = rs!FOR_GRUPO
            cRec.Fields.Item("FRM_BOTAO").Value = rs!FRM_BOTAO
            cRec.Fields.Item("FRM_ACESSO").Value = rs!FRM_ACESSO
            cRec.Fields.Item("FRM_DTA").Value = rs!FRM_DTA
            cRec.Fields.Item("FRM_DTI").Value = rs!FRM_DTI
            cRec.Fields.Item("FRM_USU").Value = IIf(IsNull(rs!FRM_USU), 0, rs!FRM_USU)
            cRec.Fields.Item("FRM_FOR_CODIGO").Value = rs!FRM_FOR_CODIGO
            cRec.Update
         End If
      
      End If
      rs.MoveNext

Wend

If cRec.RecordCount = 0 Then
   MsgBox "Não exite sub-grupos para este iten!"
Else
   Set GEXPesquisa.ADORecordset = cRec

   With GEXPesquisa
        .Columns(1).Caption = "Ver"
        .Columns(1).Width = TextWidth("wwww")
        .Columns(1).ColumnType = 3
        .Columns(2).Caption = "Hab"
        .Columns(2).Width = TextWidth("wwww")
        .Columns(2).ColumnType = 3
        .Columns(3).Caption = "Menu"
        .Columns(3).Width = TextWidth("wwwwwwwwwwwww0")
        .Columns(3).EditType = jgexEditNone
        .Columns(4).Visible = False
        .Columns(5).Caption = "Função"
        .Columns(5).Width = TextWidth("wwwwwwwwwwWWWWWwwwwwwwwww0")
        .Columns(5).EditType = jgexEditNone
        For nx = 6 To .Columns.Count
            .Columns(nx).Visible = False
        Next nx
   End With
End If

Me.MousePointer = vbDefault

Me.GEXPesquisa.BackColorRowGroup = &H808000
Me.GEXPesquisa.IsGroupItem (2)

Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub GEXPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Me.GEXPesquisa.MovePrevious
   GEXPesquisa_DblClick
End If
If KeyCode = 27 Then cmdfechar_Click
End Sub
Function Carrega_januspesquisa()

Dim cRec As ADODB.Recordset
Dim nx As Integer

On Error GoTo Erro

Set rs = New ADODB.Recordset

Me.MousePointer = vbHourglass

Set rs = CCTempneUsuario.USUARIO_Consultar_Acesso(sNomeBanco, ccodigo_pesquisa)

Set cRec = New ADODB.Recordset

cRec.Fields.Append "FRM_VISUALIZA", adBSTR  '1
cRec.Fields.Append "FRM_HABILITAR", adBSTR '2
cRec.Fields.Append "FOR_NOME_EDITOR", adBSTR '3
cRec.Fields.Append "FOR_NOME_FORM", adBSTR '4
cRec.Fields.Append "FOR_DESCRICAO", adBSTR '5
cRec.Fields.Append "FOR_GRUPO", adBSTR '6
cRec.Fields.Append "FRM_BOTAO", adBSTR '7
cRec.Fields.Append "FRM_ACESSO", adBSTR '8
cRec.Fields.Append "FRM_DTA", adBSTR '9
cRec.Fields.Append "FRM_DTI", adBSTR '10
cRec.Fields.Append "FRM_USU", adBSTR '11
cRec.Fields.Append "FRM_FOR_CODIGO", adBSTR '12
cRec.Open

While Not rs.EOF
      If rs!FOR_NOME_EDITOR = "MnuMain" Then
         cRec.AddNew
         cRec.Fields.Item("FRM_VISUALIZA").Value = rs!FRM_VISUALIZA
         cRec.Fields.Item("FRM_HABILITAR").Value = rs!FRM_HABILITAR
         cRec.Fields.Item("FOR_NOME_EDITOR").Value = IIf(IsNull(rs!FOR_NOME_EDITOR), " ", rs!FOR_NOME_EDITOR)
         cRec.Fields.Item("FOR_NOME_FORM").Value = IIf(IsNull(rs!FOR_NOME_FORM), " ", rs!FOR_NOME_FORM)
         cRec.Fields.Item("FOR_DESCRICAO").Value = rs!FOR_DESCRICAO
         cRec.Fields.Item("FOR_GRUPO").Value = rs!FOR_GRUPO
         cRec.Fields.Item("FRM_BOTAO").Value = rs!FRM_BOTAO
         cRec.Fields.Item("FRM_ACESSO").Value = rs!FRM_ACESSO
         cRec.Fields.Item("FRM_DTA").Value = rs!FRM_DTA
         cRec.Fields.Item("FRM_DTI").Value = rs!FRM_DTI
         cRec.Fields.Item("FRM_USU").Value = IIf(IsNull(rs!FRM_USU), 0, rs!FRM_USU)
         cRec.Fields.Item("FRM_FOR_CODIGO").Value = rs!FRM_FOR_CODIGO
         cRec.Update
      End If
      rs.MoveNext

Wend

Set GEXPesquisa.ADORecordset = cRec

With GEXPesquisa
     .Columns(1).Caption = "Ver"
     .Columns(1).Width = TextWidth("wwww")
     .Columns(1).ColumnType = 3
     .Columns(2).Caption = "Hab"
     .Columns(2).Width = TextWidth("wwww")
     .Columns(2).ColumnType = 3
     .Columns(3).Caption = "Menu"
     .Columns(3).Width = TextWidth("wwwwwwwwwwwww0")
     .Columns(3).EditType = jgexEditNone
     .Columns(4).Visible = False
     .Columns(5).Caption = "Função"
     .Columns(5).Width = TextWidth("wwwwwwwwwwwwwwWWWWWWWwwwwww0")
     .Columns(5).EditType = jgexEditNone
     For nx = 6 To .Columns.Count
         .Columns(nx).Visible = False
     Next nx
End With
Me.MousePointer = vbDefault

Me.GEXPesquisa.BackColorRowGroup = &H808000
Me.GEXPesquisa.IsGroupItem (2)

Rem Set rs = Nothing
Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
End Function


Private Sub GEXPesquisa_ColumnHeaderClick(ByVal Column As GridEX16.JSColumn)
If Me.GEXPesquisa.SortKeys.Count > 0 Then
    If Me.GEXPesquisa.SortKeys(1).ColIndex = Column.Index Then
        If Me.GEXPesquisa.SortKeys(1).SortOrder = jgexSortAscending Then
            Me.GEXPesquisa.SortKeys(1).SortOrder = jgexSortDescending
        Else
            Me.GEXPesquisa.SortKeys(1).SortOrder = jgexSortAscending
        End If
        Exit Sub
    End If
    Me.GEXPesquisa.SortKeys.Clear
End If
Me.GEXPesquisa.SortKeys.Add Column.Index, jgexSortAscending

End Sub


Private Sub Carrega_TreeViewpesquisa()

On Error GoTo Erro
Dim nTipo As Integer
'Set rs = New ADODB.Recordset

Me.MousePointer = vbHourglass

'Set rs = CCTempneFormulario.TipoForm_Consultar()
'           "EXP_REL_USUARIO_FORMULARIO.FRM_VISUALIZA,  " & _
'           "EXP_REL_USUARIO_FORMULARIO.FRM_HABILITAR,  " & _
'           "ESC_TAB_FORMULARIO.FOR_NOME_EDITOR,  " & _
'           "ESC_TAB_FORMULARIO.FOR_NOME_FORM,  " & _
'           "ESC_TAB_FORMULARIO.FOR_DESCRICAO,  " & _
'           "ESC_TAB_FORMULARIO.FOR_GRUPO,  " & _
'           "EXP_REL_USUARIO_FORMULARIO.FRM_BOTAO,  " & _
'           "EXP_REL_USUARIO_FORMULARIO.FRM_ACESSO,  " & _
'           "EXP_REL_USUARIO_FORMULARIO.FRM_DTA,  " & _
'           "EXP_REL_USUARIO_FORMULARIO.FRM_DTI,  " & _
'           "EXP_REL_USUARIO_FORMULARIO.FRM_USU, " & _
'           "EXP_REL_USUARIO_FORMULARIO.FRM_FOR_CODIGO, "

If Not rs.BOF Then
   Me.Trv_Acesso.Nodes.Clear
   rs.MoveFirst
   sGrupo = Mid$(rs!FOR_GRUPO, 1, 2)
   Set PaiTre = Trv_Acesso.Nodes.Add(, , "A0000000000", "Menu do sistema")
   PaiTre.Expanded = True
'   Set Parent = Trv_Acesso.Nodes.Add(, , "j" + rs!FOR_GRUPO, Mid$(rs!FOR_GRUPO, 1, 2) & "-" & rs!FOR_NOME_EDITOR, IIf(Mid$(rs!FOR_GRUPO, 3, 2) = "00", 1, 2), IIf(Mid$(rs!FOR_GRUPO, 3, 2), 1, 2))
   While Not rs.EOF
        i = 0: j = 0
        
        sGrupo = Mid$(rs!FOR_GRUPO, 1, 2)
        
        If Mid$(rs!FOR_GRUPO, 9, 2) <> "00" Then
           While Not rs.EOF
             Rem tipo de visualizacao do usuario
             If rs!FRM_VISUALIZA = 0 Then
                nTipo = 1 'o usuario esta desabilitado
             ElseIf rs!FRM_HABILITAR = 0 Then
                nTipo = 2 'o usuario apenas visualiza
             Else
                nTipo = 3 'o usuario esta habilitado
             End If
             Set Parent9 = Trv_Acesso.Nodes.Add(Parent7, tvwChild, _
                                             Format(rs!FRM_FOR_CODIGO, "000") & "K" + rs!FOR_GRUPO, _
                                             "[" & Format(rs!FRM_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]", nTipo, nTipo)
             rs.MoveNext
           Wend
        End If
        
        If Not rs.EOF Then
           If Mid$(rs!FOR_GRUPO, 7, 2) <> "00" Then
              While Not rs.EOF
                Rem tipo de visualizacao do usuario
                If rs!FRM_VISUALIZA = 0 Then
                   nTipo = 1 'o usuario esta desabilitado
                ElseIf rs!FRM_HABILITAR = 0 Then
                   nTipo = 2 'o usuario apenas visualiza
                Else
                   nTipo = 3 'o usuario esta habilitado
                End If
                Set Parent7 = Trv_Acesso.Nodes.Add(Parent5, tvwChild, _
                                                  Format(rs!FRM_FOR_CODIGO, "000") & "K" + rs!FOR_GRUPO, _
                                                  "[" & Format(rs!FRM_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]", nTipo, nTipo)
                rs.MoveNext
             Wend
           End If
        End If
        
        If Not rs.EOF Then
           If Mid$(rs!FOR_GRUPO, 5, 2) <> "00" Then
              While Not rs.EOF
                Rem tipo de visualizacao do usuario
                If rs!FRM_VISUALIZA = 0 Then
                   nTipo = 1 'o usuario esta desabilitado
                ElseIf rs!FRM_HABILITAR = 0 Then
                   nTipo = 2 'o usuario apenas visualiza
                Else
                   nTipo = 3 'o usuario esta habilitado
                End If
                Set Parent5 = Trv_Acesso.Nodes.Add(Parent3, tvwChild, _
                                                  Format(rs!FRM_FOR_CODIGO, "000") & "K" + rs!FOR_GRUPO, _
                                                  "[" & Format(rs!FRM_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]", nTipo, nTipo)
                rs.MoveNext
              Wend
           End If
        End If
        
        If Not rs.EOF Then
           If Mid$(rs!FOR_GRUPO, 3, 2) <> "00" Then
              While Not rs.EOF And Mid$(rs!FOR_GRUPO, 3, 2) <> "00"
                Rem tipo de visualizacao do usuario
                If rs!FRM_VISUALIZA = 0 Then
                   nTipo = 1 'o usuario esta desabilitado
                ElseIf rs!FRM_HABILITAR = 0 Then
                   nTipo = 2 'o usuario apenas visualiza
                Else
                   nTipo = 3 'o usuario esta habilitado
                End If
                Set Parent3 = Trv_Acesso.Nodes.Add(Parent, tvwChild, _
                                                  Format(rs!FRM_FOR_CODIGO, "000") & "K" + rs!FOR_GRUPO, _
                                                  "[" & Format(rs!FRM_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]", nTipo, nTipo)
                rs.MoveNext
                If rs.EOF Then GoTo SAIDA_LOOP
                If Mid$(rs!FOR_GRUPO, 5, 2) <> "00" Then
                   While Not rs.EOF And Mid$(rs!FOR_GRUPO, 5, 2) <> "00"
                     Rem tipo de visualizacao do usuario
                     If rs!FRM_VISUALIZA = 0 Then
                        nTipo = 1 'o usuario esta desabilitado
                     ElseIf rs!FRM_HABILITAR = 0 Then
                        nTipo = 2 'o usuario apenas visualiza
                     Else
                        nTipo = 3 'o usuario esta habilitado
                     End If
                     Set Parent5 = Trv_Acesso.Nodes.Add(Parent3, tvwChild, _
                                                       Format(rs!FRM_FOR_CODIGO, "000") & "K" + rs!FOR_GRUPO, _
                                                       "[" & Format(rs!FRM_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]", nTipo, nTipo)
                     rs.MoveNext
                     If rs.EOF Then GoTo SAIDA_LOOP
                     If Mid$(rs!FOR_GRUPO, 7, 2) <> "00" Then
                        Rem tipo de visualizacao do usuario
                        If rs!FRM_VISUALIZA = 0 Then
                           nTipo = 1 'o usuario esta desabilitado
                        ElseIf rs!FRM_HABILITAR = 0 Then
                           nTipo = 2 'o usuario apenas visualiza
                        Else
                           nTipo = 3 'o usuario esta habilitado
                        End If
                        While Not rs.EOF And Mid$(rs!FOR_GRUPO, 7, 2) <> "00"
                          Set Parent7 = Trv_Acesso.Nodes.Add(Parent5, tvwChild, _
                                                            Format(rs!FRM_FOR_CODIGO, "000") & "K" + rs!FOR_GRUPO, _
                                                            "[" & Format(rs!FRM_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]", nTipo, nTipo)
                          rs.MoveNext
                          If rs.EOF Then GoTo SAIDA_LOOP
                          If Mid$(rs!FOR_GRUPO, 9, 2) <> "00" Then
                             While Not rs.EOF And Mid$(rs!FOR_GRUPO, 9, 2) <> "00"
                               Set Parent9 = Trv_Acesso.Nodes.Add(Parent7, tvwChild, _
                                                                 Format(rs!FRM_FOR_CODIGO, "000") & "K" + rs!FOR_GRUPO, _
                                                                 "[" & Format(rs!FRM_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]", nTipo, nTipo)
                               rs.MoveNext
                             Wend
                          End If
                       Wend
                     End If
                   Wend
                End If
              Wend
           End If
        End If
SAIDA_LOOP:
        If Not rs.EOF Then
           Rem tipo de visualizacao do usuario
           If rs!FRM_VISUALIZA = 0 Then
              nTipo = 1 'o usuario esta desabilitado
           ElseIf rs!FRM_HABILITAR = 0 Then
              nTipo = 2 'o usuario apenas visualiza
           Else
              nTipo = 3 'o usuario esta habilitado
           End If
           Set Parent = Trv_Acesso.Nodes.Add(PaiTre, tvwChild, Format(rs!FRM_FOR_CODIGO, "000") & "J" + rs!FOR_GRUPO, Mid$(rs!FOR_GRUPO, 1, 2) & "-" & rs!FOR_DESCRICAO, nTipo, nTipo)
'           Parent.Expanded = True
           rs.MoveNext
        End If
    Wend
End If



Me.MousePointer = vbDefault
'Set rs = Nothing

Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub Trv_Acesso_DblClick()
Dim nValor As Integer
Dim nvalor1 As Integer

If Me.Trv_Acesso.SelectedItem.SelectedImage = 1 Then 'de desabilitado para apenas visualizar
   Me.Trv_Acesso.SelectedItem.SelectedImage = 2: Me.Trv_Acesso.SelectedItem.Image = 2
   nValor = 1: nvalor1 = 0
ElseIf Me.Trv_Acesso.SelectedItem.SelectedImage = 2 Then 'de visualizar para habilitar
   Me.Trv_Acesso.SelectedItem.SelectedImage = 3: Me.Trv_Acesso.SelectedItem.Image = 3
   nValor = 1: nvalor1 = 1
ElseIf Me.Trv_Acesso.SelectedItem.SelectedImage = 3 Then 'del habilitado para desabilitado
   Me.Trv_Acesso.SelectedItem.SelectedImage = 1: Me.Trv_Acesso.SelectedItem.Image = 1
   nValor = 0: nvalor1 = 0
End If

Me.Trv_Acesso.Refresh

If Mid$(Me.Trv_Acesso.SelectedItem.Key, 1, 3) <> "A00" Then
   Call CCTempneUsuario.USUARIO_Permissao_Alterar(sNomeBanco, Mid$(Me.Trv_Acesso.SelectedItem.Key, 1, 3), _
                                                  nValor, _
                                                  nvalor1, _
                                                  ccodigo_pesquisa, _
                                                  "", _
                                                  Mid$(Me.Trv_Acesso.SelectedItem.Key, 5, 10))
End If


End Sub


