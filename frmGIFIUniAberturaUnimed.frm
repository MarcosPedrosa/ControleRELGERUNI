VERSION 5.00
Begin VB.Form frmGIFIUniAberturaUnimed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abertura Mov. Mensal"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7710
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2250
      Width           =   1275
   End
   Begin VB.CommandButton cmdSelecionar 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Confirmar"
      Height          =   330
      Left            =   5010
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2250
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione empresa e atualize os dados abaixo"
      Height          =   2025
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CheckBox chk_situacao 
         Caption         =   "Movimento Fechado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         TabIndex        =   11
         Top             =   1590
         Width           =   2115
      End
      Begin VB.TextBox TXT_VERBA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6570
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "6,00"
         Top             =   1050
         Width           =   795
      End
      Begin VB.ComboBox cbo_coligada 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   7215
      End
      Begin VB.TextBox TXT_DESCONTO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4590
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "6,00"
         Top             =   1050
         Width           =   765
      End
      Begin VB.TextBox TXT_ANO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1830
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "2010"
         Top             =   1050
         Width           =   765
      End
      Begin VB.TextBox TXT_MES 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "01"
         Top             =   1050
         Width           =   435
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   7530
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "VERBA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5520
         TabIndex        =   10
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DESCONTO%:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2790
         TabIndex        =   9
         Top             =   1080
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "MÊS/ANO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   90
         TabIndex        =   8
         Top             =   1080
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmGIFIUniAberturaUnimed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Private cRec As ADODB.Recordset

Private Sub cbo_coligada_Change()
Call carregar_Tabela_coligada
End Sub

Private Sub cbo_coligada_Click()
Call carregar_Tabela_coligada
End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub cmdSelecionar_Click()
Dim RESPOSTA As Integer

On Error GoTo Erro

RESPOSTA = MsgBox("'Confirma atualizar dados da empresa? " & "Digite a resposta", 20, "Sim/Não?")

If RESPOSTA = 6 Then
   Call CCTempneUniColigada.Coligada_Alterar(sBancoUnimed, _
                                             Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), _
                                             "", _
                                             Me.TXT_DESCONTO.Text, _
                                             "", _
                                             Me.TXT_VERBA.Text, _
                                             Me.TXT_ANO.Text & Me.TXT_MES.Text, _
                                             Me.chk_situacao.Value)
   MsgBox "Alteração realizada com Sucesso!"
    
End If

Exit Sub

Erro:

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
If Flag_ativo = True Then
   Exit Sub
End If
Me.Top = 0
Me.Left = 0
Flag_ativo = True
Call carregar_coligada
Call carregar_Tabela_coligada

End Sub

Private Sub Form_Load()
Dim nx As Integer

Me.Top = 0
Me.Left = 0

End Sub


Private Sub carregar_coligada()
Dim nx As Integer
Dim cRec As ADODB.Recordset

On Error GoTo Erro

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass
Set cRec = rRec_cliente
Set cRec = CCTempneUniColigada.Coligada_Consultar(sBancoUnimed)

Me.cbo_coligada.Clear

nx = 0

If cRec.RecordCount > 0 Then
   cRec.MoveFirst
   While Not cRec.EOF
       If Not IsNull(cRec!TCO_CODIGO) Then
          Me.cbo_coligada.AddItem cRec!TCO_CODIGO & " - " & Trim(cRec!TCO_DESCRICAO)
          Me.cbo_coligada.ItemData(nx) = cRec!TCO_CODIGO
          nx = nx + 1
       End If
       cRec.MoveNext
   Wend
   If nx < 2 Then
      Me.cbo_coligada.ListIndex = 0
      Call carregar_Tabela_coligada
   End If
Else
   MsgBox "Não existem empresas Coligadas, procure o responsável."
End If

Me.MousePointer = vbDefault

Set cRec = Nothing
Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub carregar_Tabela_coligada()
Dim nx As Integer
Dim cRec As ADODB.Recordset

On Error GoTo Erro

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass
Set cRec = rRec_cliente
Set cRec = CCTempneUniColigada.Coligada_Consultar(sBancoUnimed, Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex))

If cRec.RecordCount > 0 Then
   cRec.MoveFirst
   Me.TXT_MES.Text = Mid$(cRec!TCO_ANO_MES_PROC, 5, 2)
   Me.TXT_ANO.Text = Mid$(cRec!TCO_ANO_MES_PROC, 1, 4)
   Me.TXT_DESCONTO.Text = cRec!TCO_DESCONTO
   Me.TXT_VERBA.Text = cRec!TCO_VERBA
   Me.chk_situacao.Value = cRec!TCO_MOV_ABERTO
Else
   MsgBox "Não existem empresas Coligadas, procure o responsável."
End If

Me.MousePointer = vbDefault

Set cRec = Nothing
Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub


