VERSION 5.00
Object = "{1C676460-C867-49DB-B514-DA5901BE8B91}#1.0#0"; "TMGNumberBox.ocx"
Begin VB.Form frmEscRelImpressora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escolha a impressora"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Impressora"
      Height          =   3375
      Left            =   90
      TabIndex        =   7
      Top             =   90
      Width           =   5625
      Begin VB.ComboBox cbo_formato 
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
         ItemData        =   "frmEscRelImpressora.frx":0000
         Left            =   3540
         List            =   "frmEscRelImpressora.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1140
         Width           =   1875
      End
      Begin VB.Frame Frame3 
         Caption         =   "Configuração fixas."
         Height          =   675
         Left            =   180
         TabIndex        =   10
         Top             =   1770
         Width           =   5265
         Begin VB.TextBox txt_impressora 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   2670
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   2505
         End
         Begin TMGNumberBox.uNumberBox txt_numLinha 
            Height          =   315
            Left            =   1290
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            MaxLength       =   2
            Text            =   ""
            BackColor       =   -2147483633
            Enabled         =   0   'False
         End
         Begin VB.Label Label3 
            Caption         =   "Impressora : "
            Height          =   225
            Left            =   1800
            TabIndex        =   13
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Nº linhas p/pag.:"
            Height          =   255
            Left            =   60
            TabIndex        =   11
            Top             =   300
            Width           =   1245
         End
      End
      Begin VB.ComboBox cbo_fonte 
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
         ItemData        =   "frmEscRelImpressora.frx":0021
         Left            =   3540
         List            =   "frmEscRelImpressora.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   750
         Width           =   945
      End
      Begin VB.CheckBox chk_generico 
         Caption         =   $"frmEscRelImpressora.frx":0064
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   4845
      End
      Begin VB.Frame frm_tipo_imp 
         Caption         =   "Tipo impresora"
         Height          =   885
         Left            =   180
         TabIndex        =   8
         Top             =   660
         Width           =   1665
         Begin VB.OptionButton Opt_tinta 
            Caption         =   "Tinta/Laser"
            Height          =   255
            Left            =   180
            TabIndex        =   2
            Top             =   540
            Width           =   1155
         End
         Begin VB.OptionButton Opt_Matricial 
            Caption         =   "Matricial"
            Height          =   255
            Left            =   180
            TabIndex        =   1
            Top             =   270
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.ComboBox Cbo_IMPRESSORA 
         Height          =   315
         ItemData        =   "frmEscRelImpressora.frx":00F3
         Left            =   180
         List            =   "frmEscRelImpressora.frx":00F5
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   5265
      End
      Begin VB.Label Label4 
         Caption         =   "Formato impressão.:"
         Height          =   225
         Left            =   2070
         TabIndex        =   15
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Tamanho da fonte.: :"
         Height          =   225
         Left            =   2070
         TabIndex        =   9
         Top             =   780
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "&Confirmar"
      Height          =   330
      Left            =   4380
      TabIndex        =   6
      Top             =   3570
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "&Desistir"
      Height          =   330
      Left            =   3060
      TabIndex        =   5
      Top             =   3570
      Width           =   1275
   End
End
Attribute VB_Name = "frmEscRelImpressora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bCancelado As Boolean 'Caso tenha cancelado a impressao.
Public sImpressora As String 'nome da impresora escolhida.
Public bTemVideo As Boolean 'caso o relatorio seja pedido pelo form previw de impressao = false

Private Sub Cbo_IMPRESSORA_Change()
sImpressora = Me.Cbo_IMPRESSORA.List(Me.Cbo_IMPRESSORA.ListIndex)
End Sub

Private Sub Cbo_IMPRESSORA_Click()
sImpressora = Me.Cbo_IMPRESSORA.List(Me.Cbo_IMPRESSORA.ListIndex)
End Sub

Private Sub Cbo_IMPRESSORA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.cmdConfirma.SetFocus
ElseIf KeyAscii = 27 Then
       Unload Me
End If
End Sub

Private Sub cmdCancela_Click()
Unload Me
End Sub

Private Sub cmdConfirma_Click()
If Me.Cbo_IMPRESSORA.ListIndex = -1 Then
   MsgBox "Escolha a saida do seu relatório.. ", , "Impressora não selecionada"
   Me.Cbo_IMPRESSORA.SetFocus
   Exit Sub
End If
If Me.Cbo_IMPRESSORA.List(Me.Cbo_IMPRESSORA.ListIndex) = "Formatação em arquivo texto (só dados)" Then
   frmGloPreview.bFormatacaoDisco = True
Else
   frmGloPreview.bFormatacaoDisco = False
End If
   
Me.bCancelado = False
Me.Hide
End Sub

Private Sub Form_Load()
Dim nx As Integer
Dim X As Printer
Dim sNome_Imp As String

If bTemVideo Then Me.Cbo_IMPRESSORA.AddItem "Listagem em Video"

For Each X In Printers
   Me.Cbo_IMPRESSORA.AddItem X.DeviceName
Next

If frmGloPreview.Lst_Exel.ListCount > 0 Then
   Me.Cbo_IMPRESSORA.AddItem "Formatação em arquivo texto (só dados)"
End If
nx = 0

For Each X In Printers
   If X.DeviceName = Me.Cbo_IMPRESSORA.List(nx) Then
      Me.Cbo_IMPRESSORA.ListIndex = nx
      Exit For
   End If
Next

Me.txt_impressora.Text = sNome_Imp

For nx = 0 To Me.Cbo_IMPRESSORA.ListCount
    If Me.txt_impressora.Text = Mid$(Me.Cbo_IMPRESSORA.List(nx), 1, Len(Trim(sNome_Imp))) Then
       Me.Cbo_IMPRESSORA.ListIndex = nx
       Exit For
    End If
Next

'If nSel_Imp = 0 Then
'   If nx = Me.Cbo_IMPRESSORA.ListCount Then
'      MsgBox "Sua impressora padrao não foi encontrada e conforme a parametrização, pede seu cadastramento, cadastre-a! --> " & sNome_Imp
'      Me.bCancelado = True
'      Me.cmdConfirma.Enabled = False
'      Me.cmdCancela.Enabled = True
'   End If
'End If

Me.txt_numLinha.Text = 55
Me.cbo_fonte.ListIndex = 1
Me.cbo_formato.ListIndex = 0
Me.bCancelado = True

'Me.chk_generico.Value = 55

'If nSel_Imp = 0 Then
'   Me.Cbo_IMPRESSORA.Enabled = False
'   Me.frm_tipo_imp.Enabled = False
'   Me.cbo_fonte.Enabled = False
'   Me.chk_generico.Enabled = False
'   Me.cbo_formato.Enabled = False
'End If

If Me.Cbo_IMPRESSORA.ListIndex < 0 Then Me.Cbo_IMPRESSORA.ListIndex = 0

End Sub

Private Sub Opt_Matricial_Click()
Me.chk_generico.Enabled = True
Me.cbo_formato.Enabled = False
End Sub

Private Sub Opt_tinta_Click()
Me.chk_generico.Enabled = False
Me.chk_generico.Value = 0
Me.cbo_formato.Enabled = True
End Sub
