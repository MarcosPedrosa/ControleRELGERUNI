VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmGloPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preview - Relatório "
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox Lst_Exel 
      Height          =   255
      Left            =   8580
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Procure o arquivo para impressão na tela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   2820
      TabIndex        =   16
      Top             =   870
      Visible         =   0   'False
      Width           =   6675
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   90
         TabIndex        =   21
         Top             =   570
         Width           =   6525
      End
      Begin VB.CommandButton cmdfechar 
         Caption         =   "&Fechar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5250
         TabIndex        =   20
         Top             =   4170
         Width           =   1275
      End
      Begin VB.CommandButton cmd_confirmar 
         Caption         =   "&Confirma"
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
         Height          =   330
         Left            =   3930
         TabIndex        =   19
         Top             =   4170
         Width           =   1275
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   60
         Pattern         =   "*.txt"
         TabIndex        =   18
         Top             =   2070
         Width           =   6495
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   210
         Width           =   6525
      End
   End
   Begin VB.ListBox ListPreVw 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   60
      TabIndex        =   22
      Top             =   600
      Width           =   11145
   End
   Begin VB.CommandButton cmdPrevW 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   6
      Left            =   7680
      Picture         =   "frmGloPreview.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Grava o Relatório em um Arquivo"
      Top             =   0
      Width           =   585
   End
   Begin ComctlLib.ProgressBar PBar1 
      Height          =   315
      Left            =   750
      TabIndex        =   12
      Top             =   5400
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdPrevW 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   5
      Left            =   10680
      Picture         =   "frmGloPreview.frx":044A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Proximo registro de procura."
      Top             =   5820
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   30
      MaxLength       =   25
      TabIndex        =   7
      Top             =   5790
      Width           =   10005
   End
   Begin VB.CommandButton cmdPrevW 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   30
      Picture         =   "frmGloPreview.frx":088C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Diminui o tamanho da Fonte"
      Top             =   0
      Width           =   585
   End
   Begin VB.CheckBox ChkAcompanha 
      Caption         =   "&Acompanhar Impressão na Tela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1410
      TabIndex        =   3
      Top             =   330
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.CheckBox ChkParte 
      Caption         =   "Impressão Parcial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1410
      TabIndex        =   2
      Top             =   30
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrevW 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   9
      Left            =   10620
      Picture         =   "frmGloPreview.frx":0CCE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   30
      Width           =   585
   End
   Begin VB.CommandButton cmdPrevW 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   4
      Left            =   630
      Picture         =   "frmGloPreview.frx":0FD8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aumenta o tamanho da Fonte"
      Top             =   0
      Width           =   585
   End
   Begin VB.CommandButton cmdPrevW 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   10080
      Picture         =   "frmGloPreview.frx":141A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Localiza uma palavra ou texto (60 posiçoes) no Relatório"
      Top             =   5820
      Width           =   555
   End
   Begin VB.CommandButton cmdPrevW 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   4800
      Picture         =   "frmGloPreview.frx":1724
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Grava o Relatório em um Arquivo"
      Top             =   0
      Width           =   585
   End
   Begin VB.CommandButton cmdPrevW 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   10020
      Picture         =   "frmGloPreview.frx":1A2E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Impressão para IMPRESSORA"
      Top             =   30
      Width           =   585
   End
   Begin VB.Label LBL_PROCESSADO 
      Height          =   255
      Left            =   60
      TabIndex        =   24
      Top             =   5430
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOME DO RELATÓRIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   14
      Top             =   0
      Width           =   2280
   End
   Begin VB.Label LBL_NOMEREL 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RELQUALQUERR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5400
      TabIndex        =   13
      Top             =   210
      Width           =   2280
   End
   Begin VB.Label LBLLinha 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LBLLinha 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmGloPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Dim Ki As Long  ' Impressao:  Linha Inicial
Dim Kf As Long  ' Impressao:  Linha Final
Dim ZIni As Double
Dim Zi As Long
Dim cNomeArq As String
Public bFormatacaoDisco As Boolean


Private Sub ChkParte_Click()

If ChkParte.Value = 1 Then
   LBLLinha(0).Visible = True
   LBLLinha(1).Visible = True
   MsgBox "Marque com um DUPLO CLICK a linha INICIAL e depois com outro CLICK a FINAL", 16, "Mensagem 0000: Impressão Parcial.  Escolha o Trecho INICIAL e FINAL"
   ListPreVw.SetFocus
Else
   LBLLinha(0) = "": LBLLinha(1) = ""
   LBLLinha(0).Visible = False
   LBLLinha(1).Visible = False
End If
End Sub
Private Sub cmd_confirmar_Click()
Dim sNome_Arq As String
Dim RESPOSTA As Integer
Dim X1 As Double
Dim Nada As String
Dim X As Variant

On Error GoTo Erro

sNome_Arq = Me.Dir1.Path & "\" & Me.File1.FileName

RESPOSTA = MsgBox("'Carregando relatório de nome : " & sNome_Arq, 20, "Sim/Não?")

If RESPOSTA = 6 Then
    
    '-------------------------------------> Localizando o Banco de Dados

    Nada = Dir(sNome_Arq, vbNormal)
    If Nada <> Me.File1.FileName Then
       MsgBox "Arquivo não encontrado em " + sNome_Arq + "COPIA NÃO pode ser feita"
       Exit Sub
    End If

   Open sNome_Arq For Random Access Read Write As #21 Len = Len(Arq_Impressao)
   Ki = 0: Kf = ListPreVw.ListCount - 1
'   If ChkParte.Value = 1 Then
'      If RTrim(LBLLinha(0)) <> "" Then Ki = Val(LBLLinha(0))
'      If RTrim(LBLLinha(1)) <> "" Then Kf = Val(LBLLinha(1))
'      If Kf < Ki Then
'         MsgBox "Linha Inicial =" + Str(Ki) + "   MAIOR  que a Linha Final =" + Str(Kf), 16, "Impressão PARCIAL com Intervalo ERRADO"
'         Exit Sub
'      End If
'   End If
   If ChkAcompanha.Value = 0 Then
      ListPreVw.Visible = False
   End If
   Me.ListPreVw.Clear
   Me.PBar1.Visible = True
   Me.PBar1.Min = 0
   Me.PBar1.Max = Int(LOF(21) / Len(Arq_Impressao))
   Ki = 1
   Kf = Int(LOF(21) / Len(Arq_Impressao))
   X1 = 0
   For Zi = Ki To Kf
       X1 = X1 + 1
       Get 21, X1, Arq_Impressao
       Me.ListPreVw.AddItem Arq_Impressao.FCampo136
       Me.PBar1.Value = Zi
   Next
   Close #21
   ListPreVw.Visible = True
   Me.PBar1.Visible = False
   ListPreVw.ListIndex = 0
   Me.PBar1.Min = 0
   Me.Frame1.Visible = False
End If
Exit Sub

Erro:

If Err.Number = 52 Then
   MsgBox "Arquivo não encontrado em " + sNome_Arq + "Lançamento NÃO pode ser feito"
Else
   MsgBox "Erro não localizado, anote o numero e chame o responsável. Numero = " & Err.Number
   Close #21
   ListPreVw.Clear
   ListPreVw.Visible = True
   Me.PBar1.Visible = False
   Me.Frame1.Visible = False
End If

Me.MousePointer = vbDefault

End Sub

Private Sub cmdFechar_Click()
Me.Frame1.Visible = False
End Sub

Private Sub cmdPrevW_Click(Index As Integer)
Dim Ki As Double
Dim Kf As Double
Dim W2S As String
Dim W1S As String
Dim K As Integer
Dim z As Double
Dim X1 As Double
Dim RESPOSTA As Integer
Dim sNome_Arq As String
Dim Nada As String

On Error GoTo Erro

'***********index=0 = impresao ***********************************************************
        If Index = 0 Then
           Dim oTela As frmEscRelImpressora
           Dim X As Printer
           Dim slinha As String * 136
           
           If Me.ListPreVw.ListCount = 0 Then Exit Sub
           
           Set oTela = New frmEscRelImpressora
           oTela.bTemVideo = False
           oTela.Show 1
           If oTela.bCancelado Then
               Unload oTela: Set oTela = Nothing
               Me.MousePointer = vbDefault
               Exit Sub
           
           Else
               
              If bFormatacaoDisco = False Then
                 If oTela.Cbo_IMPRESSORA.ListIndex > -1 Then
                 
                    For Each X In Printers
                       If X.DeviceName = oTela.sImpressora Then
                          Set Printer = X
                          Exit For
                       End If
                    Next
                    
                    If oTela.Opt_tinta.Value = True Then
                        Printer.Font = "Courier New"
                        Printer.FontBold = False
                        Printer.Font.Size = Val(oTela.cbo_fonte.List(oTela.cbo_fonte.ListIndex))
                        Printer.Orientation = oTela.cbo_formato.ItemData(oTela.cbo_formato.ListIndex)
                    End If
                     
                  End If
              
              Else
                  
                  cNomeArq = IIf(Len(Trim(cNomeArq)) = 0, "c:\", cNomeArq)
                  sNome_Arq = cNomeArq
                  sNome_Arq = cNomeArq & Me.LBL_NOMEREL.Caption & "_" & Format(Date, "DDMMYYYY") & "_" & Format(Time(), "HHMMSS") & ".txt"
'                  If Dir(cNomeArq) = 1 Then
                     Open sNome_Arq For Random Access Read Write As #21 Len = Len(Arq_ImpressaoE)
'                  Else
'                  End If
              
              End If
           End If
           
           Unload oTela: Set oTela = Nothing
           Ki = 0: Kf = ListPreVw.ListCount - 1
           
           If ChkParte.Value = 1 Then
              If RTrim(LBLLinha(0)) <> "" Then Ki = Val(LBLLinha(0))
              If RTrim(LBLLinha(1)) <> "" Then Kf = Val(LBLLinha(1))
              If Kf < Ki Then
                 MsgBox "Linha Inicial =" + Str(Ki) + "   MAIOR  que a Linha Final =" + Str(Kf), 16, "Impressão PARCIAL com Intervalo ERRADO"
                 Exit Sub
              End If
           End If
'           If ChkAcompanha.Value = 0 Then
              ListPreVw.Visible = False
'           End If
           z = 0
           Me.PBar1.Visible = True
           Me.PBar1.Min = 0
           Me.PBar1.Max = Kf - Ki
           
           For Zi = Ki To Kf
               
               If bFormatacaoDisco = True Then
                  If Zi < Me.Lst_Exel.ListCount Then
                     If bFormatacaoDisco Then
                        Me.Lst_Exel.ListIndex = Zi
                        Arq_ImpressaoE.FCampo1000 = Me.Lst_Exel.List(Zi)
                     
                        Arq_ImpressaoE.FFinal = Chr$(13) + Chr$(10)
                        Put 21, z + 1, Arq_ImpressaoE
                     Else
                        Me.Lst_Exel.ListIndex = Zi
                        Arq_Impressao.FCampo136 = Me.Lst_Exel.List(Zi)
                     
                        Arq_Impressao.FFinal = Chr$(13) + Chr$(10)
                        Put 21, z + 1, Arq_Impressao
                     End If
                  End If
                  slinha = " "
               Else
                  ListPreVw.ListIndex = Zi
                  If ChkAcompanha.Value = 1 Then
                     ListPreVw.Refresh
                  End If
               
                  slinha = ListPreVw
                  If Mid$(slinha, 1, 34) = "========= <<<<<<<>>>>>>> =========" Then
                     slinha = ""
                     Printer.NewPage
                  Else
                     slinha = ListPreVw
                     Printer.Print slinha
                     
                  End If
               End If
               
               slinha = ""
               Me.PBar1.Value = z
               z = z + 1
           Next
           
           If bFormatacaoDisco = True Then
              Close #21
           Else
              Printer.EndDoc
           End If
           ListPreVw.Visible = True
           Me.PBar1.Visible = False
           ListPreVw.ListIndex = 0
           Me.PBar1.Min = 0
           
        End If
        ChkParte.Value = 0
'***********index=0 = impresao ***********************************************************
        
'***********index=1 = impressao para texto ***********************************************
        If Index = 1 Then
        
           If Me.ListPreVw.ListCount = 0 Then Exit Sub
           
           
           Dim Message, Title, Default, MyValue
           Message = "Gerando relatório de nome"   ' Set prompt.
           Title = "Confirmar onde será gerado o relatorio em Disco"   ' Set title.
           sNome_Arq = cNomeArq
           sNome_Arq = cNomeArq & Me.LBL_NOMEREL.Caption & "_" & Format(Date, "DDMMYYYY") & "_" & Format(Time(), "HHMMSS") & ".txt"
           sNome_Arq = InputBox(Message, Title, sNome_Arq)
           
           Nada = Dir(sNome_Arq, vbNormal)
           
'           If Nada <> sNome_arq Then
'              MsgBox "Arquivo ou Diretório não encontrado em " + sNome_arq + " COPIA NÃO pode ser feita!"
'              Exit Sub
'           End If
           
           If Len(Trim(sNome_Arq)) > 0 Then
           
              Open sNome_Arq For Random Access Read Write As #21 Len = Len(Arq_Impressao)
              
              Ki = 0
              Kf = ListPreVw.ListCount - 1
              
              If ChkParte.Value = 1 Then
                 If RTrim(LBLLinha(0)) <> "" Then Ki = Val(LBLLinha(0))
                 If RTrim(LBLLinha(1)) <> "" Then Kf = Val(LBLLinha(1))
                 If Kf < Ki Then
                    MsgBox "Linha Inicial =" + Str(Ki) + "   MAIOR  que a Linha Final =" + Str(Kf), 16, "Impressão PARCIAL com Intervalo ERRADO"
                    Exit Sub
                 End If
              End If
              
              Me.PBar1.Visible = True
              Me.PBar1.Min = Ki
              Me.PBar1.Max = Kf
              X1 = 0
              
              For Zi = Ki To Kf
                  
                  ListPreVw.ListIndex = Zi
                  
                  X1 = X1 + 1
                  
                  Arq_Impressao.FCampo136 = Me.ListPreVw.List(Zi)
                  
                  Arq_Impressao.FFinal = Chr$(13) + Chr$(10)
                  Put 21, X1, Arq_Impressao
                  slinha = " "
                  Me.PBar1.Value = Zi
              
              Next
              
              Close #21
              ListPreVw.Visible = True
              Me.PBar1.Visible = False
              ListPreVw.ListIndex = 0
              Me.PBar1.Min = 0
           End If
        
        End If
'***********index=1 = impressao para texto ***********************************************
        
'***********index=2 = localizar texto no list ********************************************
        If Index = 2 Then
           If ListPreVw.ListCount = 0 Then Exit Sub
           Me.ListPreVw.SetFocus
           Me.ListPreVw.ListIndex = 0
           ListPreVw.Visible = False
           W2S = Text1.Text
           K = 0
           For z = 0 To ListPreVw.ListCount - 1
               ListPreVw.ListIndex = z
               W1S = ListPreVw
               K = InStr(1, W1S, W2S)
               If K > 0 Then
                  Exit For
               End If
           Next
           ListPreVw.Visible = True
           ZIni = z
           If K = 0 Then
              ListPreVw.ListIndex = 0
              MsgBox "Texto NAO encontrado no Relatório", 20, "Mensagem 0000:  Texto NAO encontrado"
              ZIni = 0
           End If
           cmdPrevW(5).Visible = True
        End If
'***********index=2 = localizar texto no list ********************************************
        
'***********index=3 = Aumentar fonte no list *********************************************
        If Index = 3 Then
           If ListPreVw.FontSize > 7 Then
              ListPreVw.FontSize = ListPreVw.FontSize - 1
              ListPreVw.Height = 4935
           End If
              If Me.ListPreVw.Width <= 11145 Then
                 Me.ListPreVw.Width = 11145
              Else
                 Me.ListPreVw.Width = Me.ListPreVw.Width - 1000
                 Me.Width = Me.Width - 1000
              End If
        End If
'***********index=3 = Aumentar fonte no list *********************************************
        
'***********index=4 = Diminuir fonte no list *********************************************
        If Index = 4 Then
           If ListPreVw.FontSize < 20 Then
              ListPreVw.FontSize = ListPreVw.FontSize + 1
              ListPreVw.Height = 4935
'              If Me.ListPreVw.Width < 11145 Then
'                 Me.ListPreVw.Width = 11145
'              Else
                 Me.ListPreVw.Width = Me.ListPreVw.Width + 1000
                 Me.Width = Me.Width + 1000
'              End If
           End If
        End If
'***********index=4 = Diminuir fonte no list *********************************************
        
'***********index=5 = localizar PROXIMO texto no list ********************************************
        If Index = 5 Then
           If ListPreVw.ListIndex = -1 Then Exit Sub
           Me.ListPreVw.SetFocus
           ListPreVw.Visible = False
           W2S = Text1.Text
           K = 0
           For z = ZIni + 1 To ListPreVw.ListCount - 1
               ListPreVw.ListIndex = z
               W1S = ListPreVw
               K = InStr(1, W1S, W2S)
               If K > 0 Then
                  Exit For
               End If
           Next
           ListPreVw.Visible = True
           ZIni = z
           If K = 0 Then
              ListPreVw.ListIndex = 0
              MsgBox "Texto NAO encontrado no Relatório", 20, "Mensagem 0000:  Texto NAO encontrado"
           End If
           cmdPrevW(5).Visible = True
        End If
        
'***********index=5 = localizar PROXIMO texto no list ********************************************

'***********index=6 = ler o arquivo no disco **********************************************
        If Index = 6 Then
           If Dir$(cNomeArq) = "" Then
              MsgBox "Diretorio de importação não encontrado, Procure o responsável! " & cNomeArq, 16, "Ação cancelada!"
              Exit Sub
           End If
'
           Me.Frame1.Visible = True
           Me.Dir1.Path = cNomeArq
        End If
'***********index=6 = lero arquivo no disco **********************************************

'***********index=9 = saida **************************************************************
        If Index = 9 Then
           Close 11
           Me.ListPreVw.Clear
           Me.Hide
        End If
'***********index=9 = saida **************************************************************

Exit Sub

Erro:

If Err.Number = 71 Then
   MsgBox "Diretótio para gravação do arquivo não encontrado. Arquivo - " & sNome_Arq
Else
   MsgBox "Erro não identificado, informe o responsavel o numero do erro - " & Err.Number
End If

End Sub

'Private Sub cmdPrevW_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
''MsgBox "cheguei44"
'End Sub

'Private Sub cmdPrevW_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Index = 7 Then
'   MsgBox "cheguei"
'End If
'End Sub

Private Sub File1_Click()
Me.cmd_confirmar.Enabled = True
End Sub
Private Sub Form_Activate()
Me.ListPreVw.FontName = "Courier New"
Me.ListPreVw.FontSize = 7
Me.ListPreVw.Font.Bold = False
Me.ListPreVw.Width = 11145
Me.Width = 11385
'Me.ListPreVw.Clear

If Flag_ativo = True Then
   Exit Sub
End If

Flag_ativo = True
Me.ChkAcompanha.Value = 0
Call Carrega_Diretorio_Impressao
Me.ListPreVw.FontName = "Courier New"
Me.ListPreVw.FontSize = 7
Me.ListPreVw.Font.Bold = False
bFormatacaoDisco = False
End Sub
'
Private Sub Form_Load()

Me.Top = 0
Me.Left = 0
Me.Width = 11800
Me.ListPreVw.FontName = "Courier New"
Me.ListPreVw.FontSize = 7
Me.ListPreVw.Font.Bold = False
frmGloPreview.Lst_Exel.Clear
'   W1S = "SERGIO MORAES VIEIRA - PRACA FLEMING 783 APARTAMENTO 501 JAQUEIRA RECIFE PERNAMBUCO CEP 52050-180 FONE/FAX 3221-1022"
End Sub
Public Function Carrega_Diretorio_Impressao()
'Dim cRec_func As ADODB.Recordset
'Dim Y As Integer
'
'On Error GoTo Erro
'Me.MousePointer = vbHourglass
'Set cRec_func = New ADODB.Recordset
'Set cRec_func = CCTempneAtributo.ESC_ATRIBUTO_Consultar("GES_GLOBAL", "GLO_DIRETORIO_IMP")
'
'If cRec_func Is Nothing Then
'   MsgBox "Não Existe parametro do arquivo de leitura do diretorio do arquivo, crie em GES_GLOBAL, variavel ->GLO_DIRETORIO_IMP e o caminho do arquivo."
'   Exit Function
'End If
'
'If cRec_func.RecordCount > 0 Then
'   cNomeArq = cRec_func!ges_valor
'   If Len(Trim(cNomeArq)) = 0 Then MsgBox "Não Existe parametro do arquivo de leitura do diretorio do arquivo, crie em GES_GLOBAL, variavel ->GLO_DIRETORIO_IMP e o caminho do arquivo."
'
'
'Else
'   MsgBox "Não Existe parametro do arquivo de leitura do diretorio do arquivo, crie em GES_GLOBAL, variavel ->GLO_DIRETORIO_IMP e o caminho do arquivo."
'End If
'
'Set cRec_func = Nothing
'Me.MousePointer = vbDefault
'
'Exit Function
'
'Erro:
'   Set cRec_func = Nothing
'   MsgBox Err.Description, , Me.Caption
'   Me.MousePointer = vbDefault

End Function
Private Sub ListPreVw_DblClick()

If ChkParte.Value = 1 Then
   If RTrim(LBLLinha(0)) = "" Then
      LBLLinha(0) = ListPreVw.ListIndex
   ElseIf RTrim(LBLLinha(1)) = "" Then
      LBLLinha(1) = ListPreVw.ListIndex
   Else
      LBLLinha(0) = ListPreVw.ListIndex
      LBLLinha(1) = ""
   End If
End If
End Sub

Private Sub ListPreVw_GotFocus()
       Me.ListPreVw.FontName = "Courier New"
'       Me.ListPreVw.FontSize = 7
End Sub

Private Sub ListPreVw_KeyPress(KeyAscii As Integer)
        If KeyAscii = 27 Then
           cmdPrevW(9).SetFocus
           SendKeys "{ENTER}"
        End If

End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
Me.cmd_confirmar.Enabled = False

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

'hoje eu preciso te encontrar de qualquer jeito
'nem que seja so para te levar para casa
'depois de um dia normal
'olhar teus olhos de promessas faceis e de beijar a boca
'de um jeito que te faça rir, que te faça rir
'hoje eu preciso te abraçar
'sentir teu cheiro de roupa limpa, pra esquecer os meus anseios e dormir em paz

''hoje eu preciso ouvir qualquer palavra tua, qualquer frase exagerada que me faça
''sentir alegria e estar vivo
''hoje eu presciso tomar um café ouvindo voce suspirar e dizendo
''que eu sou o causador da tua insônia que eu faço tudo errado sempre.... sempre....

'hoje preciso de voce com qualquer humor com qualquer sorriso
'hoje só tua presenca vai me deixar feliz.... só hoje

'Lára. larara larara larara

''hoje eu preciso ouvir qualquer palavra tua, qualquer frase exagerada que me faça
''sentir alegria e estar vivo
''hoje eu presciso tomar um café ouvindo voce suspirar e dizendo
''que eu sou o causador da tua insônia que eu faço tudo errado sempre.... sempre....

'hoje preciso de voce com qualquer humor com qualquer sorriso
'hoje só tua presenca vai me deixar feliz.... só hoje

'hoje preciso de voce com qualquer humor com qualquer sorriso
'hoje só tua presenca vai me deixar feliz.... feliz.... só hoje

'***************************************************************************

'Éu quero ficar so .... mas comigo so eu não consigo
'eu quero ficar junto mas sozinho só não é possivel
'é preciso amar direito um amor de qualquer jeito
'ser amor a qualquer hora ser amor de corpo inteiro
'amor de dentro para fora amor que eu desconheço

'quero um amor maior.... amor maior que eu (DEUS)
'quero um amor maior.... amor maior que eu (DEUS)
'
'Eu quero ficar so .... mas comigo so eu não consigo
'eu quero ficar junto mas sozinho assim não é possivel
'é preciso amar direito um amor de qualquer jeito
'ser amor a qualquer hora ser amor de corpo inteiro
'o amor de dentro para fora o amor que eu desconheço

'quero um amor maior.... amor maior que eu (DEUS)
'quero um amor maior.... amor maior que eu (DEUS)

'heir hier
'entao seguirei  meu coracao ate o fim pra saber se é amor
'pactuarei mesmo assim mesmo sem querer pra saber se é amor
'eu estarei mais feliz mesmo morrendo de dor
'heir
'para saber se é amor se é amor
'
'quero um amor maior.... o amor maior que eu (DEUS)
'quero um amor maior.... o amor maior que eu (DEUS)  o amor maior que eu....
'o amor maior que eu.....

'Lara rara
'Lara rara
'o amor maior que eu.... hammm
