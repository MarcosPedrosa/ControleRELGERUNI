VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGIFDisComposicaopallet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Composi��o do pallet"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   11625
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8490
      Top             =   5610
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   10230
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5670
      Width           =   1275
   End
   Begin VB.CommandButton cmdSelecionar 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Selecionar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   8880
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5670
      Width           =   1275
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   23
      Top             =   5610
      Width           =   1005
   End
   Begin vsOcx6LibCtl.vsIndexTab Vst_Pallet 
      Height          =   5505
      Left            =   120
      TabIndex        =   14
      Top             =   90
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   9710
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "&Filtro|&Composi��o"
      Align           =   0
      Appearance      =   1
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin vsOcx6LibCtl.vsElastic vsElastic2 
         Height          =   5130
         Left            =   12045
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   330
         Width           =   11310
         _ExtentX        =   19950
         _ExtentY        =   9049
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.CommandButton cmd_impresao 
            BackColor       =   &H00C0C0C0&
            Height          =   555
            Left            =   10410
            Picture         =   "frmGIFDisComposicaopallet.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   90
            Width           =   795
         End
         Begin MSFlexGridLib.MSFlexGrid mfl_gridcomp 
            Height          =   4365
            Left            =   60
            TabIndex        =   17
            Top             =   720
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   7699
            _Version        =   393216
            Cols            =   9
            AllowBigSelection=   0   'False
            TextStyle       =   3
            TextStyleFixed  =   2
            HighLight       =   2
            ScrollBars      =   2
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbl_PC 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   6990
            TabIndex        =   40
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Qt.Pc.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6360
            TabIndex        =   39
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lbl_CX 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   5490
            TabIndex        =   38
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Qt.Cx.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4860
            TabIndex        =   37
            Top             =   120
            Width           =   600
         End
         Begin VB.Label lbl_LOTE 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LOTE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   6780
            TabIndex        =   36
            Top             =   450
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Lote.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6150
            TabIndex        =   35
            Top             =   450
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lbl_CLIENTE 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CLIENTE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   3090
            TabIndex        =   34
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nome.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2460
            TabIndex        =   33
            Top             =   420
            Width           =   615
         End
         Begin VB.Label lbl_STATUS 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "STATUS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   3090
            TabIndex        =   32
            Top             =   90
            Width           =   810
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Status.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2460
            TabIndex        =   31
            Top             =   90
            Width           =   675
         End
         Begin VB.Label lbl_ID_CLIENTE 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "id_cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   870
            TabIndex        =   30
            Top             =   420
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cliente.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   29
            Top             =   420
            Width           =   720
         End
         Begin VB.Label lbl_PALLET 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "pallet"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   870
            TabIndex        =   28
            Top             =   90
            Width           =   540
         End
         Begin VB.Label lbl_1 
            AutoSize        =   -1  'True
            Caption         =   "Pallet.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   27
            Top             =   90
            Width           =   615
         End
      End
      Begin vsOcx6LibCtl.vsElastic vsElastic1 
         Height          =   5130
         Left            =   45
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   330
         Width           =   11310
         _ExtentX        =   19950
         _ExtentY        =   9049
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   0   'False
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.CommandButton cmd_imprime_pallet 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   405
            Left            =   10770
            Picture         =   "frmGIFDisComposicaopallet.frx":0532
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Imprimir os Pallet's encontrados"
            Top             =   4590
            Width           =   525
         End
         Begin VB.Frame Frame3 
            Height          =   1995
            Left            =   480
            TabIndex        =   19
            Top             =   30
            Width           =   10245
            Begin VB.ComboBox CBO_INVENTARIO 
               Height          =   315
               Left            =   7200
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   1590
               Width           =   1665
            End
            Begin VB.Frame Frame5 
               Caption         =   "Cliente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   180
               TabIndex        =   26
               Top             =   1050
               Width           =   6915
               Begin VB.ComboBox CBO_CLIENTE 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   90
                  TabIndex        =   5
                  Text            =   "Combo1"
                  Top             =   300
                  Width           =   6735
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1245
               Left            =   8370
               TabIndex        =   25
               Top             =   150
               Width           =   1665
               Begin VB.OptionButton Opt_faturado 
                  Caption         =   "Faturado"
                  Height          =   255
                  Left            =   150
                  TabIndex        =   6
                  Top             =   210
                  Width           =   945
               End
               Begin VB.OptionButton Opt_Embarcou 
                  Caption         =   "Embarcou"
                  Height          =   255
                  Left            =   150
                  TabIndex        =   7
                  Top             =   465
                  Value           =   -1  'True
                  Width           =   1155
               End
               Begin VB.OptionButton Opt_NFaturado 
                  Caption         =   "N�o Faturado"
                  Height          =   255
                  Left            =   150
                  TabIndex        =   8
                  Top             =   720
                  Width           =   1305
               End
               Begin VB.OptionButton Opt_Nafabrica 
                  Caption         =   "Na fabrica"
                  Height          =   255
                  Left            =   150
                  TabIndex        =   9
                  Top             =   960
                  Width           =   1215
               End
            End
            Begin VB.CommandButton cmd_Pesquisar 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Pesquisar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   9060
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   1530
               Width           =   975
            End
            Begin VB.Frame Frame1 
               Caption         =   "Intervalo de palet's"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   180
               TabIndex        =   20
               Top             =   180
               Width           =   7305
               Begin VB.CheckBox chk_pallet 
                  Caption         =   "Todos pallets"
                  Height          =   255
                  Left            =   6000
                  TabIndex        =   4
                  Top             =   360
                  Width           =   1275
               End
               Begin VB.ComboBox CBO_SEQ2 
                  Height          =   315
                  Left            =   5070
                  Style           =   2  'Dropdown List
                  TabIndex        =   3
                  Top             =   330
                  Width           =   780
               End
               Begin VB.ComboBox CBO_SEQ1 
                  Height          =   315
                  Left            =   2280
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   330
                  Width           =   780
               End
               Begin MSComCtl2.DTPicker dt_inicio 
                  Height          =   315
                  Left            =   870
                  TabIndex        =   0
                  Top             =   330
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   66256897
                  CurrentDate     =   40148
               End
               Begin MSComCtl2.DTPicker dt_final 
                  Height          =   315
                  Left            =   3660
                  TabIndex        =   2
                  Top             =   330
                  Width           =   1425
                  _ExtentX        =   2514
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   66256897
                  CurrentDate     =   40148
               End
               Begin VB.Label Label2 
                  Caption         =   "at�"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   22
                  Top             =   330
                  Width           =   435
               End
               Begin VB.Label Label3 
                  Caption         =   "De.:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   210
                  TabIndex        =   21
                  Top             =   330
                  Width           =   465
               End
            End
            Begin VB.Label Label10 
               Caption         =   "Invent�rio"
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
               Left            =   7230
               TabIndex        =   43
               Top             =   1290
               Width           =   1035
            End
         End
         Begin VB.Frame Frame2 
            Height          =   2955
            Left            =   450
            TabIndex        =   18
            Top             =   2070
            Width           =   10275
            Begin MSFlexGridLib.MSFlexGrid mfl_grid 
               Height          =   2685
               Left            =   90
               TabIndex        =   11
               Top             =   180
               Width           =   10125
               _ExtentX        =   17859
               _ExtentY        =   4736
               _Version        =   393216
               Cols            =   6
               AllowBigSelection=   0   'False
               TextStyle       =   3
               TextStyleFixed  =   2
               HighLight       =   2
               ScrollBars      =   2
               SelectionMode   =   1
               AllowUserResizing=   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
   End
   Begin VB.Label lbl_msg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Aguarde, Selecionando os Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3270
      TabIndex        =   45
      Top             =   5670
      Width           =   4245
   End
   Begin VB.Label Label4 
      Caption         =   "Total registros : "
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   5640
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFDisComposicaopallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Flag_ativo As Boolean 'Conter� true se o form ja foi ativado
Private cRec As ADODB.Recordset 'conter� os dados do registro corrente
Private Declare Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)


Private Sub cmd_impresao_Click()
Dim oTela As frmRelCristalReport
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim rs As ADODB.Recordset
Dim nx As Integer
Dim nValor As Double

On Error GoTo Erro

Set rs = New ADODB.Recordset

rs.Fields.Append "NUM_CAIXA", ADODB.DataTypeEnum.adVarChar, 10
rs.Fields.Append "TIPO_CAIXA", ADODB.DataTypeEnum.adVarChar, 8
rs.Fields.Append "PECA", ADODB.DataTypeEnum.adVarChar, 15
rs.Fields.Append "LOTE", ADODB.DataTypeEnum.adVarChar, 15
rs.Fields.Append "NF_VENDA", ADODB.DataTypeEnum.adVarChar, 16
rs.Fields.Append "ORDEM_VENDA", ADODB.DataTypeEnum.adVarChar, 11
rs.Fields.Append "SEQUENCIA", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "PLACA", ADODB.DataTypeEnum.adVarChar, 7
rs.Fields.Append "QTDEITENS", ADODB.DataTypeEnum.adDouble

rs.Open

cRec.MoveFirst

If cRec.RecordCount > 0 Then
'   Me.txtlidos.Text = cRec.RecordCount
   cRec.MoveFirst
   While Not cRec.EOF
       rs.AddNew "NUM_CAIXA", cRec!NUM_CAIXA
       rs.Fields("TIPO_CAIXA").Value = cRec!TIPO_CAIXA
       rs.Fields("PECA").Value = cRec!PECA
       rs.Fields("LOTE").Value = cRec!LOTE
       rs.Fields("NF_VENDA").Value = cRec!NF_VENDA
       rs.Fields("ORDEM_VENDA").Value = cRec!ORDEM_VENDA
       rs.Fields("SEQUENCIA").Value = cRec!SEQUENCIA
       rs.Fields("PLACA").Value = cRec!PLACA
       rs.Fields("QTDEITENS").Value = cRec!QTDE_NA_CAIXA
       rs.Update
       cRec.MoveNext
   Wend
Else
   MsgBox "Sem movimenta��o, Retorne."
   Exit Sub
End If

Set oTela = New frmRelCristalReport

Me.MousePointer = vbHourglass

Set CrystalReport1 = Application.OpenReport(App.Path & "\crptConferenciaPalllet.rpt")

CrystalReport1.Database.SetDataSource rs

CrystalReport1.ParameterFields(1).AddCurrentValue Me.lbl_PALLET.Caption & " Qtde. Cx.: " & lbl_CX.Caption & " Qtde.Pc.: " & lbl_PC.Caption & " Status : " & Me.lbl_STATUS.Caption
CrystalReport1.ParameterFields(1).DiscreteOrRangeKind = crDiscreteValue

CrystalReport1.ParameterFields(2).AddCurrentValue Me.lbl_CLIENTE.Caption
CrystalReport1.ParameterFields(2).DiscreteOrRangeKind = crDiscreteValue


oTela.CRV_RELATORIO.ReportSource = CrystalReport1
oTela.CRV_RELATORIO.ViewReport

rs.Clone

Me.MousePointer = vbDefault

oTela.Show 0

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_imprime_pallet_Click()
Dim oTela As frmRelCristalReport
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim rs As ADODB.Recordset
Dim nx As Integer
Dim nValor As Double
Dim sStatus As String

On Error GoTo Erro

'Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass

Set rs = New ADODB.Recordset

rs.Fields.Append "PALLET", ADODB.DataTypeEnum.adVarChar, 15
rs.Fields.Append "ID_CLIENTE", ADODB.DataTypeEnum.adVarChar, 10
rs.Fields.Append "CONF_PALLET", ADODB.DataTypeEnum.adVarChar, 10
rs.Fields.Append "CLIENTE", ADODB.DataTypeEnum.adVarChar, 25
rs.Fields.Append "CAIXA", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "QTDE", ADODB.DataTypeEnum.adDouble
rs.Open

cRec.MoveFirst

If cRec.RecordCount > 0 Then
   cRec.MoveFirst
   While Not cRec.EOF
       rs.AddNew "PALLET", IIf(IsNull(cRec!PALLET), " ", cRec!PALLET)
       rs.Fields("ID_CLIENTE").Value = IIf(IsNull(cRec!ID_CLIENTE), " ", cRec!ID_CLIENTE)
       rs.Fields("CONF_PALLET").Value = IIf(IsNull(cRec!CONF_PALLET), " ", cRec!CONF_PALLET)
       rs.Fields("CLIENTE").Value = IIf(IsNull(cRec!CLIENTE), " ", Mid$(cRec!CLIENTE, 1, 25))
       rs.Fields("CAIXA").Value = IIf(IsNull(cRec!CAIXA), " ", cRec!CAIXA)
       rs.Fields("QTDE").Value = IIf(IsNull(cRec!QTDE), " ", cRec!QTDE)
       rs.Update
       cRec.MoveNext
   Wend
Else
   MsgBox "Sem movimenta��o, Retorne."
   Exit Sub
End If

Set oTela = New frmRelCristalReport

Me.MousePointer = vbHourglass

Set CrystalReport1 = Application.OpenReport(App.Path & "\crptComposicaoPalllet.rpt")

CrystalReport1.Database.SetDataSource rs

CrystalReport1.ParameterFields(1).AddCurrentValue " "
CrystalReport1.ParameterFields(1).DiscreteOrRangeKind = crDiscreteValue

CrystalReport1.ParameterFields(2).AddCurrentValue Me.lbl_CLIENTE.Caption
CrystalReport1.ParameterFields(2).DiscreteOrRangeKind = crDiscreteValue


oTela.CRV_RELATORIO.ReportSource = CrystalReport1
oTela.CRV_RELATORIO.ViewReport

rs.Clone

Me.MousePointer = vbDefault

oTela.Show 0

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_Pesquisar_Click()
Dim nx As Integer
Dim sDataINI As String
Dim sDataFIM As String
Dim sSEQ1 As String
Dim sSEQ2 As String
Dim sStatus As String
Dim sCliente As String


On Error GoTo Erro

If Me.CBO_CLIENTE.ListCount = 0 Then Exit Sub

Me.txtlidos.Text = 0

Call Limpar_mfl_grid


If Me.chk_pallet.Value = 1 Then
   sDataINI = ""
   sDataFIM = ""
   sSEQ1 = ""
   sSEQ2 = ""
Else
   If Me.dt_inicio.Value > Me.dt_final.Value Then
      MsgBox "Data dos pallets, o inicio est� maior que o final, redigite!"
      Me.dt_inicio.SetFocus
      Exit Sub
   Else
      sDataINI = Format(Me.dt_inicio.Value, "yyyymmdd")
      sDataFIM = Format(Me.dt_final.Value, "yyyymmdd")
   End If
   sSEQ1 = Me.CBO_SEQ1.List(Me.CBO_SEQ1.ListIndex)
   sSEQ2 = Me.CBO_SEQ2.List(Me.CBO_SEQ2.ListIndex)
End If
   

If Me.Opt_faturado.Value = True Then
   sStatus = "1"
ElseIf Me.Opt_Embarcou.Value = True Then
   sStatus = "2"
ElseIf Me.Opt_NFaturado.Value = True Then
   sStatus = "3"
Else
   sStatus = "4"
End If

If Me.CBO_CLIENTE.ListIndex = 0 Then
   sCliente = ""
Else
   sCliente = Format(Me.CBO_CLIENTE.ItemData(Me.CBO_CLIENTE.ListIndex), "0000000000")
End If

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass

Set cRec = CCTemp.EXPEDICAO_Consultar_Comp_Pallet(sNomeBanco, _
                                                  sDataINI, _
                                                  sDataFIM, _
                                                  sSEQ1, _
                                                  sSEQ2, _
                                                  sStatus, _
                                                  sCliente, _
                                                  Me.CBO_INVENTARIO.ListIndex)

Me.mfl_grid.Visible = False
mfl_grid.Row = 0
mfl_grid.Col = 0: mfl_grid.ColWidth(0) = 1500: mfl_grid.Text = "PALLET"
mfl_grid.Col = 1: mfl_grid.ColWidth(1) = 1400: mfl_grid.Text = "ID_CLIENTE"
mfl_grid.Col = 2: mfl_grid.ColWidth(2) = 1200:  mfl_grid.Text = "STATUS"
mfl_grid.Col = 3: mfl_grid.ColWidth(3) = 4200: mfl_grid.Text = "CLIENTE"
'mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 1100: mfl_grid.Text = "LOTE"
'mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 1100: mfl_grid.Text = "PECA"
mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 700:  mfl_grid.Text = "CX"
mfl_grid.Col = 5: mfl_grid.ColWidth(5) = 700:  mfl_grid.Text = "PC"
mfl_grid.Col = 2: mfl_grid.BackColor = &H80FFFF

mfl_grid.Row = 0

mfl_grid.HighLight = False
mfl_grid.ColAlignment(0) = flexAlignLeftCenter
mfl_grid.ColAlignment(1) = flexAlignLeftCenter
mfl_grid.ColAlignment(2) = flexAlignLeftCenter
mfl_grid.ColAlignment(3) = flexAlignLeftCenter
mfl_grid.ColAlignment(4) = flexAlignCenterCenter
mfl_grid.ColAlignment(5) = flexAlignCenterCenter
'mfl_grid.ColAlignment(6) = flexAlignCenterCenter
'mfl_grid.ColAlignment(7) = flexAlignCenterCenter
mfl_grid.Row = 1

If cRec.RecordCount > 0 Then
   Me.cmd_imprime_pallet.Enabled = True
   Me.txtlidos.Text = cRec.RecordCount
   Me.cmdSelecionar.Enabled = True
   cRec.MoveFirst
   For nx = 1 To cRec.RecordCount
       mfl_grid.Col = 0: mfl_grid.Text = cRec.Fields("PALLET")
       mfl_grid.Col = 1: mfl_grid.Text = cRec.Fields("ID_CLIENTE")
       mfl_grid.Col = 2: mfl_grid.Text = cRec.Fields("CONF_PALLET")
       mfl_grid.Col = 3: mfl_grid.Text = cRec.Fields("CLIENTE")
'       mfl_grid.Col = 4: mfl_grid.Text = cRec.Fields("LOTE")
'       mfl_grid.Col = 4: mfl_grid.Text = cRec.Fields("ID_PECA")
       mfl_grid.Col = 4: mfl_grid.Text = cRec.Fields("CAIXA")
       mfl_grid.Col = 5: mfl_grid.Text = cRec.Fields("QTDE")
       cRec.MoveNext
       If Not cRec.EOF Then
          mfl_grid.Rows = mfl_grid.Rows + 1
          mfl_grid.Row = mfl_grid.Row + 1
       End If
   Next
Else
   Me.cmdSelecionar.Enabled = False
End If

Me.mfl_grid.Visible = True

''''Set GEXPesquisa.ADORecordset = rs
''''txtlidos.Text = rs.RecordCount
''''With GEXPesquisa
''''     .Columns(1).Caption = "C�digo"
''''     .Columns(1).Width = TextWidth("wwwwwww")
''''     .Columns(2).Caption = "nome"
''''     .Columns(2).Width = TextWidth("wwwwwwwww0wwwwwwwww0wwwwwwwww0")
'''''     For nx = 3 To .Columns.Count
'''''         .Columns(nx).Visible = False
'''''     Next nx
''''End With
''''Me.GEXPesquisa.SortKeys.Add 2, jgexSortAscending
Me.MousePointer = vbDefault

Rem Set rs = Nothing
Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub cmdSelecionar_Click()
Call Limpar_mfl_gridcompcomp
Call Pesquisar_composicao
End Sub


Private Sub Form_Activate()
If Flag_ativo = True Then
   Exit Sub
End If
Me.Top = 0
Me.Left = 0
Flag_ativo = True
Me.lbl_msg.Caption = "Aguarde, Selecionando os Clientes"
Me.Refresh

Call carregar_cliente
Me.lbl_msg.Caption = ""

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
      SendKeys "{TAB}"
ElseIf KeyCode = 27 Then
'   If Me.ActiveControl.TabIndex < 8 Then
'      If Me.CMD_SALVAR.Enabled = True Then
'        If 6 = MsgBox("Deseja realmente sair deste m�dulo?", 32 + 4) Then
'           Unload Me
'        End If
'      Else
        Unload Me
'      End If
'   Else
'       SendKeys "+{TAB}" ' retornar campo
'   End If
End If


End Sub

Private Sub Form_Load()
Dim nx As Integer

Me.Top = 0
Me.Left = 0

For nx = 1 To 99
    Me.CBO_SEQ1.AddItem nx
    Me.CBO_SEQ2.AddItem nx
Next

Me.CBO_INVENTARIO.AddItem "Todos"
Me.CBO_INVENTARIO.AddItem "N�o inventariado"
Me.CBO_INVENTARIO.AddItem "Com contagem"
Me.CBO_INVENTARIO.AddItem "1a Contagem"
Me.CBO_INVENTARIO.AddItem "2a Contagem"
Me.CBO_INVENTARIO.AddItem "3a Contagem"
Me.CBO_INVENTARIO.ListIndex = 0

Me.CBO_SEQ1.ListIndex = 0
Me.CBO_SEQ2.ListIndex = 0
Me.Vst_Pallet.CurrTab = 0
Me.Vst_Pallet.TabVisible(1) = False
Me.dt_inicio.Value = Format(Now(), "dd/mm/yyyy")
Me.dt_final.Value = Format(Now(), "dd/mm/yyyy")

mfl_grid.Row = 0
mfl_grid.Col = 0: mfl_grid.ColWidth(0) = 1500: mfl_grid.Text = "PALLET"
mfl_grid.Col = 1: mfl_grid.ColWidth(1) = 1400: mfl_grid.Text = "ID_CLIENTE"
mfl_grid.Col = 2: mfl_grid.ColWidth(2) = 1200:  mfl_grid.Text = "STATUS"
mfl_grid.Col = 3: mfl_grid.ColWidth(3) = 3500: mfl_grid.Text = "CLIENTE"
'mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 1100: mfl_grid.Text = "LOTE"
'mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 1100: mfl_grid.Text = "PECA"
mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 400:  mfl_grid.Text = "CX"
mfl_grid.Col = 5: mfl_grid.ColWidth(5) = 700:  mfl_grid.Text = "PC"
mfl_grid.Col = 2: mfl_grid.BackColor = &H80FFFF

End Sub

Public Function CCTemp() As neExpedicao
     Set CCTemp = New neExpedicao
End Function

Function carregar_cliente()
Dim nx As Integer

On Error GoTo Erro

Set cRec = New ADODB.Recordset
Sleep (1000)

Me.MousePointer = vbHourglass
Set cRec = rRec_cliente
'Set cRec = CCTemp.EXPEDICAO_Consultar_Cliente(sNomeBanco)

Me.CBO_CLIENTE.Clear
Me.CBO_CLIENTE.AddItem "TODOS OS CLIENTES"
Me.CBO_CLIENTE.ItemData(0) = 0

If cRec.RecordCount > 0 Then
   cRec.MoveFirst
   While Not cRec.EOF
       If Not IsNull(cRec!CLIENTE) Then
          nx = nx + 1
          Me.CBO_CLIENTE.AddItem cRec!ID_CLIENTE & " - " & Trim(cRec!CLIENTE)
          Me.CBO_CLIENTE.ItemData(nx) = cRec!ID_CLIENTE
       End If
       cRec.MoveNext
   Wend
Else
   MsgBox "N�o existem clientes na tabela - mov_etiq, procure o respons�vel."
End If

Me.CBO_CLIENTE.ListIndex = 0
Me.MousePointer = vbDefault

Rem Set cRec = Nothing
Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Function

Private Sub Limpar_mfl_grid()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

Me.mfl_grid.Visible = False
mfl_grid.Clear
nLinhas = mfl_grid.Rows

If mfl_grid.Rows > 2 Then
   For nx = mfl_grid.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then mfl_grid.RemoveItem (nx)
   Next
End If

mfl_grid.Row = 0
mfl_grid.Col = 0: mfl_grid.ColWidth(0) = 1500: mfl_grid.Text = "PALLET"
mfl_grid.Col = 1: mfl_grid.ColWidth(1) = 1400: mfl_grid.Text = "ID_CLIENTE"
mfl_grid.Col = 2: mfl_grid.ColWidth(2) = 1200:  mfl_grid.Text = "STATUS"
mfl_grid.Col = 3: mfl_grid.ColWidth(3) = 3500: mfl_grid.Text = "CLIENTE"
'mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 1100: mfl_grid.Text = "LOTE"
'mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 1100: mfl_grid.Text = "PECA"
mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 400:  mfl_grid.Text = "CX"
mfl_grid.Col = 5: mfl_grid.ColWidth(5) = 700:  mfl_grid.Text = "PC"
mfl_grid.Col = 2: mfl_grid.BackColor = &H80FFFF

mfl_grid.Row = 0

mfl_grid.HighLight = False
mfl_grid.ColAlignment(0) = flexAlignCenterCenter
mfl_grid.ColAlignment(1) = flexAlignLeftCenter
mfl_grid.ColAlignment(2) = flexAlignLeftCenter
mfl_grid.ColAlignment(3) = flexAlignLeftCenter
mfl_grid.ColAlignment(4) = flexAlignLeftCenter
mfl_grid.ColAlignment(5) = flexAlignCenterCenter
'mfl_grid.ColAlignment(6) = flexAlignCenterCenter
Me.mfl_grid.Visible = True
Me.cmd_imprime_pallet.Enabled = False


End Sub

Private Function Pesquisar_composicao()
Dim nx As Integer


On Error GoTo Erro

mfl_grid.Col = 0: Me.lbl_PALLET.Caption = mfl_grid.Text
mfl_grid.Col = 1: If Len(Trim(Me.lbl_PALLET.Caption)) = 0 Then Exit Function

Me.Vst_Pallet.TabVisible(1) = True
Me.Vst_Pallet.CurrTab = 1

mfl_grid.Col = 0: Me.lbl_PALLET.Caption = mfl_grid.Text
mfl_grid.Col = 1: Me.lbl_ID_CLIENTE.Caption = mfl_grid.Text
mfl_grid.Col = 2: Me.lbl_STATUS.Caption = mfl_grid.Text
mfl_grid.Col = 3: Me.lbl_CLIENTE.Caption = mfl_grid.Text
mfl_grid.Col = 5: Me.lbl_PC.Caption = mfl_grid.Text
mfl_grid.Col = 4: Me.lbl_CX.Caption = mfl_grid.Text
'mfl_grid.Col = 6: Me.lbl_PC.Caption = mfl_grid.Text

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass

Set cRec = CCTemp.EXPEDICAO_Consultar_Um_Pallet(sNomeBanco, Me.lbl_PALLET.Caption)

Me.mfl_gridcomp.Visible = False
mfl_gridcomp.Row = 0
mfl_gridcomp.Col = 0: mfl_gridcomp.ColWidth(0) = 1500: mfl_gridcomp.Text = "N�CAIXA"
mfl_gridcomp.Col = 1: mfl_gridcomp.ColWidth(1) = 1200: mfl_gridcomp.Text = "CAIXA"
mfl_gridcomp.Col = 2: mfl_gridcomp.ColWidth(2) = 1200: mfl_gridcomp.Text = "PECA"
mfl_gridcomp.Col = 3: mfl_gridcomp.ColWidth(3) = 1200: mfl_gridcomp.Text = "LOTE"
mfl_gridcomp.Col = 4: mfl_gridcomp.ColWidth(4) = 700: mfl_gridcomp.Text = "QTDE"
mfl_gridcomp.Col = 5: mfl_gridcomp.ColWidth(5) = 1250: mfl_gridcomp.Text = "N.FISCAL"
mfl_gridcomp.Col = 6: mfl_gridcomp.ColWidth(6) = 1300: mfl_gridcomp.Text = "ORD.VENDA"
mfl_gridcomp.Col = 7: mfl_gridcomp.ColWidth(7) = 1400: mfl_gridcomp.Text = "SEQUENCIA"
mfl_gridcomp.Col = 8: mfl_gridcomp.ColWidth(8) = 1100: mfl_gridcomp.Text = "PLACA"
'mfl_gridcomp.Col = 2: mfl_gridcomp.BackColor = &H80FFFF

mfl_gridcomp.Row = 0

mfl_gridcomp.HighLight = False
mfl_gridcomp.ColAlignment(0) = flexAlignCenterCenter
mfl_gridcomp.ColAlignment(1) = flexAlignCenterCenter
mfl_gridcomp.ColAlignment(2) = flexAlignLeftCenter
mfl_gridcomp.ColAlignment(3) = flexAlignLeftCenter
mfl_gridcomp.ColAlignment(4) = flexAlignLeftCenter
mfl_gridcomp.Row = 1

If cRec.RecordCount > 0 Then
   Me.txtlidos.Text = cRec.RecordCount
   cRec.MoveFirst
   For nx = 1 To cRec.RecordCount
       mfl_gridcomp.Col = 0: mfl_gridcomp.Text = cRec.Fields("NUM_CAIXA")
       mfl_gridcomp.Col = 1: mfl_gridcomp.Text = cRec.Fields("TIPO_CAIXA")
       mfl_gridcomp.Col = 2: mfl_gridcomp.Text = cRec.Fields("PECA")
       mfl_gridcomp.Col = 3: mfl_gridcomp.Text = cRec.Fields("LOTE")
       mfl_gridcomp.Col = 4: mfl_gridcomp.Text = cRec.Fields("QTDE_NA_CAIXA")
       mfl_gridcomp.Col = 5: mfl_gridcomp.Text = cRec.Fields("NF_VENDA")
       mfl_gridcomp.Col = 6: mfl_gridcomp.Text = cRec.Fields("ORDEM_VENDA")
       mfl_gridcomp.Col = 7: mfl_gridcomp.Text = cRec.Fields("SEQUENCIA")
       mfl_gridcomp.Col = 8: mfl_gridcomp.Text = cRec.Fields("PLACA")
       cRec.MoveNext
       If Not cRec.EOF Then
          mfl_gridcomp.Rows = mfl_gridcomp.Rows + 1
          mfl_gridcomp.Row = mfl_gridcomp.Row + 1
       End If
   Next
End If

Me.mfl_gridcomp.Visible = True

Me.MousePointer = vbDefault

Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault


End Function


Private Sub mfl_grid_DblClick()
Call Limpar_mfl_gridcompcomp
Call Pesquisar_composicao
End Sub

Private Sub Timer1_Timer()
If Len(Me.lbl_msg.Caption) > 0 Then
   If Me.lbl_msg.BackColor = &H8000000D Then
      Me.lbl_msg.BackColor = &HFF&
   Else
      Me.lbl_msg.BackColor = &H8000000D
   End If
Else
   Me.lbl_msg.Visible = False
End If
End Sub

Private Sub Vst_Pallet_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)

If NewTab = 0 Then
   Call Limpar_mfl_gridcompcomp
   cmd_Pesquisar_Click
   Me.Vst_Pallet.TabVisible(1) = False
   Me.cmdSelecionar.Visible = True
Else
   Me.cmdSelecionar.Visible = False
End If

End Sub
Private Sub Limpar_mfl_gridcompcomp()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

Me.mfl_gridcomp.Visible = False
mfl_gridcomp.Clear
nLinhas = mfl_gridcomp.Rows

If mfl_gridcomp.Rows > 2 Then
   For nx = mfl_gridcomp.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then mfl_gridcomp.RemoveItem (nx)
   Next
End If

mfl_gridcomp.Row = 0
mfl_gridcomp.Col = 0: mfl_gridcomp.ColWidth(0) = 1500: mfl_gridcomp.Text = "N�CAIXA"
mfl_gridcomp.Col = 1: mfl_gridcomp.ColWidth(1) = 1400: mfl_gridcomp.Text = "NF.VENDA"
mfl_gridcomp.Col = 2: mfl_gridcomp.ColWidth(2) = 1200: mfl_gridcomp.Text = "DT.PALET"
mfl_gridcomp.Col = 3: mfl_gridcomp.ColWidth(3) = 3500: mfl_gridcomp.Text = "TP.CAIXA"
mfl_gridcomp.Col = 4: mfl_gridcomp.ColWidth(4) = 1200: mfl_gridcomp.Text = "ORD.VENDA"
mfl_gridcomp.Col = 2: mfl_gridcomp.BackColor = &H80FFFF

mfl_gridcomp.Row = 0

mfl_gridcomp.HighLight = False
mfl_gridcomp.ColAlignment(0) = flexAlignCenterCenter
mfl_gridcomp.ColAlignment(1) = flexAlignLeftCenter
mfl_gridcomp.ColAlignment(2) = flexAlignLeftCenter
mfl_gridcomp.ColAlignment(3) = flexAlignLeftCenter
mfl_gridcomp.ColAlignment(4) = flexAlignLeftCenter
Me.mfl_gridcomp.Visible = True

End Sub

