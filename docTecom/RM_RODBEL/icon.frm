VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNotfIc 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualização Remota Rm -> RodBel, Versão 15/07/2010 1.0"
   ClientHeight    =   9810
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9690
   Icon            =   "icon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMD_INCLUIR_SECCAO 
      Caption         =   "INCLUIR SECAO"
      Height          =   555
      Left            =   6630
      TabIndex        =   67
      Top             =   8250
      Width           =   2025
   End
   Begin VB.Frame Frame7 
      Caption         =   "Mudança de Turno"
      DragIcon        =   "icon.frx":030A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1245
      Left            =   3390
      TabIndex        =   64
      Top             =   8310
      Width           =   2655
      Begin MSComCtl2.DTPicker txt_time_Func_Muda_Turno 
         Height          =   315
         Left            =   1170
         TabIndex        =   65
         Top             =   510
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16252930
         UpDown          =   -1  'True
         CurrentDate     =   39716.0416666667
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "1a Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   240
         Left            =   180
         TabIndex        =   66
         Top             =   540
         Width           =   960
      End
   End
   Begin VB.Frame FrmAcesso 
      BackColor       =   &H00008000&
      Caption         =   "      Digite a senha de Acesso     "
      Height          =   975
      Left            =   180
      TabIndex        =   56
      Top             =   5610
      Width           =   2595
      Begin VB.TextBox txt_senha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   11
         PasswordChar    =   "*"
         TabIndex        =   57
         Text            =   "20080101ACS"
         ToolTipText     =   "Digite a senha para obter acesso a tela de manutenção."
         Top             =   330
         Width           =   2085
      End
   End
   Begin MSComctlLib.ProgressBar Pr_Prog 
      Height          =   435
      Left            =   3150
      TabIndex        =   35
      Top             =   4620
      Visible         =   0   'False
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Frame frm_impressao 
      BackColor       =   &H00FF8080&
      Caption         =   "Filtro das informações do LOG"
      Height          =   2535
      Left            =   3240
      TabIndex        =   20
      Top             =   5460
      Visible         =   0   'False
      Width           =   5745
      Begin VB.CommandButton cmd_Impressao 
         Height          =   645
         Left            =   4590
         Picture         =   "icon.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1800
         Width           =   765
      End
      Begin VB.ComboBox CBO_ACOES2 
         BackColor       =   &H00FFC0C0&
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
         Left            =   2190
         TabIndex        =   26
         Text            =   "CBO_ACOES2"
         Top             =   1290
         Width           =   3165
      End
      Begin MSComCtl2.DTPicker DT_Filtro_ini 
         Height          =   285
         Left            =   1170
         TabIndex        =   22
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16252929
         CurrentDate     =   39720
      End
      Begin MSComCtl2.DTPicker DT_Filtro_fim 
         Height          =   285
         Left            =   3990
         TabIndex        =   23
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16252929
         CurrentDate     =   39720
      End
      Begin MSComCtl2.DTPicker DT_Hora_ini 
         Height          =   315
         Left            =   1170
         TabIndex        =   24
         Top             =   720
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16252930
         UpDown          =   -1  'True
         CurrentDate     =   39716.0000115741
      End
      Begin MSComCtl2.DTPicker DT_Hora_fim 
         Height          =   315
         Left            =   3990
         TabIndex        =   25
         Top             =   720
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   32768
         Format          =   16252930
         UpDown          =   -1  'True
         CurrentDate     =   39716.9999884259
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   5670
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   5670
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Atualização de.:"
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
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1680
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "H.Final.:"
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
         Left            =   2970
         TabIndex        =   30
         Top             =   750
         Width           =   870
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "H.Inicial.:"
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
         Left            =   120
         TabIndex        =   29
         Top             =   750
         Width           =   975
      End
      Begin VB.Line Line7 
         X1              =   90
         X2              =   5640
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Dt.Final.:"
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
         Left            =   3030
         TabIndex        =   28
         Top             =   300
         Width           =   930
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Dt.Inicial.:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmd_Log 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Log das Ações realizadas"
      Height          =   405
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Clique para aparecer/Desaparecer filtro da impressão do LOG."
      Top             =   4650
      Width           =   2865
   End
   Begin VB.Frame frm_principal 
      Caption         =   "Horários Para Começar Procesos de :"
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
      Height          =   4485
      Left            =   90
      TabIndex        =   11
      Top             =   90
      Width           =   9495
      Begin VB.Frame Frame9 
         Caption         =   "Funcionários Atestado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   1245
         Left            =   150
         TabIndex        =   59
         ToolTipText     =   "Ajuste os horários a que horas o sistema vai fazer as novas Atualizações dos funcionários afastados"
         Top             =   1710
         Width           =   2655
         Begin MSComCtl2.DTPicker txt_time_Func_Atest_Ini 
            Height          =   315
            Left            =   1140
            TabIndex        =   60
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   32768
            CalendarTitleForeColor=   32768
            Format          =   16252930
            UpDown          =   -1  'True
            CurrentDate     =   39716.5208333333
         End
         Begin MSComCtl2.DTPicker txt_time_Func_Atest_Fim 
            Height          =   315
            Left            =   1140
            TabIndex        =   61
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   32768
            Format          =   16252930
            UpDown          =   -1  'True
            CurrentDate     =   39716.9791666667
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "1a Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   240
            Left            =   90
            TabIndex        =   63
            Top             =   390
            Width           =   960
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "2a Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   240
            Left            =   90
            TabIndex        =   62
            Top             =   840
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Funcionários Afastados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1245
         Left            =   6630
         TabIndex        =   51
         ToolTipText     =   "Ajuste os horários a que horas o sistema vai fazer as novas Atualizações dos funcionários afastados"
         Top             =   390
         Width           =   2655
         Begin MSComCtl2.DTPicker txt_time_Func_Afast_Ini 
            Height          =   315
            Left            =   1110
            TabIndex        =   52
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   32768
            CalendarTitleForeColor=   32768
            Format          =   16252930
            UpDown          =   -1  'True
            CurrentDate     =   39716.5208333333
         End
         Begin MSComCtl2.DTPicker txt_time_Func_Afast_Fim 
            Height          =   315
            Left            =   1140
            TabIndex        =   53
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   32768
            Format          =   16252930
            UpDown          =   -1  'True
            CurrentDate     =   39716.9791666667
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "2a Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   90
            TabIndex        =   55
            Top             =   810
            Width           =   960
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "1a Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   90
            TabIndex        =   54
            Top             =   390
            Width           =   960
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Período de pesquisa de acesso aos dados dos funcionários no RM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   3090
         TabIndex        =   32
         Top             =   3090
         Width           =   6225
         Begin VB.TextBox txt_Dias_Antecedencia 
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
            Left            =   2520
            MaxLength       =   2
            TabIndex        =   5
            Text            =   "05"
            Top             =   300
            Width           =   345
         End
         Begin MSComCtl2.DTPicker DT_Pesquisa 
            Height          =   345
            Left            =   4470
            TabIndex        =   6
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   16252929
            CurrentDate     =   39721
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Dt.Pesquisa.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   3030
            TabIndex        =   34
            Top             =   330
            Width           =   1395
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Dias de antecedência.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   90
            TabIndex        =   33
            Top             =   330
            Width           =   2415
         End
      End
      Begin VB.CommandButton CMD_ACAO 
         BackColor       =   &H0080FF80&
         Caption         =   "Confirma Atualização"
         Height          =   825
         Left            =   150
         Picture         =   "icon.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3510
         Width           =   2715
      End
      Begin VB.ComboBox CBO_ACOES 
         BackColor       =   &H0080FF80&
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
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2745
      End
      Begin VB.Frame Frame6 
         Caption         =   "Mudança de Setor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   1245
         Left            =   6630
         TabIndex        =   18
         ToolTipText     =   "Ajuste o minuto, que o sistema verificará mudança de Setor. Será feita no minuto/segundos de cada hora."
         Top             =   1710
         Width           =   2655
         Begin MSComCtl2.DTPicker txt_time_Func_Muda_Setor 
            Height          =   315
            Left            =   630
            TabIndex        =   4
            Top             =   630
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   16252930
            UpDown          =   -1  'True
            CurrentDate     =   39716.0416666667
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "No minuto da Hora"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C0C0&
            Height          =   240
            Left            =   390
            TabIndex        =   19
            Top             =   300
            Width           =   1950
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Funcionários Desligados"
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
         Left            =   3390
         TabIndex        =   17
         ToolTipText     =   "Ajuste o minuto e segundo, que o sistema verificará os funcionários desligados. Será feita no minuto/segundo  de cada hora."
         Top             =   1710
         Width           =   2655
         Begin MSComCtl2.DTPicker txt_time_Func_Desligados 
            Height          =   315
            Left            =   630
            TabIndex        =   2
            Top             =   630
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   16252930
            UpDown          =   -1  'True
            CurrentDate     =   39716.0416666667
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "No minuto da Hora"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   300
            TabIndex        =   36
            Top             =   270
            Width           =   1950
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Funcionários de Férias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1245
         Left            =   3390
         TabIndex        =   15
         ToolTipText     =   "Verificará os funcionarios que estaram entrando ou saindo de férias. Verifica na hora exata que você atualizou."
         Top             =   390
         Width           =   2655
         Begin MSComCtl2.DTPicker txt_time_Func_ferias 
            Height          =   315
            Left            =   1200
            TabIndex        =   3
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12632319
            CalendarForeColor=   16711680
            CalendarTitleBackColor=   65535
            CalendarTitleForeColor=   16711680
            CalendarTrailingForeColor=   65535
            Format          =   16252930
            UpDown          =   -1  'True
            CurrentDate     =   39716.0416666667
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Na Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   150
            TabIndex        =   16
            Top             =   510
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Funcionários Novatos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1245
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "Ajuste os horários a que horas o sistema vai fazer as novas inclusões dos novos funcionários"
         Top             =   390
         Width           =   2655
         Begin MSComCtl2.DTPicker txt_time_Func_Novo1 
            Height          =   315
            Left            =   1140
            TabIndex        =   0
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   32768
            CalendarTitleForeColor=   32768
            Format          =   16252930
            UpDown          =   -1  'True
            CurrentDate     =   39716.5208333333
         End
         Begin MSComCtl2.DTPicker txt_time_Func_Novo2 
            Height          =   315
            Left            =   1140
            TabIndex        =   1
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   32768
            Format          =   16252930
            UpDown          =   -1  'True
            CurrentDate     =   39716.9791666667
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "1a Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Left            =   90
            TabIndex        =   14
            Top             =   390
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "2a Hora.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Left            =   90
            TabIndex        =   13
            Top             =   810
            Width           =   960
         End
      End
      Begin VB.Label LBL_MSG 
         Caption         =   "PARA ACESSO AOS PARAMETROS TECLE <ALT> + ""S"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3150
         TabIndex        =   58
         Top             =   4020
         Width           =   6255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10710
      Top             =   4860
   End
   Begin VB.PictureBox pichook 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9210
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   5130
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Frame Frame4 
      Caption         =   "Funcionários Afastados por :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3945
      Left            =   0
      TabIndex        =   37
      Top             =   6210
      Visible         =   0   'False
      Width           =   3105
      Begin MSComCtl2.DTPicker txt_time_Func_AF_LM 
         Height          =   315
         Left            =   1320
         TabIndex        =   38
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16252930
         UpDown          =   -1  'True
         CurrentDate     =   39716.0416666667
      End
      Begin MSComCtl2.DTPicker txt_time_Func_AF_PRE1 
         Height          =   315
         Left            =   1320
         TabIndex        =   39
         Top             =   1590
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16252930
         UpDown          =   -1  'True
         CurrentDate     =   39716.0416666667
      End
      Begin MSComCtl2.DTPicker txt_time_Func_AF_PRE2 
         Height          =   315
         Left            =   1320
         TabIndex        =   40
         Top             =   2010
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16252930
         UpDown          =   -1  'True
         CurrentDate     =   39716.0416666667
      End
      Begin MSComCtl2.DTPicker txt_time_Func_AF_LMA1 
         Height          =   315
         Left            =   1350
         TabIndex        =   41
         Top             =   2850
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16252930
         UpDown          =   -1  'True
         CurrentDate     =   39716.0416666667
      End
      Begin MSComCtl2.DTPicker txt_time_Func_AF_LMA2 
         Height          =   315
         Left            =   1350
         TabIndex        =   42
         Top             =   3270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16252930
         UpDown          =   -1  'True
         CurrentDate     =   39716.0416666667
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "1a Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   300
         TabIndex        =   50
         Top             =   750
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   120
         X2              =   780
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Licença Médica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   840
         TabIndex        =   49
         Top             =   390
         Width           =   1365
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   2220
         X2              =   2880
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   120
         X2              =   990
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Previdëncia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1020
         TabIndex        =   48
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   2100
         X2              =   2910
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "1a Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   300
         TabIndex        =   47
         Top             =   1620
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "2a Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   300
         TabIndex        =   46
         Top             =   2040
         Width           =   960
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   150
         X2              =   570
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Licença Maternidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   630
         TabIndex        =   45
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "1a Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   300
         TabIndex        =   44
         Top             =   2880
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "2a Hora.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   300
         TabIndex        =   43
         Top             =   3300
         Width           =   960
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   2460
         X2              =   2880
         Y1              =   2640
         Y2              =   2640
      End
   End
   Begin VB.Menu mnu_taskbar 
      Caption         =   "mnu_taskbar"
      Visible         =   0   'False
      Begin VB.Menu mnu_sobre 
         Caption         =   "&Atualizar Parametros..."
      End
      Begin VB.Menu mnutraco 
         Caption         =   "-"
      End
      Begin VB.Menu mnusair 
         Caption         =   "Sa&ir"
      End
   End
End
Attribute VB_Name = "frmNotfIc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sBancoRodbel As String
Public sBancoRM As String
Public rs As ADODB.Recordset
Public nMinuto As Integer

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA


Private Sub CBO_ACOES_Change()

End Sub


Private Sub CMD_ACAO_Click()

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Novatos" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Novatos
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Férias" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Ferias
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Afastamentos" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Afastamentos
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Desligados" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Desligados
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Mudança Setor" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Secoes
End If

If Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "Atestado" Or _
   Me.CBO_ACOES.List(Me.CBO_ACOES.ListIndex) = "TODAS AS AÇÕES ACIMA" Then
   Call Atualizacoes_Atestados
End If

Me.Pr_Prog.Visible = False

End Sub

Private Sub cmd_Impressao_Click()

Dim oTela As frmGloPreview
Dim rs As New ADODB.Recordset
Dim sTipo As String
Dim Y As Double
Dim nx As Integer
Dim sString As String
Dim nreg As Integer

On Error GoTo Erro

Set oTela = New frmGloPreview

Me.MousePointer = vbHourglass

Open App.Path & "\LogRmRodBel.TXT" For Random Access Read Write As #11 Len = Len(sTexto)

Y = LOF(11) / Len(sTexto)
nreg = 0

For nx = 1 To Y


   Get 11, nx, sTexto
   If (CDate(Mid$(sTexto.Texto, 3, 10)) >= CDate(Me.DT_Filtro_ini.Value) And _
      CDate(Mid$(sTexto.Texto, 3, 10)) <= CDate(Me.DT_Filtro_fim.Value)) And _
      (Mid$(sTexto.Texto, 14, 5) >= Mid$(DT_Hora_ini.Value, 12, 5) And _
      Mid$(sTexto.Texto, 14, 5) <= Mid$(DT_Hora_fim.Value, 12, 5)) Then
      If Me.CBO_ACOES2.ListIndex = 0 Or _
         Me.CBO_ACOES2.List(Me.CBO_ACOES2.ListIndex) = Trim(Mid$(sTexto.Texto, 26, 29)) Then
         nreg = nreg + 1
         If Mid$(sTexto.Texto, 1, 1) = "0" Then
            sString = "Ok "
         ElseIf Mid$(sTexto.Texto, 1, 1) = "1" Then
            sString = "Er "
         Else
            sString = "Es "
         End If
         sString = sString & Mid$(sTexto.Texto, 3, 10) & " "
         sString = sString & Mid$(sTexto.Texto, 14, 5) & " "
         sString = sString & Mid$(sTexto.Texto, 20, 5) & " "
         sString = sString & Mid$(sTexto.Texto, 26, 29) & " "
         sString = sString & Mid$(sTexto.Texto, 56, 60)
         oTela.ListPreVw.AddItem sString
      End If
   End If
Next

Me.MousePointer = vbDefault

If nreg = 0 Then
   MsgBox "Não há registros com este filtro. Altere o filtro para nova consulta."
   Exit Sub
End If

oTela.Show 0

Me.frm_impressao.Visible = False
Me.frm_principal.Visible = True
Me.frm_impressao.Top = 1170

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault
Close (11)
Me.frm_impressao.Visible = False
Me.frm_principal.Visible = True
Me.frm_impressao.Top = 1170

End Sub

Private Sub cmd_Log_Click()
If Me.frm_impressao.Visible = True Then
   Me.frm_impressao.Top = 1170
   Me.frm_impressao.Visible = False
   Me.frm_principal.Visible = True
Else
   Me.frm_impressao.Top = 1170
   Me.frm_impressao.Visible = True
   Me.frm_principal.Visible = False
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 83 And Shift = 4 Then
   Me.FrmAcesso.Top = 1770
   Me.FrmAcesso.Left = 3640
   Me.FrmAcesso.Enabled = True
   Me.txt_senha.Text = ""
   Me.txt_senha.SetFocus
End If

End Sub

Private Sub Form_Load()
Dim dDate As Date
Dim cFields As Collection

Static vShowMsg As Variant 'mostra mensagem 1a vez

Rem teste de acesso(retirar)
Rem Call Ajusta_Acesso

If App.PrevInstance Then
   MsgBox "Este Programa JÁ esta sendo processado neste computador", 16, "<ENTER>=Para Finalizar"
   Close: End
End If


If IsDate(Mid(Now(), 1, 10)) = False Then
   MsgBox "O seu computador está com o formato da DATA DIFERENTE DO PADRÃO dd/mm/yyyy. Altere as Configurações Regionais , no Painel de Controle."
   End
End If

dDate = Mid(Now(), 1, 10)
If Len(Trim(dDate)) <> 10 Then
   MsgBox "O seu computador está com o formato da DATA DIFERENTE DO PADRÃO dd/mm/yyyy. Altere as Configurações Regionais , no Painel de Controle."
   End
End If



Rem ***************************************
Rem ********  VARIAVEIS DE ACESSO A BANCO
Rem ***************************************
Call Variaveis_Acesso_Banco
Rem ***************************************

If IsEmpty(vShowMsg) Or vShowMsg = 1 Then
    MsgBox "A aplicação será reduzida a um ícone no lado direito da Barra de Tarefas do Windows.", 64, "Atualização Remota Rm -> RodBel"
    vShowMsg = 2
End If
 
t.cbSize = Len(t)
t.hWnd = pichook.hWnd
t.uId = 1&
t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
t.ucallbackMessage = WM_MOUSEMOVE
t.hIcon = Me.Icon
t.szTip = "RM_RODBEL, Comunicação Remota..." & Chr$(0) 'Texto a ser exibido quando o mouse é movido sobre o ícone.
Shell_NotifyIcon NIM_ADD, t
Me.Hide
App.TaskVisible = False

Rem CARREGAR OS COMBO E SUAS ACOES

Me.CBO_ACOES.AddItem "Novatos"
Me.CBO_ACOES.AddItem "Férias"
Me.CBO_ACOES.AddItem "Afastamentos"
Me.CBO_ACOES.AddItem "Desligados"
Me.CBO_ACOES.AddItem "Mudança Setor"
'Me.CBO_ACOES.AddItem "Mudança de Turno"
Me.CBO_ACOES.AddItem "Atestado"
Me.CBO_ACOES.AddItem "TODAS AS AÇÕES ACIMA"
Me.CBO_ACOES.ListIndex = 0

Me.CBO_ACOES2.AddItem "Novatos"
Me.CBO_ACOES2.AddItem "Férias"
Me.CBO_ACOES2.AddItem "Afastamentos"
Me.CBO_ACOES2.AddItem "Desligados"
Me.CBO_ACOES2.AddItem "Mudança Setor"
'Me.CBO_ACOES2.AddItem "Mudança de Turno"
Me.CBO_ACOES2.AddItem "Atestado"
Me.CBO_ACOES2.AddItem "TODAS AS AÇÕES ACIMA"
Me.CBO_ACOES2.ListIndex = 0

'Me.CBO_ACOES2.AddItem "Todos"
'Me.CBO_ACOES2.AddItem "Novatos"
'Me.CBO_ACOES2.AddItem "Férias"
'Me.CBO_ACOES2.AddItem "Afastamentos"
'Me.CBO_ACOES2.AddItem "Desligados"
'Me.CBO_ACOES2.AddItem "Mudança Setor"
'Me.CBO_ACOES2.AddItem "Mudança de Turno"
'Me.CBO_ACOES2.AddItem "TODAS AS AÇÕES ACIMA"
'Me.CBO_ACOES2.ListIndex = 0

Rem   REGISTRAR A HORA EM QUA O SISTEMA FOI STARTADO
sStatusMsg = "0"
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Sistema Ligado"
sCodFun = "0000"
sMsg = "Sistema Ligado no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & _
                   sData & ";" & _
                   sHora & ";" & _
                   sCodFun & ";" & _
                   sTipo & ";" & _
                   sMsg

Call CCTempneRegBanco.Gerar_Situacao_Log(cFields)

Me.DT_Filtro_ini.Value = "01/" & Format(Now(), "mm/yyyy")

If IsDate("28/" & Format(Now(), "mm/yyyy")) Then Me.DT_Filtro_fim.Value = "28/" & Format(Now(), "mm/yyyy")
If IsDate("29/" & Format(Now(), "mm/yyyy")) Then Me.DT_Filtro_fim.Value = "29/" & Format(Now(), "mm/yyyy")
If IsDate("30/" & Format(Now(), "mm/yyyy")) Then Me.DT_Filtro_fim.Value = "30/" & Format(Now(), "mm/yyyy")
If IsDate("31/" & Format(Now(), "mm/yyyy")) Then Me.DT_Filtro_fim.Value = "31/" & Format(Now(), "mm/yyyy")

Me.DT_Pesquisa.Value = CDate(Format(Now(), "dd/mm/yyyy")) - Val(Me.txt_Dias_Antecedencia.Text)

Me.Caption = "Atualização Remota Rm -> RodBel - Data Sistema " & Format(Now(), "dd/mm/yyyy")
Set cFields = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim cFields As Collection

Rem   REGISTRAR A HORA EM QUA O SISTEMA FOI DESLIGADO
sStatusMsg = "0"
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Sistema Desligado"
sCodFun = "0000"
sMsg = "Sistema Desligado no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & _
                   sData & ";" & _
                   sHora & ";" & _
                   sCodFun & ";" & _
                   sTipo & ";" & _
                   sMsg

Call CCTempneRegBanco.Gerar_Situacao_Log(cFields)
    
t.cbSize = Len(t)
t.hWnd = pichook.hWnd
t.uId = 1&
Shell_NotifyIcon NIM_DELETE, t  'Remove o ícone da barra de tarefas.
    
End Sub

Private Sub Form_Resize()
    If (Me.WindowState) = 1 Then
        Me.Hide
    End If
End Sub

Private Sub mnu_sobre_Click()
    Me.WindowState = 0
    Me.Show
End Sub

Private Sub mnusair_Click()
    Unload Me
    End
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'pichook é uma picture box, utilizada pelo Windows para
'reconhecer o ícone na barra de tarefas.
    Static rec As Boolean, msg As Long
    msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case msg
            Case WM_LBUTTONDBLCLK:
                 Me.PopupMenu mnu_taskbar
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
        'Se for pressionado o botão direito
        'sobre o ícone, é exibido um menu pop-up.
                Me.PopupMenu mnu_taskbar    'mnuBar-menu criado no form.
        End Select
        rec = False
    End If
End Sub

Private Sub Timer1_Timer()

Dim tTempo As String
Dim tTempo2 As String

Rem aqui marcos

Exit Sub




tTempo = Format(Now(), "hh:mm:ss")
tTempo2 = Format(Now(), "hh:mm:ss")

Rem ***************************************
Rem ATUALIZAR A DATA DO PERIODO DE PESQUISA
Rem ***************************************
If Val(Me.txt_Dias_Antecedencia.Text) > 0 Then
   Me.DT_Pesquisa.Value = CDate(Format(Now(), "dd/mm/yyyy")) - Val(Me.txt_Dias_Antecedencia.Text)
Else
   Me.DT_Pesquisa.Value = CDate(Format(Now(), "dd/mm/yyyy"))
End If
Rem ***************************************

Rem ****************************************************************
Rem verificar tempos para Inclusao de novos funcionários
Rem ****************************************************************
    If tTempo = Format(txt_time_Func_Novo1.Value, "hh:mm:ss") Then
        Call Atualizacoes_Novatos
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If

    If tTempo = Format(txt_time_Func_Novo2.Value, "hh:mm:ss") Then
        Call Atualizacoes_Novatos
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem Verificar tempos para demissao dos funcionários
Rem ****************************************************************
    If Mid$(tTempo, 4, 5) = Mid$(Format(txt_time_Func_Desligados.Value, "hh:mm:ss"), 4, 5) Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Desligados.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Desligados.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Desligados
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem verificar tempos para FERIAS dos funcionários
Rem ****************************************************************
    If tTempo = Format(txt_time_Func_ferias.Value, "hh:mm:ss") Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_ferias.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_ferias.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Ferias
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem verificar tempos para Mudanca de Secao dos funcionários
Rem ****************************************************************
    If Mid$(tTempo, 4, 5) = Mid$(Format(txt_time_Func_Muda_Setor.Value, "hh:mm:ss"), 4, 5) Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Muda_Setor.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Muda_Setor.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Secoes
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem verificar tempos para Afastamentos de funcionários
Rem ****************************************************************
    If tTempo = Format(txt_time_Func_Afast_Ini.Value, "hh:mm:ss") Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Afast_Ini.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Afast_Ini.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Afastamentos
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
    
    If tTempo = Format(txt_time_Func_Afast_Fim.Value, "hh:mm:ss") Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Afast_Fim.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Afast_Fim.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Afastamentos
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************

Rem ****************************************************************
Rem verificar tempos para Atestados de funcionários
Rem ****************************************************************
    If tTempo = Format(txt_time_Func_Atest_Ini.Value, "hh:mm:ss") Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Atest_Ini.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Atest_Ini.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Atestados
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
    
    If tTempo = Format(txt_time_Func_Atest_Fim.Value, "hh:mm:ss") Or _
       (Mid$(tTempo2, 4, 5) >= Mid$(Format(txt_time_Func_Atest_Fim.Value, "hh:mm:ss"), 4, 5) And _
        Mid$(tTempo, 4, 5) <= Mid$(Format(txt_time_Func_Atest_Fim.Value, "hh:mm:ss"), 4, 5)) Then
        Call Atualizacoes_Atestados
        tTempo2 = Format(Now(), "hh:mm:ss")
    End If
Rem ****************************************************************
Rem ****************************************************************




Me.Pr_Prog.Visible = False

Rem ****************************************************
Rem CONTROLE DE MENSAGEM PISCANTE E MENSAGEM A SER DITA
Rem ****************************************************

If Me.LBL_MSG.Visible = True Then
   Me.LBL_MSG.Visible = False
Else
   Me.LBL_MSG.Visible = True
End If
Rem ****************************************************

Rem ****************************************************
Rem CONTROLE DE MENSAGEM PISCANTE E MENSAGEM A SER DITA
Rem ****************************************************
If Me.frm_principal.Enabled = True Then
   nMinuto = nMinuto - 1
   If nMinuto < 0 Then
      Me.frm_principal.Enabled = False
   End If
   Me.LBL_MSG.Caption = "ACESSO LIBERADO ATUALIZE OS PARAMETROS.(T) " & Str(nMinuto) & " s."
'   Me.LBL_MSG.ForeColor = &H8000&
Else
   Me.LBL_MSG.Caption = "PARA ACESSO AOS PARAMETROS TECLE <ALT> + 'S'"
'   Me.LBL_MSG.ForeColor = &H8000000F
End If
Rem ****************************************************

Me.Caption = "Atualização Remota Rm -> RodBel - Data Sistema " & Format(Now(), "dd/mm/yyyy") & " Hora : " & Format(Now(), "hh:mm:ss")

End Sub

Private Sub Variaveis_Acesso_Banco()
Rem teste da data do computador para o formato dd/mm/yyyy

If IsDate(Mid(Now(), 1, 10)) = False Then
   MsgBox "O seu computador está com o formato da DATA DIFERENTE DO PADRÃO dd/mm/yyyy. Altere as Configurações Regionais , no Painel de Controle."
   End
End If

If Len(Trim(Mid(Now(), 1, 10))) <> 10 Then
   MsgBox "O seu computador está com o formato da DATA DIFERENTE DO PADRÃO dd/mm/yyyy. Altere as Configurações Regionais , no Painel de Controle."
   End
End If
 
 
Rem *************************  A T E N Ç Ã O *****************************************
Rem *************************  BASE TESTE SQL   *****************************************
Rem **********************************************************************************
' sBancoRodbel = "Provider=SQLOLEDB.1;" & _
'                 "Password=sa;" & _
'                 "Persist Security Info=True;" & _
'                 "User ID=sa;" & _
'                 "Initial Catalog=RBACESSO_V100;" & _
'                 "Data Source=msb-4"
' Me.BackColor = &HFF&
' sBancoRM = "Provider=SQLOLEDB.1;" & _
'                 "Password=sa;" & _
'                 "Persist Security Info=True;" & _
'                 "User ID=sa;" & _
'                 "Initial Catalog=BkpRM;" & _
'                 "Data Source=msb-2"


Rem *************************  A T E N Ç Ã O *****************************************
Rem *************************  BASE TESTE ORACLE *****************************************
Rem **********************************************************************************

 sBancoRodbel = "Driver={Microsoft ODBC for Oracle};Server=XE;uid=Default_Acesso;pwd=Default;"
 
 Me.Caption = &H8000000A
 
 sBancoRM = "Driver={Microsoft ODBC for Oracle};Server=ERP05;uid=rm;pwd=rm;"

Rem **********************************************************************************
Rem *************************  A T E N Ç Ã O *****************************************
Rem *************************  A T E N Ç Ã O *****************************************


End Sub

Private Sub Atualizacoes_Novatos()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(DT_Pesquisa.Value, "dd/mm/yyyy"))

Rem VERSAO SQL sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Funcionarios_Novatos(sBancoRM, sBancoRodbel, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Desligados()
Dim nx As Double
Dim nLinhas As Double
Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(DT_Pesquisa.Value, "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Funcionarios_Desligados(sBancoRM, sBancoRodbel, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Ferias()

On Error GoTo Erro

Set rs = New ADODB.Recordset

Set rs = CCTempneRegBanco.Funcionarios_Ferias(sBancoRM, sBancoRodbel, CDate(Format(Now(), "dd/mm/yyyy")))

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Secoes()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(DT_Pesquisa.Value, "dd/mm/yyyy"))

'sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Funcionario_Historico_Secao(sBancoRM, sBancoRodbel, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Afastamentos()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(Now(), "dd/mm/yyyy"))

sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Funcionarios_Afastado(sBancoRM, sBancoRodbel, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub
Private Sub Atualizacoes_Atestados()

Dim sDataAdm As String

On Error GoTo Erro

Set rs = New ADODB.Recordset

sDataAdm = CDate(Format(Now(), "dd/mm/yyyy"))

'sDataAdm = "'" & Mid$(sDataAdm, 7, 4) & Mid$(sDataAdm, 4, 2) & Mid$(sDataAdm, 1, 2) & "'"

Set rs = CCTempneRegBanco.Funcionarios_Atestado(sBancoRM, sBancoRodbel, sDataAdm)

Exit Sub

Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
End Sub

Public Function CCTempneRegBanco() As neRegBanco
     Set CCTempneRegBanco = New neRegBanco
End Function

Private Sub txt_Dias_Antecedencia_Change()
If Val(Me.txt_Dias_Antecedencia.Text) > 0 Then
   Me.DT_Pesquisa.Value = CDate(Format(Now(), "dd/mm/yyyy")) - Val(Me.txt_Dias_Antecedencia.Text)
Else
   Me.DT_Pesquisa.Value = CDate(Format(Now(), "dd/mm/yyyy"))
End If

End Sub

Private Sub txt_senha_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   
Rem   If UCase(Me.txt_senha.Text) <> Format(Now(), "YYYYMMDD") & "ACS" Then
   If UCase(Me.txt_senha.Text) <> Format(Now(), "YYYY") Then
      MsgBox "SENHA NÃO CONFERE, TENTE NOVAMENTE OU TECLE <ESC> "
      Me.txt_senha.Text = ""
      Me.txt_senha.SetFocus
   Else
      MsgBox "Você terá 10 (Dez) minutos para atualização dos parametros"
      Me.frm_principal.Enabled = True
      Me.txt_senha.Text = ""
      Me.frm_principal.Enabled = True
      Me.FrmAcesso.Top = 5630
      Me.FrmAcesso.Left = 210
      nMinuto = 600
   End If
End If

If KeyAscii = 27 Then
   Me.frm_principal.Enabled = False
   Me.FrmAcesso.Top = 5630
   Me.FrmAcesso.Left = 210
   Me.txt_senha.Text = ""
   Me.frm_principal.Enabled = False
End If
   

End Sub



Function Ajusta_Acesso()
Dim ADOConnectRm As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim rRs As ADODB.Recordset 'record set que pesquisara se existe o funcionario no Rodbel
Dim ssql As String

'On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set cConect = New daAbertura

Rem ssql = "Driver={Microsoft ODBC for Oracle};Server=XE;uid=Default_Acesso;pwd=Default;"
ssql = "Driver={Microsoft ODBC for Oracle};Server=ERP05;uid=rm;pwd=rm;"

Set ADOConnectRm = cConect.Coneccao_OLEDB(ssql, "A")


ssql = "SELECT " & _
       "CONVERT(BIGINT,PFUNC.CHAPA) AS CHAPA, " & _
       "PFUNC.NOME, " & _
       "PFUNC.INTEGRGERENCIAL, " & _
       "PPESSOA.CPF, " & _
       "PPESSOA.CARTIDENTIDADE, " & _
       "PFUNC.CODFUNCAO, " & _
       "PFUNCAO.CARGO "

ssql = ssql & _
       "FROM PFUNC " & _
       "INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO " & _
       "INNER JOIN PFUNCAO ON PFUNC.CODFUNCAO = PFUNCAO.CODIGO AND PFUNC.CODCOLIGADA = PFUNCAO.CODCOLIGADA"

ssql = ssql & _
       " WHERE PFUNC.CODCOLIGADA = 1"


Set rRs = New ADODB.Recordset

ADOConnectRm.CursorLocation = adUseClientBatch

rRs.Open ssql, ADOConnectRm

Set rRs = Nothing

End Function

