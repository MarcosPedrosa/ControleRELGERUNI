VERSION 5.00
Begin VB.Form frmGIFDisUsuarioBanco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuário do Sistema"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cbo_Usuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "frmGIFDisUsuarioBanco.frx":0000
      Left            =   5040
      List            =   "frmGIFDisUsuarioBanco.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox TXTSENHA 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   570
      IMEMode         =   3  'DISABLE
      Left            =   5040
      MaxLength       =   9
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2790
      Width           =   3255
   End
   Begin VB.ComboBox CboCliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "frmGIFDisUsuarioBanco.frx":0004
      Left            =   5040
      List            =   "frmGIFDisUsuarioBanco.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   570
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1650
      Left            =   1500
      OLEDropMode     =   1  'Manual
      Picture         =   "frmGIFDisUsuarioBanco.frx":0008
      ScaleHeight     =   412.222
      ScaleMode       =   0  'User
      ScaleWidth      =   2025
      TabIndex        =   3
      Top             =   750
      Width           =   2085
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "        MUSASHI DO BRASIL"
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
      Height          =   225
      Left            =   1140
      TabIndex        =   9
      Top             =   420
      Width           =   3225
   End
   Begin VB.Label LBLSENHA 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SENHA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5040
      TabIndex        =   8
      Top             =   2250
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5040
      TabIndex        =   7
      Top             =   1140
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5040
      TabIndex        =   6
      Top             =   30
      Width           =   3255
   End
   Begin VB.Label lbl_release 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000002&
      Caption         =   "Release "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   3360
      Width           =   5055
   End
   Begin VB.Label lbl_Fone 
      BackStyle       =   0  'Transparent
      Caption         =   "        MUSASHI INFORMATICA"
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
      Height          =   225
      Left            =   5070
      TabIndex        =   4
      Top             =   3390
      Width           =   3225
   End
End
Attribute VB_Name = "frmGIFDisUsuarioBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bAchou As Boolean
Public sChave As String

Rem dados do banco principal
Public bNomeBd As Collection
Public bNomeIP As Collection
Public bNomeUsu As Collection
Public bNomeSen As Collection

Rem dados dos bancos da rm e rodbel
Public bNomeBd2 As Collection
Public bNomeIP2 As Collection
Public bNomeUsu2 As Collection
Public bNomeSen2 As Collection
Public nxv As Integer
Private Sub Cbo_Usuario_Change()
If Len(sNomeBanco) > 0 Then
   Me.TXTSENHA.Enabled = True
End If
End Sub

Private Sub Cbo_Usuario_Click()
If Len(sNomeBanco) > 0 Then
   Me.TXTSENHA.Enabled = True
End If
End Sub

Private Sub Cbo_Usuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Len(sNomeBanco) > 0 Then
      Me.TXTSENHA.Enabled = True
      Me.TXTSENHA.SetFocus
   End If
End If

End Sub

Private Sub CboCliente_Change()

Call Muda_Banco

If Len(sNomeBanco) > 0 Then Call Carrega_Usuario

End Sub

Private Sub CboCliente_Click()

Call Muda_Banco

If Len(sNomeBanco) > 0 Then Carrega_Usuario

End Sub



Private Sub Form_Load()
Dim Nada As String
Dim nx As Integer
Dim nRotina As Integer
Dim ny As Integer
Dim Nz As Integer

On Error GoTo Erro

Me.MousePointer = vbHourglass

sNomeBanco = ""
sBancoRM = ""
Nada = App.Path & "\LOCALIZA.TXT"
If Dir$(Nada) = "" Then
   MsgBox "Arquivo de inicialização não encontrado, Procure o responsável!", 16, "Programa Cancelado"
   End
End If
Open Nada For Random Access Read Write Shared As #1 Len = 89
Set bNomeBd = New Collection
Set bNomeIP = New Collection
Set bNomeUsu = New Collection
Set bNomeSen = New Collection

Set bNomeBd2 = New Collection
Set bNomeIP2 = New Collection
Set bNomeUsu2 = New Collection
Set bNomeSen2 = New Collection

For nx = 1 To 100
    Get 1, nx, ARQUIVO_TEXTO
    If Asc(Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 1, 1))) = 13 _
       Or Asc(Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 1, 1))) = 0 Then
       nx = 100
    Else
       If Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 1, 1)) <> "#" Then
          If Mid$(ARQUIVO_TEXTO.LinhaTexto, 1, 1) = "0" Then
             ny = ny + 1
             CboCliente.AddItem Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 19, 12))
             CboCliente.ItemData(ny - 1) = Val(Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 16, 3)))
             bNomeBd.Add Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 2, 14))
             bNomeIP.Add Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 31, 15))
             bNomeUsu.Add Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 46, 15))
             bNomeSen.Add Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 61, 15))
          Else
             bNomeBd2.Add Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 2, 14))
             bNomeIP2.Add Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 31, 15))
             bNomeUsu2.Add Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 46, 15))
             bNomeSen2.Add Trim(Mid$(ARQUIVO_TEXTO.LinhaTexto, 61, 15))
          End If
       End If
       
     End If
Next
Me.CboCliente.Visible = True
Me.lbl_release.Caption = "Release " & App.Major & "." & Format(App.Revision, "000")

Close #1

Me.MousePointer = vbDefault
Exit Sub

Erro:
MsgBox "Existe um erro na linha " & Str(nx) & ", do arquivo LOCALIZA.TXT.", , Me.Caption
Me.MousePointer = vbDefault
End

'App.Path & "\help\Aeshelp.hlp"
'   Me.CboCliente.ListIndex = 0
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   bAchou = False
   TXTSENHA.Text = ""
   Me.Hide

End Sub





Private Sub TXTSENHA_KeyPress(KeyAscii As Integer)

sUsuario = ""
If KeyAscii = 13 Then
   If Confirmar_Senha Then
      sUsuario = Format(Me.Cbo_Usuario.ItemData(Cbo_Usuario.ListIndex), "000")
'      bNomeUsuario = Me.cbo_usuario.List(cbo_usuario.ListIndex)
      bAchou = True
      Me.Hide
      Exit Sub
   End If
   nxv = nxv + 1
   If nxv > 3 And sUsuario = "" Then
      MsgBox "Número de tentativas excedeu o limite!"
      Me.Hide
      bAchou = False
      Exit Sub
   End If
End If

If KeyAscii = 27 Then
   bAchou = False
   TXTSENHA.Text = ""
   Me.Hide
End If

End Sub
Private Function Confirmar_Senha() As Boolean
Dim nx As Integer
Dim cRec As ADODB.Recordset

On Error GoTo Erro
Me.MousePointer = vbHourglass

If Me.Cbo_Usuario.ListCount = 0 Then Carrega_Usuario

Set cRec = New ADODB.Recordset

If Me.Cbo_Usuario.ListIndex = -1 Then
   Exit Function
End If

Set cRec = CCTempneUsuario.USUARIO_Consultar(sNomeBanco, Me.Cbo_Usuario.ItemData(Me.Cbo_Usuario.ListIndex))

bAchou = False
Confirmar_Senha = False
'        If UnCripta(cRec!USU_SENHA) = Me.TXTSENHA.Text Then

If cRec.RecordCount > 0 Then
    cRec.MoveFirst
    For nx = 0 To cRec.RecordCount - 1
        If UnCripta(Trim(cRec!USU_SENHA)) = Trim(Me.TXTSENHA.Text) Then
           sUsuario = Format(Me.Cbo_Usuario.ItemData(Me.Cbo_Usuario.ListIndex), "000")
           sNome_Usuario = cRec!USU_USUARIO
           bAchou = True
           Confirmar_Senha = True
           Exit For
        End If
        cRec.MoveNext
    Next
End If
Me.MousePointer = vbDefault
Set cRec = Nothing
If nxv < 4 Then
   If sNome_Usuario = "" Or bAchou = False Then
      MsgBox "Senha Inválida, Tente novamente!"
      Me.TXTSENHA.Text = ""
      Me.TXTSENHA.SetFocus
   End If
End If

Exit Function

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Function


Public Function Carrega_Usuario()
Dim cRecAux As ADODB.Recordset
Dim nx As Integer

On Error GoTo Erro
Me.MousePointer = vbHourglass
Me.Cbo_Usuario.Clear

If Not ADOConnection Is Nothing Then
   ADOConnection.CommitTrans
   ADOConnection.Close
End If

'MsgBox "VOU LER OS CLIENTES" & sNomeBanco
Set cRecAux = CCTempneUsuario.USUARIO_Consultar(sNomeBanco)
'MsgBox "LI OS CLIENTES" & sNomeBanco

If cRecAux Is Nothing Then
   MsgBox "Não Existem Usuarios cadastrados!"
   Exit Function
   
ElseIf cRecAux.RecordCount = 0 Then
   MsgBox "Não Existem Usuarios cadastrados ou habilitados!"
   Exit Function
End If

cRecAux.MoveFirst

If cRecAux.RecordCount > 0 Then
   Cbo_Usuario.Clear
   For nx = 1 To cRecAux.RecordCount
       Cbo_Usuario.AddItem cRecAux!USU_USUARIO
       Cbo_Usuario.ItemData(nx - 1) = cRecAux!USU_CODIGO
       cRecAux.MoveNext
   Next
Else
   MsgBox "Não Existem usuarios cadastrados!"
End If

Me.MousePointer = vbDefault
Set cRecAux = Nothing

Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Function

Private Function Muda_Banco()
If Me.CboCliente.ListIndex = -1 Then Exit Function

sCodEempresa = Format(Me.CboCliente.ItemData(Me.CboCliente.ListIndex), "000")
sNomeEmpresa = Me.CboCliente.List(Me.CboCliente.ListIndex)

sNomeBanco = "Provider=SQLOLEDB.1;" & _
             "Persist Security Info=True;" & _
             "Data Source=" & bNomeIP(Me.CboCliente.ListIndex + 1) & ";" & _
             "Initial Catalog=" & bNomeBd(Me.CboCliente.ListIndex + 1) & ";" & _
             "User ID=" & bNomeUsu(Me.CboCliente.ListIndex + 1) & ";" & _
             "Password=" & bNomeSen(Me.CboCliente.ListIndex + 1) & ";"


If (Me.CboCliente.ListIndex + 1) = 1 Then
   sBancoRM = "Provider=SQLOLEDB.1;" & _
              "Persist Security Info=True;" & _
              "Data Source=" & bNomeIP2(1) & ";" & _
              "Initial Catalog=" & bNomeBd2(1) & ";" & _
              "User ID=" & bNomeUsu2(1) & ";" & _
              "Password=" & bNomeSen2(1) & ";"
              
   sBancoRodbel = "Provider=SQLOLEDB.1;" & _
                  "Persist Security Info=True;" & _
                  "Data Source=" & bNomeIP2(2) & ";" & _
                  "Initial Catalog=" & bNomeBd2(2) & ";" & _
                  "User ID=" & bNomeUsu2(2) & ";" & _
                  "Password=" & bNomeSen2(2) & ";"

   sBancoUnimed = "Provider=SQLOLEDB.1;" & _
                  "Persist Security Info=True;" & _
                  "Data Source=" & bNomeIP2(3) & ";" & _
                  "Initial Catalog=" & bNomeBd2(3) & ";" & _
                  "User ID=" & bNomeUsu2(3) & ";" & _
                  "Password=" & bNomeSen2(3) & ";"
Else
   sBancoRM = "Provider=SQLOLEDB.1;" & _
              "Persist Security Info=True;" & _
              "Data Source=" & bNomeIP2(4) & ";" & _
              "Initial Catalog=" & bNomeBd2(4) & ";" & _
              "User ID=" & bNomeUsu2(4) & ";" & _
              "Password=" & bNomeSen2(4) & ";"
              
   sBancoRodbel = "Provider=SQLOLEDB.1;" & _
                  "Persist Security Info=True;" & _
                  "Data Source=" & bNomeIP2(5) & ";" & _
                  "Initial Catalog=" & bNomeBd2(5) & ";" & _
                  "User ID=" & bNomeUsu2(5) & ";" & _
                  "Password=" & bNomeSen2(5) & ";"
   
   sBancoUnimed = "Provider=SQLOLEDB.1;" & _
                  "Persist Security Info=True;" & _
                  "Data Source=" & bNomeIP2(6) & ";" & _
                  "Initial Catalog=" & bNomeBd2(6) & ";" & _
                  "User ID=" & bNomeUsu2(6) & ";" & _
                  "Password=" & bNomeSen2(6) & ";"
End If

'sNumIP & ";" & 10.3.0.3;
'sDataBase teklogix
'User ID etiquetas
'Password=etiquetas

End Function
