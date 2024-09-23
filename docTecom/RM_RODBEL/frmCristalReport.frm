VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmCristalReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cristal"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   12270
   StartUpPosition =   2  'CenterScreen
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   6495
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   12015
      lastProp        =   500
      _cx             =   21193
      _cy             =   11456
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   -1  'True
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmCristalReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Rrec As ADODB.Recordset
Public nTipoTela As Integer

Private Sub Form_Load()
'  Screen.MousePointer = vbHourglass
'
''  Set report = app1.OpenReport(App.Path & "crptVale03p.rpt")
''  report.Database.SetDataSource Rrec
''
''  CRViewer91.ReportSource = report
''  CRViewer91.ViewReport
''  Me.Top = 0
''  Me.Left = 0
'
'  Screen.MousePointer = vbDefault
  
End Sub
Private Sub Form_Activate()
  Me.CRViewer91.Refresh
End Sub

