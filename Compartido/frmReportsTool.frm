VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmReportsTool 
   Caption         =   "Informes"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11685
   Icon            =   "frmReportsTool.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRV 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11295
      lastProp        =   500
      _cx             =   19923
      _cy             =   9763
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmReportsTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    CRV.Top = 0
    CRV.Left = 0
    CRV.Height = ScaleHeight
    CRV.Width = ScaleWidth
End Sub


Public Sub PrintReport()
    'Print the Report
    CRV.PrintReport

End Sub

Public Sub ViewReport()
    'View the Report
    CRV.ViewReport

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim i As Integer
'
''Call ReportsTool.Reset
''Set crReport = Nothing
 
 
' Connect = glogon.ConectRPT
'
' For i = 0 To 20
'   Formulas(i) = ""
' Next i
'
' For i = 0 To 20
'   Parametros(i) = ""
' Next i
End Sub
