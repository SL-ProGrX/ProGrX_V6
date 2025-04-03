VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCR_Pagaré_Mail 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Confección de Pagarés"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnPrueba 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
      _Version        =   1310722
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Prueba"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
      _Version        =   1310722
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      Text            =   "17912"
   End
   Begin VB.Timer TimerX 
      Left            =   9000
      Top             =   240
   End
End
Attribute VB_Name = "frmCR_Pagaré_Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub btnPrueba_Click()
Call sbPagare_PDF(txtOperacion.Text)
End Sub

Private Sub Form_Load()


'Timer 65000

End Sub





Private Sub sbPagare_PDF(pOperacion As Long)

Dim vCedula As String, vSecciones As String

Dim pProvincia As String, pCedulaSwitch As Integer, pReemplazar As Integer
Dim pRuta As String, pReporte As String


On Error GoTo vError

'Inicializa
vSecciones = "0" 'El Pagaré utiliza las secciones

pProvincia = "Heredia"
pCedulaSwitch = 1
pReemplazar = 0

pReporte = App.Path & "\Credito_Pagarev2.rpt"

pRuta = "C:\Pagares\ASODXC_Pagare_" & pOperacion & ".pdf"

strSQL = "exec spCrd_Operacion_Pagare_Registra " & pOperacion & "," & pReemplazar _
       & "," & pCedulaSwitch & ",'" & pProvincia & "'"
Call OpenRecordSet(rs, strSQL)
  vCedula = Trim(rs!Cedula)
  vSecciones = Trim(rs!Secciones)
rs.Close

''Imprimir el Reporte
'With frmContenedor.Crt
'  .Reset
'  .WindowShowExportBtn = True
'  .WindowShowPrintSetupBtn = True
'
'  .Connect = glogon.ConectRPT
'
'  .WindowState = crptMaximized
'  .WindowTitle = "Emisión del Pagaré"
'  .ReportFileName = pReporte
'  .Formulas(0) = "fxCedula = '" & vCedula & "'"
'  .Formulas(1) = "fxSecciones = '" & vSecciones & "'"
'  .Formulas(2) = "fxBarras = '*" & pOperacion & "*'"
'
'   .Destination = crptToPrinter
'   .PrinterName = "Microsoft Print to PDF"
'   .PrinterPort = "Ne04:"
'   .PrinterDriver = "winspool"
'
'   .DialogParentHandle
'
'   .PrintFileName = pRuta
'   .Destination = crptToFile
'
'
'  .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & pOperacion
'
'  .PrintReport
'End With


Dim oXApp As CRAXDRT.Application
Dim oXRpt As CRAXDRT.Report
Dim oXOpt As CRAXDRT.ExportOptions

Dim objTable As CRAXDRT.DatabaseTable

pReporte = App.Path & "\Credito_Pagare.rpt"

pRuta = "C:\Pagares\ASODXC_Pagare_" & pOperacion & ".pdf"
Set oXApp = CreateObject("CrystalRuntime.Application")

Set oXRpt = oXApp.OpenReport(pReporte)

oXRpt.RecordSelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & pOperacion

With oXRpt
    .EnableParameterPrompting = False
    .MorePrintEngineErrorMessages = True
    
    .Database.LogOnServer "pdsodbc.dll", "PGX_Core.dsn", glogon.BaseDatos, glogon.Core_User, glogon.Core_Key
    
    .FormulaFields.Item(1).Text = "1111"
    
        For Each objTable In .Database.Tables
            objTable.SetLogOnInfo glogon.Servidor, glogon.BaseDatos, glogon.Core_User, glogon.Core_Key

        Next
    
    '"pdsodbc.dll", "PGX_Core.dsn", glogon.BaseDatos, glogon.Core_User, glogon.Core_Key
'         .Database.LogOnServer "pdsodbc.dll", "PGX_Core.dsn", glogon.BaseDatos, glogon.Core_User, glogon.Core_Key

'        For Each objTable In .Database.Tables
''            objTable.SetLogOnInfo glogon.Servidor, glogon.BaseDatos, glogon.Core_User, glogon.Core_Key
'
'        Next

End With

Set oXOpt = oXRpt.ExportOptions

With oXOpt
    .DestinationType = crEDTDiskFile
    .DiskFileName = pRuta
    .FormatType = crEFTPortableDocFormat
    .PDFExportAllPages = True
End With

oXRpt.Export False  'throws missing or out-of-date dll error




'    Dim objReport As CRAXDRT.Report
'    Dim objExportOptions As CRAXDRT.ExportOptions
'    Dim objTable As CRAXDRT.DatabaseTable
'
'    Set objApp = New CRAXDRT.Application
'
'    Set objReport = objApp.OpenReport(pReporte)
'
'    With objReport
'
'         .Database.LogOnServer "pdsodbc.dll", "Pagare"", glogon.BaseDatos, glogon.Core_User, glogon.Core_Key"
'
'        For Each objTable In .Database.Tables
''            objTable.SetLogOnInfo glogon.Servidor, glogon.BaseDatos, glogon.Core_User, glogon.Core_Key
'
'        Next
'
'
'        .RecordSelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & pOperacion
'
'        Set objExportOptions = .ExportOptions
'
'        With objExportOptions
'            .DestinationType = crEDTDiskFile
'            .DiskFileName = pRuta
'            .FormatType = crEFTPortableDocFormat
'            .PDFExportAllPages = True
'        End With
'
'        .ReadRecords
'        .DisplayProgressDialog = False
'        .Export False
'
'    End With
'
'    Set objTable = Nothing
'    Set objExportOptions = Nothing
'    Set objReport = Nothing
'    Set objApp = Nothing




'
'Dim crxApplication As New CRAXDRT.Application
'Dim Report As CRAXDRT.Report 'Object
'Set Report = crxApplication.OpenReport(pReporte)
'
'
''crxApplication.LogOnServer "p2sodbc.dll", glogon.Servidor, glogon.BaseDatos, glogon.Core_User, glogon.Core_Key
''crxApplication.LogOnServer "pdsodbc.dll", "PGX_Core.dsn", glogon.BaseDatos, glogon.Core_User, glogon.Core_Key
'
''crxApplication.LogOnServer glogon.ConectRPT
'
'
'
'With Report
''  .Connect = glogon.ConectRPT
''pdssql.dll
'
'  .FormulaFields(1) = "fxCedula = '" & vCedula & "'"
'  .FormulaFields(2) = "fxSecciones = '" & vSecciones & "'"
'  .FormulaFields(3) = "fxBarras = '*" & pOperacion & "*'"
'
'  .RecordSelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & pOperacion
'
'' Report.Database.Tables.Item(0).SetLogOnInfo glogon.ConectRPT
'  .Database.Tables(1).SetLogOnInfo glogon.Servidor, glogon.BaseDatos, glogon.Core_User, glogon.Core_Key
'  .Database.Tables(2).SetLogOnInfo glogon.Servidor, glogon.BaseDatos, glogon.Core_User, glogon.Core_Key
'  .Database.Tables(3).SetLogOnInfo glogon.Servidor, glogon.BaseDatos, glogon.Core_User, glogon.Core_Key
'
''Report.Database.Tables.Item(0).SetLogOnInfo glogon.ConectRPT
'
''    CrxReport.Database.Tables.Item(1).SetLogOnInfo "DNS_NAME", "DB_NAME",
''"userID", "password"
'
''crxApplication.LogOnServer "pdsodbc.dll", "PGX_Core.dsn", glogon.BaseDatos, glogon.Core_User, glogon.Core_Key
'
'End With
'
'
'
'With Report.ExportOptions
'    .DestinationType = crEDTDiskFile
'    .FormatType = crEFTPortableDocFormat
'    .PDFExportAllPages = True
'    .DiskFileName = pRuta
'End With
'
'Report.Export False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

