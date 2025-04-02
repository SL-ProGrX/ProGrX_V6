VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmSYS_APP_LOG 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "System: Registro de Bitácoras de Consumo de los Apps"
   ClientHeight    =   6870
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton Btn_Buscar 
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   960
      Width           =   1455
      _Version        =   1441792
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Consultar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_APP_LOG.frx":0000
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   68288515
      CurrentDate     =   42949
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5175
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   10575
      _Version        =   524288
      _ExtentX        =   18653
      _ExtentY        =   9128
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   490
      ScrollBars      =   2
      SpreadDesigner  =   "frmSYS_APP_LOG.frx":0700
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rango"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "APPs: Estadísticas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   4452
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11292
   End
End
Attribute VB_Name = "frmSYS_APP_LOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub sbCargaGrid_Local(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim pCon As New ADODB.Connection, pStrCon As String
Dim pUser As String, pKey As String

Me.MousePointer = vbHourglass

On Error GoTo vError

pUser = glogon.RootName
pKey = glogon.RootKey

pStrCon = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" & glogon.Servidor _
       & ";Database=PGX_BASE;APP=PGX_Portal_Admin;tcp:" & glogon.Servidor _
       & "," & SIFGlobal.PuertosDisponibles & ";"

With pCon
  .CommandTimeout = 15
  .Mode = adModeReadWrite
  .CursorLocation = adUseClient
  
  .Open pStrCon, pUser, pKey
  .CommandTimeout = 360
End With


vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 0 'Reset
vGrid.MaxRows = 1 'Inicia
vGrid.Row = vGrid.MaxRows

rs.Open strSQL, pCon, adOpenStatic
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    vGrid.Text = CStr(rs.Fields(i - 1).Value)
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

vError:

Me.MousePointer = vbDefault


End Sub


Private Sub Btn_Buscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAPP_Estadistica " & gPortal.Empresa_Id & ",'" & Format(dtpInicio.Value, "yyyy/mm/dd") _
        & " 00:00:00','" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

Call sbCargaGrid_Local(vGrid, 5, strSQL)


Me.MousePointer = vbDefault
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 13
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -1, dtpCorte.Value)
vGrid.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

