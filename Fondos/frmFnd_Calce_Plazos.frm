VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFnd_Calce_Plazos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fondos: Calce de Plazos"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodos 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   9975
      _Version        =   1572864
      _ExtentX        =   17590
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   615
         Left            =   8040
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   16777215
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
         Picture         =   "frmFnd_Calce_Plazos.frx":0000
      End
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   4815
      _Version        =   1572864
      _ExtentX        =   8493
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Informe de Calce de Plazos"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmFnd_Calce_Plazos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vScroll As Boolean


Private Sub btnReporte_Click()
    Call sbReporte
End Sub

Private Sub sbReporte()

On Error GoTo vError

If cboPeriodos.Text = "" Then Exit Sub

Me.MousePointer = vbHourglass


With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Fondos: Informe de Calce de Plazos"
    
    .Connect = glogon.ConectRPT
    .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Calce_Plazos.rpt")
    .Formulas(0) = "SUBTITULO='PERIODO: " & UCase(cboPeriodos.Text) & "'"
    .Formulas(1) = "FECHA='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "EMPRESA='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"

    strSQL = "select * from fnd_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
    Call OpenRecordSet(rs, strSQL)
    
    .StoredProcParam(0) = rs!Anio
    .StoredProcParam(1) = rs!Mes
  
    .SubreportToChange = "sbResumen"
    .SelectionFormula = "{FND_CALCE_PLAZOS_RSM.ANIO} = " & rs!Anio & " AND {FND_CALCE_PLAZOS_RSM.MES} = " & rs!Mes
    rs.Close
    
    .Action = 1
End With



Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()

vModulo = 18 'Fondo de Inversion

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


strSQL = "select * from fnd_per_historico order by anio desc,mes desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboPeriodos.AddItem rs!Anio & "-" & rs!Mes
 cboPeriodos.ItemData(cboPeriodos.ListCount - 1) = CStr(rs!id_per_historico)
 rs.MoveNext
Loop
If rs.RecordCount > 1 Then
  rs.MoveFirst
  cboPeriodos.Text = rs!Anio & "-" & rs!Mes
End If
rs.Close


End Sub
