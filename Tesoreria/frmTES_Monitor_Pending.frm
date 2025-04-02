VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmTES_Monitor_Pending 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Monitor de Procesos"
   ClientHeight    =   7932
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7548
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7932
   ScaleWidth      =   7548
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6732
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   7332
      _Version        =   1245185
      _ExtentX        =   12933
      _ExtentY        =   11874
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   5400
      Top             =   480
   End
   Begin XtremeSuiteControls.PushButton btnMonitor 
      Height          =   372
      Index           =   0
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1332
      _Version        =   1245185
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Refrescar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmTES_Monitor_Pending.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtCasos 
      Height          =   312
      Left            =   4320
      TabIndex        =   4
      Top             =   7560
      Width           =   972
      _Version        =   1245185
      _ExtentX        =   1714
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotal 
      Height          =   312
      Left            =   5280
      TabIndex        =   5
      Top             =   7560
      Width           =   2172
      _Version        =   1245185
      _ExtentX        =   3831
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   7560
      Width           =   972
      _Version        =   1245185
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Totales:"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   612
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5292
      _Version        =   1245185
      _ExtentX        =   9334
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Monitor de Transacciones pendientes de envío a Bancos para su desembolso"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTES_Monitor_Pending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnMonitor_Click(Index As Integer)

Select Case Index
    Case 0
      Call sbRefresca
    Case 1
      Unload Me
End Select

End Sub

Private Sub Form_Load()

vModulo = 9
Me.BackColor = RGB(70, 111, 178)

With lsw.ColumnHeaders
    .Clear
    .Add , , "Sistema:", 3500
    .Add , , "Casos", 1100, vbRightJustify
    .Add , , "Monto", 2600, vbRightJustify
End With

End Sub


Private Sub sbRefresca()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lCasos As Long, curTotal As Currency
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spTes_Monitor_Pending"
Call OpenRecordSet(rs, strSQL)

vPaso = True

curTotal = 0
lCasos = 0

lsw.ListItems.Clear

Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!Modulo_Desc)
        itmX.Tag = rs!Modulo
        itmX.SubItems(1) = rs!Casos
        itmX.SubItems(2) = Format(rs!Monto, "Standard")
  lCasos = lCasos + rs!Casos
  curTotal = curTotal + rs!Monto
  
  rs.MoveNext
Loop
rs.Close

txtCasos.Text = CStr(lCasos)
txtTotal.Text = Format(curTotal, "Standard")

vPaso = False
Me.MousePointer = vbDefault


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

Select Case Item.Tag
    Case "BEN"
        Call sbClassCall("BENEFICIOS", 7, "frmAF_BeneficiosTraslado")
    Case "AFI"
        Call sbFormsCall("frmAF_LiquidacionAsientos")
    Case "CRD"
        Call sbClassCall("PROCESOS", 4, "frmCR_TraspasoTesoreria")
    Case "FND"
        Call sbFormsCall("frmFNDTraspasoTesoreria")
    Case "CxC"
        Call sbFormsCall("frmCxC_RemesasTesoreria")
    Case "CxP"
        Call sbFormsCall("frmCxPControlEjecucion")
End Select

Me.MousePointer = vbDefault

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbRefresca

End Sub
