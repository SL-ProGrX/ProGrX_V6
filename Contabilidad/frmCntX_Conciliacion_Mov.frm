VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCntX_Conciliacion_Mov 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Conciliacion de Movimiento"
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15780
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   15780
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4335
      Left            =   0
      TabIndex        =   13
      Top             =   5880
      Width           =   15735
      _Version        =   1310723
      _ExtentX        =   27755
      _ExtentY        =   7646
      _StockProps     =   77
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
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ListView lswC 
      Height          =   3615
      Left            =   7920
      TabIndex        =   12
      Top             =   1800
      Width           =   7815
      _Version        =   1310723
      _ExtentX        =   13785
      _ExtentY        =   6376
      _StockProps     =   77
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
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ListView lswD 
      Height          =   3615
      Left            =   0
      TabIndex        =   11
      Top             =   1800
      Width           =   7815
      _Version        =   1310723
      _ExtentX        =   13785
      _ExtentY        =   6376
      _StockProps     =   77
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
      Appearance      =   17
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      _Version        =   1310723
      _ExtentX        =   27543
      _ExtentY        =   2355
      _StockProps     =   79
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
      BorderStyle     =   1
      Begin VB.Timer TimerX 
         Interval        =   10
         Left            =   240
         Top             =   240
      End
      Begin XtremeSuiteControls.ProgressBar PrgBar 
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   8175
         _Version        =   1310723
         _ExtentX        =   14420
         _ExtentY        =   450
         _StockProps     =   93
         Value           =   10
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuenta 
         Height          =   330
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   2175
         _Version        =   1310723
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuenta_Desc 
         Height          =   330
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   6015
         _Version        =   1310723
         _ExtentX        =   10610
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnConciliar 
         Height          =   615
         Left            =   12000
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2566
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Conciliar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCntX_Conciliacion_Mov.frx":0000
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   330
         Index           =   0
         Left            =   9000
         TabIndex        =   5
         Top             =   480
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   330
         Index           =   1
         Left            =   10440
         TabIndex        =   9
         Top             =   480
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   615
         Left            =   13440
         TabIndex        =   18
         Top             =   360
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2566
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Exportar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCntX_Conciliacion_Mov.frx":0719
      End
      Begin XtremeSuiteControls.ComboBox cboExport 
         Height          =   315
         Left            =   13440
         TabIndex        =   19
         Top             =   960
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label lblProceso 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando...."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   9240
         TabIndex        =   17
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   10440
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   9000
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption scConciliados 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   5520
      Width           =   15735
      _Version        =   1310723
      _ExtentX        =   27755
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Movimientos Conciliados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scCreditos 
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   1440
      Width           =   7815
      _Version        =   1310723
      _ExtentX        =   13785
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Créditos Pendientes de Conciliar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scDebitos 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   1440
      Width           =   7815
      _Version        =   1310723
      _ExtentX        =   13785
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Débitos Pendientes de Conciliar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmCntX_Conciliacion_Mov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub btnConciliar_Click()
Dim pCuenta As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lswD.ListItems.Clear
lswC.ListItems.Clear
lsw.ListItems.Clear

pCuenta = fxgCntCuentaFormato(False, txtCuenta, 0)

PrgBar.Visible = True
lblProceso.Visible = True

lblProceso.Caption = "Inicializando...."
DoEvents

strSQL = "exec spCntX_Concilia_Inicializa '" & glogon.Usuario & "', " & gCntX_Parametros.CodigoConta _
        & ", '" & pCuenta & "','" & Format(dtpFecha(0).Value, "yyyy-mm-dd") _
        & "','" & Format(dtpFecha(1).Value, "yyyy-mm-dd") & " 23:59:59'"

Call ConectionExecute(strSQL)

lblProceso.Caption = "Procesando...."
DoEvents

strSQL = "exec spCntX_Concilia_Procesa '" & glogon.Usuario & "', " & gCntX_Parametros.CodigoConta _
        & ", '" & pCuenta & "',1"
Call OpenRecordSet(rs, strSQL)
Do While rs!Pendientes > 0
    PrgBar.Max = rs!Total
    PrgBar.Value = rs!Total - rs!Pendientes
    lblProceso.Caption = "Procesando [" & rs!Pendientes & "] de " & rs!Total
    DoEvents
    
    strSQL = "exec spCntX_Concilia_Procesa '" & glogon.Usuario & "', " & gCntX_Parametros.CodigoConta _
            & ", '" & pCuenta & "',20"
    Call OpenRecordSet(rs, strSQL)
Loop

lblProceso.Visible = False
PrgBar.Visible = False

lblProceso.Caption = "Montando Resultados...."
DoEvents

'pCntX_Concilia_Resultados(@Usuario varchar(30), @Contabilidad int, @Cuenta varchar(60), @Tipo varchar(10) = 'CON')

'Debitos
strSQL = "exec spCntX_Concilia_Resultados '" & glogon.Usuario & "', " & gCntX_Parametros.CodigoConta _
        & ", '" & pCuenta & "','DB'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswD.ListItems.Add(, , rs!Num_Linea)
     itmX.SubItems(1) = rs!Tipo_Asiento
     itmX.SubItems(2) = rs!Num_Asiento
     itmX.SubItems(3) = Format(rs!fecha, "yyyy-mm-dd")
     itmX.SubItems(4) = Format(rs!Monto, "Standard")
     itmX.SubItems(5) = rs!Tipo_Cambio
     itmX.SubItems(6) = rs!Referencia & ""
     itmX.SubItems(7) = rs!Documento & ""
     itmX.SubItems(8) = rs!Detalle & ""
 rs.MoveNext
Loop

'Creditos
strSQL = "exec spCntX_Concilia_Resultados '" & glogon.Usuario & "', " & gCntX_Parametros.CodigoConta _
        & ", '" & pCuenta & "','CR'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswC.ListItems.Add(, , rs!Num_Linea)
     itmX.SubItems(1) = rs!Tipo_Asiento
     itmX.SubItems(2) = rs!Num_Asiento
     itmX.SubItems(3) = Format(rs!fecha, "yyyy-mm-dd")
     itmX.SubItems(4) = Format(rs!Monto, "Standard")
     itmX.SubItems(5) = rs!Tipo_Cambio
     itmX.SubItems(6) = rs!Referencia & ""
     itmX.SubItems(7) = rs!Documento & ""
     itmX.SubItems(8) = rs!Detalle & ""
 rs.MoveNext
Loop

'Conciliados
strSQL = "exec spCntX_Concilia_Resultados '" & glogon.Usuario & "', " & gCntX_Parametros.CodigoConta _
        & ", '" & pCuenta & "','CON'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Num_Linea)
     itmX.SubItems(1) = rs!Tipo_Asiento
     itmX.SubItems(2) = rs!Num_Asiento
     itmX.SubItems(3) = Format(rs!fecha, "yyyy-mm-dd")
     itmX.SubItems(4) = Format(rs!Monto, "Standard")
     itmX.SubItems(5) = rs!Tipo_Cambio
     itmX.SubItems(6) = rs!Referencia & ""
     itmX.SubItems(7) = rs!Documento & ""
     itmX.SubItems(8) = rs!Detalle & ""
     
     
     itmX.SubItems(9) = rs!CR_Tipo_Asiento
     itmX.SubItems(10) = rs!CR_Num_Asiento
     itmX.SubItems(11) = Format(rs!CR_fecha, "yyyy-mm-dd")
     itmX.SubItems(12) = Format(rs!CR_Monto, "Standard")
     itmX.SubItems(13) = rs!CR_Tipo_Cambio
     itmX.SubItems(14) = rs!CR_Referencia & ""
     itmX.SubItems(15) = rs!CR_Documento & ""
     itmX.SubItems(16) = rs!CR_Detalle & ""
     
 rs.MoveNext
Loop



Me.MousePointer = vbDefault

MsgBox "Proceso Concluido, verifique resultados!", vbInformation
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExport_Click()

On Error GoTo vError

lblProceso.Visible = True
lblProceso.Caption = "Exportando..."
DoEvents

Me.MousePointer = vbHourglass

PrgBar.Visible = True

Select Case cboExport.Text
    Case "Conciliados"
        Call Excel_Exportar_Lsw(lsw, PrgBar)
    Case "Débitos"
        Call Excel_Exportar_Lsw(lswD, PrgBar)
    Case "Crébitos"
        Call Excel_Exportar_Lsw(lswC, PrgBar)
End Select

PrgBar.Visible = False
lblProceso.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

cboExport.Clear
cboExport.AddItem "Conciliados"
cboExport.AddItem "Débitos"
cboExport.AddItem "Crébitos"
cboExport.Text = "Conciliados"


With lswD.ColumnHeaders
    .Clear
    .Add , , "Linea Id", 1200
    .Add , , "T.Asiento", 1000, vbCenter
    .Add , , "N.Asiento", 2200
    .Add , , "Fecha", 1200, vbCenter
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "T.C.", 1200, vbRightJustify
    .Add , , "Referencia", 2200
    .Add , , "Documento", 2200
    .Add , , "Detalle", 2200
End With

With lswC.ColumnHeaders
    .Clear
    .Add , , "Linea Id", 1200
    .Add , , "T.Asiento", 1000, vbCenter
    .Add , , "N.Asiento", 2200
    .Add , , "Fecha", 1200, vbCenter
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "T.C.", 1200, vbRightJustify
    .Add , , "Referencia", 2200
    .Add , , "Documento", 2200
    .Add , , "Detalle", 2200
End With


With lsw.ColumnHeaders
    .Clear
    .Add , , "Linea Id", 1200
    .Add , , "T.Asiento", 1000, vbCenter
    .Add , , "N.Asiento", 2200
    .Add , , "Fecha", 1200, vbCenter
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "T.C.", 1200, vbRightJustify
    .Add , , "Referencia", 2200
    .Add , , "Documento", 2200
    .Add , , "Detalle", 2200

    .Add , , "Mt. T.Asiento", 1000, vbCenter
    .Add , , "Mt. N.Asiento", 2200
    .Add , , "Mt. Fecha", 1200, vbCenter
    .Add , , "Mt. Monto", 1800, vbRightJustify
    .Add , , "Mt. T.C.", 1200, vbRightJustify
    .Add , , "Mt. Referencia", 2200
    .Add , , "Mt. Documento", 2200
    .Add , , "Mt. Detalle", 2200

End With



End Sub

Private Sub Form_Resize()
Dim pAlto As Long

On Error Resume Next

pAlto = (Me.Height - (lswD.Top + scConciliados.Height + 750)) / 2

lswD.Height = pAlto
lswC.Height = pAlto
lsw.Height = pAlto

lswD.Width = (Me.Width - 300) / 2
lswC.Width = lswD.Width
lswC.Left = lswD.Width + 100

scDebitos.Width = lswD.Width

scCreditos.Left = lswC.Left
scCreditos.Width = lswC.Width

lsw.Width = Me.Width - 200
scConciliados.Width = lsw.Width
scConciliados.Top = lswD.Top + lswD.Height + 100
lsw.Top = scConciliados.Top + scConciliados.Height


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

dtpFecha(0).Value = fxFechaServidor
dtpFecha(1).Value = dtpFecha(0).Value

PrgBar.Visible = False
lblProceso.Visible = False

Form_Resize

End Sub



Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
     frmCntX_ConsultaCuentas.Show vbModal
     txtCuenta.Text = gCuenta
     txtCuenta_Desc.SetFocus
End If

End Sub

Private Sub txtCuenta_LostFocus()
 txtCuenta_Desc.Text = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, txtCuenta.Text))
 txtCuenta.Text = fxCntX_CuentaFormato(True, txtCuenta.Text)
End Sub
