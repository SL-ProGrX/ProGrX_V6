VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCprReportesGenerales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Compras"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   10350
   Begin XtremeSuiteControls.FlatEdit txtProveedor 
      Height          =   315
      Left            =   3120
      TabIndex        =   19
      ToolTipText     =   "Presione F4 Para Consultar"
      Top             =   3360
      Width           =   5055
      _Version        =   1441792
      _ExtentX        =   8916
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbInforme 
      Height          =   1095
      Left            =   0
      TabIndex        =   10
      Top             =   5040
      Width           =   10335
      _Version        =   1441792
      _ExtentX        =   18230
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   615
         Left            =   8280
         TabIndex        =   11
         Top             =   360
         Width           =   1695
         _Version        =   1441792
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "&Reporte"
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
         Picture         =   "frmCprReportesGenerales.frx":0000
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.CheckBox chkGrpProveedor 
         Height          =   435
         Left            =   6000
         TabIndex        =   13
         Top             =   360
         Width           =   1695
         _Version        =   1441792
         _ExtentX        =   2990
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Agrupado por Proveedor"
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
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   3120
      TabIndex        =   14
      Top             =   1680
      Width           =   5055
      _Version        =   1441792
      _ExtentX        =   8916
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   3120
      TabIndex        =   15
      Top             =   2760
      Width           =   2535
      _Version        =   1441792
      _ExtentX        =   4471
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
   Begin XtremeSuiteControls.ComboBox cboBase 
      Height          =   330
      Left            =   5640
      TabIndex        =   16
      Top             =   2760
      Width           =   2535
      _Version        =   1441792
      _ExtentX        =   4471
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
   Begin XtremeSuiteControls.ComboBox cboUser 
      Height          =   330
      Left            =   3120
      TabIndex        =   17
      Top             =   2400
      Width           =   2535
      _Version        =   1441792
      _ExtentX        =   4471
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   3120
      TabIndex        =   18
      Top             =   3720
      Width           =   5055
      _Version        =   1441792
      _ExtentX        =   8916
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
   Begin XtremeSuiteControls.ComboBox cboForma 
      Height          =   330
      Left            =   3120
      TabIndex        =   20
      Top             =   4080
      Width           =   5055
      _Version        =   1441792
      _ExtentX        =   8916
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
   Begin XtremeSuiteControls.ComboBox cboCxP 
      Height          =   330
      Left            =   3120
      TabIndex        =   21
      Top             =   4440
      Width           =   5055
      _Version        =   1441792
      _ExtentX        =   8916
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   3960
      TabIndex        =   22
      Top             =   2040
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   6480
      TabIndex        =   23
      Top             =   2040
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtUsuarios 
      Height          =   315
      Left            =   5640
      TabIndex        =   24
      ToolTipText     =   "Presione F4 Para Consultar"
      Top             =   2400
      Width           =   2535
      _Version        =   1441792
      _ExtentX        =   4471
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   315
      Left            =   8400
      TabIndex        =   25
      Top             =   2040
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todas"
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.CheckBox chkUsuarios 
      Height          =   315
      Left            =   8400
      TabIndex        =   26
      Top             =   2400
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkProveedores 
      Height          =   315
      Left            =   8400
      TabIndex        =   27
      Top             =   3360
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkTipo 
      Height          =   315
      Left            =   8400
      TabIndex        =   28
      Top             =   3720
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Appearance      =   16
      Value           =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Compras"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   10
      Left            =   1800
      TabIndex        =   12
      Top             =   360
      Width           =   7212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CxP Programa"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1080
      TabIndex        =   9
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1080
      TabIndex        =   8
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Compra"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1080
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1080
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transacción"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   3120
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   5520
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmCprReportesGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cbo_Click()

cboEstado.Clear
cboUser.Clear
cboBase.Clear
cboCxP.Enabled = False
cboForma.Enabled = False
txtProveedor.Enabled = False
chkProveedores.Enabled = False
chkGrpProveedor.Enabled = False

If Mid(cbo.Text, 1, 1) = "C" Then
    cboEstado.AddItem "Procesadas"
    cboEstado.AddItem "Anuladas"
    cboEstado.AddItem "Devoluciones"
    cboEstado.AddItem "TODOS"
    cboEstado.Text = "TODOS"
    
    cboUser.AddItem "Procesada Por"
    cboUser.AddItem "Anulada Por"
    cboUser.Text = "Procesada Por"
    
    cboBase.AddItem "Fecha Procesa"
    cboBase.AddItem "Fecha Anula"
    cboBase.Text = "Fecha Procesa"
    
    txtProveedor.Enabled = True
    chkProveedores.Enabled = True
    
    cboForma.Enabled = True
    
    chkGrpProveedor.Enabled = True
    cboCxP.Enabled = True
    
Else
    cboEstado.AddItem "Solicitados"
    cboEstado.AddItem "Autorizadas"
    cboEstado.AddItem "Rechazadas"
    cboEstado.AddItem "Pendientes Despachos"
    cboEstado.AddItem "TODOS"
    cboEstado.Text = "TODOS"
    
    cboUser.AddItem "Solicitado por"
    cboUser.AddItem "Resuelto por"
    cboUser.Text = "Solicitado por"
    
    cboBase.AddItem "Fecha Solicitud"
    cboBase.AddItem "Fecha Resolucion"
    cboBase.Text = "Fecha Solicitud"

End If

End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled
cboBase.Enabled = dtpInicio.Enabled
End Sub


Private Sub chkProveedores_Click()
If chkProveedores.Value = vbChecked Then
   txtProveedor.Enabled = False
Else
   txtProveedor.Enabled = True
End If
End Sub

Private Sub chkTipo_Click()
If chkTipo.Value = vbChecked Then
   cboTipo.Enabled = False
Else
   cboTipo.Enabled = True
End If
End Sub

Private Sub chkUsuarios_Click()
If chkUsuarios.Value = vbChecked Then
   txtUsuarios.Enabled = False
Else
   txtUsuarios.Enabled = True
End If
cboUser.Enabled = txtUsuarios.Enabled

End Sub

Private Sub sbReporteOrdenes()
Dim vSubTitulo As String, strSQL As String

Me.MousePointer = vbHourglass

strSQL = ""
vSubTitulo = ""

If Mid(cboTipo.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{CPR_ORDENES.TIPO_ORDEN} = '" & fxCodigoCbo(cboTipo) & "'"
   vSubTitulo = vSubTitulo & "/TIPO:" & fxCodigoCbo(cboTipo) & " "
Else
   vSubTitulo = vSubTitulo & "/TIPO:TODOS "
End If

If Mid(cboEstado.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{CPR_ORDENES.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
End If
vSubTitulo = vSubTitulo & "/EST:" & UCase(cboEstado.Text) & " "

If chkFechas.Value = vbUnchecked Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  Select Case Mid(cboBase.Text, 1, 1)
    Case "S" 'Solicitud
       strSQL = strSQL & "CDATE({CPR_ORDENES.GENERA_FECHA}) in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
              & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    Case "R" 'Resuelto
       strSQL = strSQL & "CDATE({CPR_ORDENES.AUTORIZA_FECHA}) in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
              & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
  End Select
   vSubTitulo = vSubTitulo & "/" & UCase(cboBase.Text) & " I:" & Format(dtpInicio.Value, "dd/mm/yyyy") & " C:" & Format(dtpCorte.Value, "dd/mm/yyyy") & " "
Else
   vSubTitulo = vSubTitulo & "/TODAS LAS FECHAS "
End If

If chkUsuarios.Value = vbUnchecked Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  Select Case Mid(cboUser.Text, 1, 1)
    Case "S" 'Solicitud
       strSQL = strSQL & "{CPR_ORDENES.GENERA_USER} = '" & txtUsuarios.Text & "'"
    Case "R" 'Resuelto
       strSQL = strSQL & "{CPR_ORDENES.AUTORIZA_USER} = '" & txtUsuarios.Text & "'"
  End Select
  vSubTitulo = vSubTitulo & "/" & UCase(cboUser.Text) & " " & txtUsuarios.Text
Else
  vSubTitulo = vSubTitulo & "/TODOS LOS USUARIOS "
End If


With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes SIF - Division Comercial"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "fxTitulo = 'ORDENES DE COMPRAS'"
 .Formulas(4) = "fxSubTitulo = '" & vSubTitulo & "'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("Compras_OrdenesListado.rpt")
 .SelectionFormula = strSQL
 
 .PrintReport

End With

Me.MousePointer = vbDefault

End Sub


Private Sub sbReporteCompras()
Dim vSubTitulo As String, strSQL As String

Me.MousePointer = vbHourglass

strSQL = ""
vSubTitulo = ""



If Mid(cboTipo.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{CPR_ORDENES.TIPO_ORDEN} = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
   vSubTitulo = vSubTitulo & "¦ TIPO:" & cboTipo.ItemData(cboTipo.ListIndex) & " "
Else
   vSubTitulo = vSubTitulo & "¦ TIPO:TODOS "
End If


If Mid(cboEstado.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{CPR_COMPRAS.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
   vSubTitulo = vSubTitulo & "¦ ESTADO:" & UCase(cboEstado.Text) & " "
Else
   vSubTitulo = vSubTitulo & "¦ ESTADO:TODOS "
End If


If chkFechas.Value = vbUnchecked Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   If Mid(cboBase.Text, 7, 1) = "P" Then
        strSQL = strSQL & "CDATE({CPR_COMPRAS.GENERA_FECHA}) in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
               & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
   Else
        strSQL = strSQL & "CDATE({CPR_COMPRAS.ANULA_FECHA}) in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
               & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
   End If
   vSubTitulo = vSubTitulo & "¦  FECHA I:" & Format(dtpInicio.Value, "dd¦ mm¦ yyyy") & " C:" & Format(dtpCorte.Value, "dd¦ mm¦ yyyy") & " "
Else
   vSubTitulo = vSubTitulo & "¦ TODAS LAS FECHAS "
End If

If chkUsuarios.Value = vbUnchecked Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  If cboUser.Text = "Procesada Por" Then
      strSQL = strSQL & "{CPR_COMPRAS.GENERA_USER} = '" & txtUsuarios.Text & "'"
  Else
   'Anulada
      strSQL = strSQL & "{CPR_COMPRAS.ANULA_USER} = '" & txtUsuarios.Text & "'"
  End If
  vSubTitulo = vSubTitulo & "¦ " & UCase(cboUser.Text) & " " & txtUsuarios.Text
Else
  vSubTitulo = vSubTitulo & "¦ TODOS LOS USUARIOS "
End If


If Mid(cboForma.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{CPR_COMPRAS.FORMA_PAGO} = '" & UCase(Mid(cboForma.Text, 1, 2)) & "'"
   vSubTitulo = vSubTitulo & "¦ FP:" & cboForma.Text & " "
Else
   vSubTitulo = vSubTitulo & "¦ FP:TODAS "
End If

If Mid(cboCxP.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{CPR_COMPRAS.CXP_ESTADO} = '" & IIf((Mid(cboCxP.Text, 1, 1) = "P"), "P", "G") & "'"
   vSubTitulo = vSubTitulo & "¦ CxP:" & cboCxP.Text & " "
Else
   vSubTitulo = vSubTitulo & "¦ CxP:TODAS "
End If

If chkProveedores.Value = vbUnchecked And txtProveedor.Tag <> "" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{CPR_COMPRAS.COD_PROVEEDOR} = " & txtProveedor.Tag
   vSubTitulo = vSubTitulo & "¦ PROV:" & txtProveedor.Tag & " "
Else
   vSubTitulo = vSubTitulo & "¦ PROV:TODOS "
End If

With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes SIF - Division Comercial"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd¦ mm¦ yyyy") & "'"
 .Formulas(3) = "fxTitulo = 'COMPRAS'"
 .Formulas(4) = "fxSubTitulo = '" & vSubTitulo & "'"
 
 If chkGrpProveedor.Value = vbChecked Then
    .ReportFileName = SIFGlobal.fxPathReportes("Compras_ComprasListadoG.rpt")
 Else
    .ReportFileName = SIFGlobal.fxPathReportes("Compras_ComprasListado.rpt")
 End If
 .SelectionFormula = strSQL
 
 .PrintReport

End With

Me.MousePointer = vbDefault


End Sub

Private Sub cmdReporte_Click()
If Mid(cbo.Text, 1, 1) = "C" Then
  Call sbReporteCompras
Else
  Call sbReporteOrdenes
End If
End Sub

Private Sub Form_Activate()
vModulo = 35
End Sub

Private Sub Form_Load()
vModulo = 35

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

Call sbCprCboTiposOrden(cboTipo)
cboTipo.Enabled = False

cboForma.Clear
cboForma.AddItem "Contado"
cboForma.AddItem "Crédito"
cboForma.AddItem "TODAS"
cboForma.Text = "TODAS"

cbo.AddItem "Compras"
cbo.AddItem "Ordenes de Compras"
cbo.Text = "Compras"

cboCxP.Clear
cboCxP.AddItem "Pendiente de Programación"
cboCxP.AddItem "Programación Generada"
cboCxP.AddItem "TODAS"
cboCxP.Text = "TODAS"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

txtUsuarios.Enabled = False
cboUser.Enabled = False

End Sub


Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Col1Name = "Id Interno"
  gBusquedas.Col2Name = "Proveedor"
  gBusquedas.Col3Name = "Identificación"
  
  gBusquedas.Consulta = "select cod_proveedor,descripcion, cedJur from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtProveedor.Tag = gBusquedas.Resultado
  txtProveedor = gBusquedas.Resultado2
End If
End Sub

Private Sub txtUsuarios_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    txtUsuarios = ""
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    
  gBusquedas.Col1Name = "Usuario"
  gBusquedas.Col2Name = "Nombre"
  
    gBusquedas.Consulta = "select nombre,descripcion from usuarios"
    gBusquedas.Filtro = ""
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    txtUsuarios = gBusquedas.Resultado
End If

End Sub


