VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmInvTransacReportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes de Movimientos a Inventarios"
   ClientHeight    =   4680
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8724
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8724
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.CheckBox chkTransac 
      Height          =   252
      Left            =   6720
      TabIndex        =   13
      Top             =   1320
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   492
      Left            =   6720
      TabIndex        =   2
      Top             =   3960
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Reporte"
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
      Appearance      =   14
      Picture         =   "frmInvTransacReportes.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cboCausa 
      Height          =   312
      Left            =   1920
      TabIndex        =   5
      Top             =   3000
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8065
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1920
      TabIndex        =   6
      Top             =   1320
      Width           =   4572
      _Version        =   1245187
      _ExtentX        =   8065
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   1920
      TabIndex        =   7
      Top             =   2400
      Width           =   2292
      _Version        =   1245187
      _ExtentX        =   4043
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboBase 
      Height          =   312
      Left            =   4200
      TabIndex        =   8
      Top             =   2400
      Width           =   2292
      _Version        =   1245187
      _ExtentX        =   4043
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboUser 
      Height          =   312
      Left            =   1920
      TabIndex        =   9
      Top             =   2040
      Width           =   2292
      _Version        =   1245187
      _ExtentX        =   4043
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuarios 
      Height          =   315
      Left            =   4200
      TabIndex        =   10
      Top             =   2040
      Width           =   2292
      _Version        =   1245187
      _ExtentX        =   4043
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
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   3000
      TabIndex        =   11
      Top             =   1680
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   550
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
      Height          =   312
      Left            =   5280
      TabIndex        =   12
      Top             =   1680
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   252
      Left            =   6720
      TabIndex        =   14
      Top             =   1680
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox chkUsuarios 
      Height          =   252
      Left            =   6720
      TabIndex        =   15
      Top             =   2040
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
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
      Appearance      =   2
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkCausa 
      Height          =   252
      Left            =   6720
      TabIndex        =   4
      Top             =   3000
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
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
      Appearance      =   2
      Value           =   1
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   4
      Left            =   720
      TabIndex        =   20
      Top             =   3000
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Causa"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   3
      Left            =   720
      TabIndex        =   19
      Top             =   2400
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   2
      Left            =   720
      TabIndex        =   18
      Top             =   2040
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Usuario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   1
      Left            =   720
      TabIndex        =   17
      Top             =   1680
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fechas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   0
      Left            =   720
      TabIndex        =   16
      Top             =   1320
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Transacción"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   852
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   6852
      _Version        =   1245187
      _ExtentX        =   12086
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Boletas de Transacciones de Inventarios"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   312
      Index           =   3
      Left            =   4200
      TabIndex        =   1
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   312
      Index           =   2
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmInvTransacReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnReporte_Click()
Call sbReporte
End Sub

Private Sub cbo_Click()
If chkCausa.Value = vbUnchecked Then
   Call chkCausa_Click
End If
End Sub

Private Sub chkCausa_Click()
If chkCausa.Value = vbChecked Then
   cboCausa.Enabled = False
Else
   cboCausa.Enabled = True
   Call sbInvESCombo(Mid(cbo.Text, 1, 1), cboCausa)
End If
End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled
End Sub

Private Sub chkTransac_Click()
If chkTransac.Value = vbChecked Then
   cbo.Enabled = False
   chkCausa.Value = vbChecked
   chkCausa.Enabled = False
Else
   cbo.Enabled = True
   chkCausa.Enabled = True
End If

Call chkCausa_Click

End Sub

Private Sub chkUsuarios_Click()
If chkUsuarios.Value = vbChecked Then
   txtUsuarios.Enabled = False
Else
   txtUsuarios.Enabled = True
End If
cboUser.Enabled = txtUsuarios.Enabled

End Sub

Private Sub sbReporte()
Dim vSubTitulo As String, strSQL As String

Me.MousePointer = vbHourglass

strSQL = ""
vSubTitulo = ""

If chkTransac.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{PV_INVTRANSAC.TIPO} = '" & Mid(cbo.Text, 1, 1) & "'"
   vSubTitulo = vSubTitulo & "/MOV:" & UCase(cbo.Text) & " "
Else
   vSubTitulo = vSubTitulo & "/MOV:TODOS "
End If

If Mid(cboEstado.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{PV_INVTRANSAC.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
End If
vSubTitulo = vSubTitulo & "/EST:" & UCase(cboEstado.Text) & " "



If chkCausa.Value = vbChecked Then
  vSubTitulo = vSubTitulo & "/CAUSA: TODAS "
Else
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{PV_INVTRANSAC.COD_ENTSAL} = '" & cboCausa.ItemData(cboCausa.ListIndex) & "'"
   vSubTitulo = vSubTitulo & "/CAUSA: " & cboCausa.ItemData(cboCausa.ListIndex) & " "
End If


If chkFechas.Value = vbUnchecked Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  Select Case Mid(cboBase.Text, 7, 1)
    Case "S" 'Solicitud
       strSQL = strSQL & "CDATE({PV_INVTRANSAC.GENERA_FECHA}) in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
              & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    Case "R" 'Resuelto
       strSQL = strSQL & "CDATE({PV_INVTRANSAC.AUTORIZA_FECHA}) in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
              & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    Case "P" 'Procesado
       strSQL = strSQL & "CDATE({PV_INVTRANSAC.PROCESA_FECHA}) in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
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
       strSQL = strSQL & "{PV_INVTRANSAC.GENERA_USER} = '" & txtUsuarios.Text & "'"
    Case "R" 'Resuelto
       strSQL = strSQL & "{PV_INVTRANSAC.AUTORIZA_USER} = '" & txtUsuarios.Text & "'"
    Case "P" 'Procesado
       strSQL = strSQL & "{PV_INVTRANSAC.PROCESA_USER} = '" & txtUsuarios.Text & "'"
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
 .Formulas(3) = "fxTitulo = 'TRANSACCIONES DE INVENTARIOS'"
 .Formulas(4) = "fxSubTitulo = '" & vSubTitulo & "'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("Inventario_Transac.rpt")
 .SelectionFormula = strSQL
 
 .PrintReport

End With

Me.MousePointer = vbDefault



End Sub

Private Sub Form_Load()


Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

cbo.AddItem "Entradas"
cbo.AddItem "Salidas"
cbo.AddItem "Traslados"
cbo.AddItem "Requisiciones"
cbo.Text = "Entradas"

cboCausa.Enabled = False


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

txtUsuarios.Enabled = False
cboUser.Enabled = False

cboEstado.AddItem "Solicitados"
cboEstado.AddItem "Autorizadas"
cboEstado.AddItem "Rechazadas"
cboEstado.AddItem "Procesadas"
cboEstado.AddItem "TODOS"
cboEstado.Text = "Solicitados"


cboUser.AddItem "Solicitado por"
cboUser.AddItem "Resuelto por"
cboUser.AddItem "Procesado por"
cboUser.Text = "Solicitado por"

cboBase.AddItem "Fecha Solicitud"
cboBase.AddItem "Fecha Resolucion"
cboBase.AddItem "Fecha Procesamiento"
cboBase.Text = "Fecha Solicitud"

End Sub


Private Sub txtUsuarios_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    txtUsuarios = ""
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Consulta = "select nombre,descripcion from usuarios"
    gBusquedas.Filtro = ""
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    frmBusquedas.Show vbModal
    txtUsuarios = gBusquedas.Resultado
End If

End Sub
