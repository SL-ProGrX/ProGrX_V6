VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmPosReportesGenerales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pos: Reportes Generales"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   11010
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4215
      Left            =   0
      TabIndex        =   19
      Top             =   1800
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
      _ExtentY        =   7435
      _StockProps     =   77
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox fraReporte 
      Height          =   1095
      Left            =   4560
      TabIndex        =   1
      Top             =   4920
      Width           =   6375
      _Version        =   1441793
      _ExtentX        =   11245
      _ExtentY        =   1931
      _StockProps     =   79
      ForeColor       =   16711680
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   615
         Left            =   4800
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   1085
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmPosReportesGenerales.frx":0000
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fechas"
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboCobro 
      Height          =   315
      Left            =   6000
      TabIndex        =   4
      Top             =   3960
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8493
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   6000
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   7200
      TabIndex        =   6
      Top             =   3600
      Width           =   3615
      _Version        =   1441793
      _ExtentX        =   6371
      _ExtentY        =   550
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
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   6000
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
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
      Left            =   7320
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.CheckBox chkCajas 
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cajas"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkAgente 
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Agente"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFormaPago 
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   3240
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Pago"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkCliente 
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cliente"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkTipoVenta 
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   3960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo Venta"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.ComboBox cboFormaPago 
      Height          =   315
      Left            =   6000
      TabIndex        =   14
      Top             =   3240
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8493
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
   Begin XtremeSuiteControls.ComboBox cboAgente 
      Height          =   315
      Left            =   6000
      TabIndex        =   15
      Top             =   2880
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8493
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
   Begin XtremeSuiteControls.ComboBox cboCajas 
      Height          =   315
      Left            =   6000
      TabIndex        =   16
      Top             =   2520
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8493
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   1320
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Informes disponibles:"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption lblReporte 
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   1320
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "[Seleccione un Informe]"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Informe de POS"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   4932
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   11892
   End
End
Attribute VB_Name = "frmPosReportesGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSubTitulo As String

Private Sub chkAgente_Click()
If chkAgente.Value = vbChecked Then
   cboAgente.Enabled = False
Else
   cboAgente.Enabled = True
End If
End Sub

Private Sub chkCajas_Click()
If chkCajas.Value = vbChecked Then
  cboCajas.Enabled = False
Else
  cboCajas.Enabled = True
End If
End Sub

Private Sub chkCliente_Click()
If chkCliente.Value = vbChecked Then
  txtCedula.Enabled = False
  txtNombre.Enabled = False
Else
  txtCedula.Enabled = True
  txtNombre.Enabled = True
End If
End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If
End Sub

Private Function fxRepFechas()
Dim vFecha As String

vFecha = "in date(" & Format(dtpInicio, "yyyy,mm,dd") & ") to date(" & Format(dtpCorte, "yyyy,mm,dd") & ")"

fxRepFechas = vFecha

End Function

Private Function fxSQL(Optional i As Integer = 0) As String
Dim vSQL As String

vSQL = ""
vSubTitulo = "" 'Variable de Formulario

Select Case i
   Case 0, 10 'Reporte de Ventas x Facturacion, y a Credito
        vSQL = "{PV_FACTURACION.ESTADO} = 'P'"
       If chkFechas.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
          vSQL = vSQL & "CDATE({PV_FACTURACION.FECHA}) " & fxRepFechas
          vSubTitulo = "Inicio " & Format(dtpInicio.Value, "yyyy/mm/dd") & " Corte " & Format(dtpCorte.Value, "yyyy/mm/dd")
       End If
       If chkCajas.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
         vSQL = vSQL & "{PV_FACTURACION.COD_CAJA} = '" & cboCajas.ItemData(cboCajas.ListIndex) & "'"
         vSubTitulo = vSubTitulo & "¦ Caja: " & cboCajas.Text
       End If
       If chkAgente.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
         vSQL = vSQL & "{PV_FACTURACION.COD_AGENTE} = '" & cboAgente.ItemData(cboAgente.ListIndex) & "'"
         vSubTitulo = vSubTitulo & "¦ Agente: " & cboAgente.Text
       End If
       
       
       
'       If chkFormaPago.Value = vbUnchecked Then
'         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
'         vSQL = vSQL & "{PV_FACTURACION.COD_FORMA_PAGO} = " & fxCodigoCbo(cboFormaPago)
'         vSubTitulo = vSubTitulo & "¦ Pago: " & cboFormaPago.Text
'       End If
       
       If chkCliente.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
         vSQL = vSQL & "{PV_FACTURACION.CEDULA} = '" & txtCedula & "'"
         vSubTitulo = vSubTitulo & "¦ Cliente: " & txtCedula
       End If
       
       If chkTipoVenta.Value = vbUnchecked Then
        If Mid(cboCobro.Text, 1, 1) <> "T" Then
            If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
            Select Case cboCobro.Text
              Case "Contado"
                vSQL = vSQL & "{PV_FACTURACION.CXC_TIPO} = 'CNT'"
              Case "Credito"
                vSQL = vSQL & "{PV_FACTURACION.CXC_TIPO} = 'CRD'"
            End Select
            vSubTitulo = vSubTitulo & "¦ Tipo de Venta: " & cboCobro.Text
        End If
       End If
       
       
   
   Case 1 'Devoluciones de Facturas
       If chkFechas.Value = vbUnchecked Then
          vSQL = "CDATE({PV_DEVOLUCIONES.FECHA}) " & fxRepFechas
          vSubTitulo = "Inicio " & Format(dtpInicio.Value, "yyyy/mm/dd") & " Corte " & Format(dtpCorte.Value, "yyyy/mm/dd")
       End If
       If chkCajas.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
         vSQL = vSQL & "{PV_DEVOLUCIONES.COD_CAJA} = '" & cboCajas.ItemData(cboCajas.ListIndex) & "'"
         vSubTitulo = vSubTitulo & "¦ Caja " & cboCajas.ItemData(cboCajas.ListIndex)
       End If
       If chkAgente.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
         vSQL = vSQL & "{PV_FACTURACION.COD_AGENTE} = '" & cboAgente.ItemData(cboAgente.ListIndex) & "'"
         vSubTitulo = vSubTitulo & "¦ Agente " & cboAgente.ItemData(cboAgente.ListIndex)
       End If
       
'       If chkFormaPago.Value = vbUnchecked Then
'         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
'         vSQL = vSQL & "{PV_FACTURACION.COD_FORMA_PAGO} = " & fxCodigoCbo(cboFormaPago)
'         vSubTitulo = vSubTitulo & " Pago " & cboFormaPago.Text
'       End If
       
       If chkCliente.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
         vSQL = vSQL & "{PV_FACTURACION.CEDULA} = '" & txtCedula & "'"
         vSubTitulo = vSubTitulo & "¦ Cliente " & txtCedula
       End If
   
       If chkTipoVenta.Value = vbUnchecked Then
        If Mid(cboCobro.Text, 1, 1) <> "T" Then
            If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
            Select Case cboCobro.Text
              Case "Contado"
                vSQL = vSQL & "{PV_FACTURACION.CXC_TIPO} = 'CNT'"
              Case "Credito"
                vSQL = vSQL & "{PV_FACTURACION.CXC_TIPO} = 'CRD'"
            End Select
            vSubTitulo = vSubTitulo & "¦ Tipo de Venta: " & cboCobro.Text
        End If
       End If
   Case 2 'Anulacion de Facturas
   
       vSQL = vSQL & "{PV_FACTURACION.ESTADO} = 'A'"
       
       If chkFechas.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
          vSQL = vSQL & "CDATE({PV_FACTURACION.ANU_FECHA}) " & fxRepFechas
          vSubTitulo = "Inicio " & Format(dtpInicio.Value, "yyyy/mm/dd") & " Corte " & Format(dtpCorte.Value, "yyyy/mm/dd")
       End If
       
       If chkCajas.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
         vSQL = vSQL & "{PV_FACTURACION.ANU_CAJACOD} = '" & cboCajas.ItemData(cboCajas.ListIndex) & "'"
         vSubTitulo = vSubTitulo & "¦ Caja " & cboCajas.ItemData(cboCajas.ListIndex)
       End If
       
       If chkAgente.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
         vSQL = vSQL & "{PV_FACTURACION.COD_AGENTE} = '" & cboAgente.ItemData(cboAgente.ListIndex) & "'"
         vSubTitulo = vSubTitulo & "¦ Agente " & cboAgente.ItemData(cboAgente.ListIndex)
       End If
       
'       If chkFormaPago.Value = vbUnchecked Then
'         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
'         vSQL = vSQL & "{PV_FACTURACION.ANU_FORMA_PAGO} = " & fxCodigoCbo(cboFormaPago)
'         vSubTitulo = vSubTitulo & " Pago " & cboFormaPago.Text
'       End If
       
       If chkCliente.Value = vbUnchecked Then
         If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
         vSQL = vSQL & "{PV_FACTURACION.CEDULA} = '" & txtCedula & "'"
         vSubTitulo = vSubTitulo & "¦ Cliente " & txtCedula
       End If
   
       If chkTipoVenta.Value = vbUnchecked Then
        If Mid(cboCobro.Text, 1, 1) <> "T" Then
            If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
            Select Case cboCobro.Text
              Case "Contado"
                vSQL = vSQL & "{PV_FACTURACION.CXC_TIPO} = 'CNT'"
              Case "Credito"
                vSQL = vSQL & "{PV_FACTURACION.CXC_TIPO} = 'CRD'"
            End Select
            vSubTitulo = vSubTitulo & "¦ Tipo de Venta: " & cboCobro.Text
        End If
       End If
   Case 3
   Case 4
   Case 5
   Case 6
   Case 7
   Case 8
   Case 9

End Select

fxSQL = vSQL

End Function


Private Sub chkFormaPago_Click()
If chkFormaPago.Value = vbChecked Then
  cboFormaPago.Enabled = False
Else
  cboFormaPago.Enabled = True
End If
End Sub

Private Sub cmdReporte_Click()

Select Case lblReporte.Tag
  Case "x00" 'Reporte de Ventas x Facturacion
     Call sbPosReportes("Facturas", lblReporte.Caption, vSubTitulo, fxSQL(0))
     
  Case "x01" 'Devoluciones de Mercaderia
     Call sbPosReportes("Devoluciones", lblReporte.Caption, vSubTitulo, fxSQL(1))
  
  Case "x02" 'Facturas Anuladas
     Call sbPosReportes("Anulaciones", lblReporte.Caption, vSubTitulo, fxSQL(2))
     
  Case "x03" 'Pedidos
     MsgBox "En Investigacion y Desarrollo...", vbInformation
  
  Case "x04" 'Cotizaciones y Proformas
     MsgBox "En Investigacion y Desarrollo...", vbInformation
  
  Case "x05" 'Saldas de Articulos x Facturacion
     Call sbPosReportes("SalidaXFactura", lblReporte.Caption, vSubTitulo, fxSQL(0))
  
  Case "x06" 'Entradas de Articulos x Facturas Anuladas
     MsgBox "En Investigacion y Desarrollo...", vbInformation
  
  Case "x07" 'Entradas de Articulos x Devoluciones
     Call sbPosReportes("EntradaXDevolucion", lblReporte.Caption, vSubTitulo, fxSQL(1))
  
  Case "x08", "x09" 'Movimientos de Inventarios x POS (Agrupado x Bodegas)
     MsgBox "En Investigacion y Desarrollo...", vbInformation
  
  Case "x10" 'Facturas a Crédito
     Call sbPosReportes("FacturasCxCI", lblReporte.Caption, vSubTitulo, fxSQL(10))
  
  Case "x11" 'Detalle de Ventas y Costos
     Call sbPosReportes("VentasCts", lblReporte.Caption, vSubTitulo, fxSQL(0))
  
  Case "x12" 'CxC en Mora x Agente
     Call sbPosReportes("CxCMoraAgente", lblReporte.Caption, vSubTitulo, fxSQL(0))
  
  
  Case "x13" 'Facturacion por Forma de Pago
     Call sbPosReportes("FacturaFP", lblReporte.Caption, vSubTitulo, fxSQL(0))
  
  
End Select

End Sub

Private Sub Form_Load()
Dim strSQL As String

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

strSQL = "select Rtrim(cod_agente) as 'IdX', rtrim(nombre) as 'itmX' from pv_agentes"
Call sbCbo_Llena_New(cboAgente, strSQL, False, True)

strSQL = "select  cod_forma_pago as 'IdX',  rtrim(descripcion) as itmX from pv_formas_pago"
Call sbCbo_Llena_New(cboFormaPago, strSQL, False, True)

strSQL = "select rtrim(cod_caja) as 'IdX', (Rtrim(cod_caja) + ' - ' + usuario) as 'ItmX'  from pv_cajas"
Call sbCbo_Llena_New(cboCajas, strSQL, False, True)

Call chkAgente_Click
Call chkCliente_Click
Call chkFormaPago_Click
Call chkCajas_Click

cboCobro.Clear
cboCobro.AddItem "Todos"
cboCobro.AddItem "Contado"
cboCobro.AddItem "Credito"
cboCobro.Text = "Todos"

With lsw.ColumnHeaders
    .Clear
    .Add , , "", lsw.Width - 100
End With

With lsw.ListItems
    .Clear
    .Add , "x00", "Reporte de Ventas"
    .Add , "x01", "Devoluciones de Mercadería"
    .Add , "x02", "Facturas Anuladas"
    .Add , "x03", "Pedidos Registrados"
    .Add , "x04", "Cotizaciones (Proformas)"
    .Add , "x05", "Salidas de Articulos x Ventas"
    .Add , "x06", "Entradas de Articulos x Fact. Anuladas"
    .Add , "x07", "Entradas de Articulos x Devoluciones"
    .Add , "x08", "Kardex del POS"
    .Add , "x09", "Kardex del POS Agrupado x Bodegas"
    .Add , "x10", "Ventas a Crédito (CxC Internas)"
    .Add , "x11", "Detalle de Ventas / Costos"
    .Add , "x12", "CxC en Mora x Agente"
    .Add , "x13", "Facturación por Forma de Pago"

    .Add , "x14", "Exportar Informe de Ventas completo"
    .Add , "x15", "Exportar Informe detalle de Ventas completo"

End With

lblReporte.Tag = "00"
lblReporte.Caption = "Reportes de Ventas"

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    lblReporte.Tag = Item.Key
    lblReporte.Caption = Item.Text
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdReporte.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If
End Sub

