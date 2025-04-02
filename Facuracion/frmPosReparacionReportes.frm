VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmPosReparacionReportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S.R.: Reportes"
   ClientHeight    =   6300
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   8424
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   8424
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   1680
   End
   Begin VB.ComboBox cboDetalle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ComboBox cboBoleta 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4920
      Width           =   2535
   End
   Begin VB.ComboBox cboUsuario 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4080
      Width           =   5655
   End
   Begin VB.ComboBox cboBase 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   3360
      Width           =   4095
   End
   Begin VB.TextBox txtCedula 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   3360
      Width           =   1575
   End
   Begin VB.ComboBox cboProveedor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3000
      Width           =   5655
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tiempos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.OptionButton opt 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Boletas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   2175
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9758
      _ExtentY        =   3831
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Reporte"
         Object.Width           =   8360
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   4920
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   195166211
      CurrentDate     =   37782
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   6960
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   195231747
      CurrentDate     =   37782
   End
   Begin XtremeSuiteControls.PushButton cmdReporte 
      Height          =   612
      Left            =   6720
      TabIndex        =   22
      Top             =   5520
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Reporte"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmPosReparacionReportes.frx":0000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   3
      Left            =   5880
      TabIndex        =   18
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Boleta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   2640
      TabIndex        =   17
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Left            =   960
      TabIndex        =   16
      Top             =   4080
      Width           =   1335
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   6240
      TabIndex        =   13
      Top             =   3720
      Width           =   735
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   4200
      TabIndex        =   12
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   960
      TabIndex        =   9
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Left            =   960
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
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
      Left            =   960
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   360
      X2              =   2880
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   240
      X2              =   2760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   240
      X2              =   8280
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblReporte 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Left            =   2760
      TabIndex        =   21
      Top             =   2280
      Width           =   5535
   End
End
Attribute VB_Name = "frmPosReparacionReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxRepFechas()
Dim vFecha As String

vFecha = "in date(" & Format(dtpInicio, "yyyy,mm,dd") & ") to date(" & Format(dtpCorte, "yyyy,mm,dd") & ")"

fxRepFechas = vFecha

End Function


Private Sub sbReporteBoletas()
Dim strSQL As String
Dim vTitulo As String, vSubTitulo As String


vTitulo = lblReporte.Caption
strSQL = ""

If cboUsuario.Text <> "Todos" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " and "
  strSQL = strSQL & "{POS_REPARACION.GENERA_USUARIO} = '" & cboUsuario.Text & "'"
End If
vSubTitulo = "[US.:" & cboUsuario.Text & "] "

If cboBase.Text <> "Todas" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " and "
  strSQL = strSQL & "{POS_REPARACION.GENERA_FECHA}" & fxRepFechas
  vSubTitulo = vSubTitulo & "[Fecha : i." & Format(dtpInicio.Value, "dd/mm/yyyy") _
             & " c." & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
             
Else
  vSubTitulo = vSubTitulo & "[Fecha : Todas]"
End If

If cboBoleta.Text <> "Todas" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " and "
  
  Select Case cboBoleta.Text
    Case "Solicitada"
      strSQL = strSQL & "{POS_REPARACION.ESTADO} = 'SC'"
    Case "Trasladada"
      strSQL = strSQL & "{POS_REPARACION.ESTADO} = 'TC'"
    Case "Trasladada (Parcial)"
      strSQL = strSQL & "{POS_REPARACION.ESTADO} = 'TP'"
    Case "Recibida"
      strSQL = strSQL & "{POS_REPARACION.ESTADO} = 'RC'"
    Case "Recibida (Parcial)"
      strSQL = strSQL & "{POS_REPARACION.ESTADO} = 'RP'"
    Case "Entregada"
      strSQL = strSQL & "{POS_REPARACION.ESTADO} = 'EC'"
    Case "Entrega (Parcial)"
      strSQL = strSQL & "{POS_REPARACION.ESTADO} = 'EP'"
  End Select
End If
vSubTitulo = vSubTitulo & "[Estado: " & cboBoleta.Text & "]"

If txtCedula.Text <> "" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " and "
  strSQL = strSQL & "{POS_REPARACION.CEDULA} = '" & txtCedula.Text & "'"
  vSubTitulo = vSubTitulo & "[Cliente: " & txtNombre.Text & "]"
End If



Me.MousePointer = vbHourglass

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
 .Formulas(3) = "fxTitulo = '" & UCase(vTitulo) & "'"
 .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"


 Select Case lblReporte.Tag
    Case "x00b" 'Listado de Boletas"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoBoletasGen.rpt")
    Case "x01b" 'Resumen de Boletas x Usuario"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoBoletasUser.rpt")
    Case "x02b" 'Resumen de Boletas x Cliente"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoBoletasCliente.rpt")
    Case "x03b" 'Listado de Boletas x Estado"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoBoletasEstado.rpt")
    Case "x04b" 'Resumen de Boletas x Estado"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoBoletasEstadoRsm.rpt")
 End Select

 .SelectionFormula = strSQL
 .PrintReport
End With

Me.MousePointer = vbDefault


End Sub


Private Sub sbReporteInventario()
Dim strSQL As String
Dim vTitulo As String, vSubTitulo As String


vTitulo = lblReporte.Caption
strSQL = ""



If cboDetalle.Text <> "Todos" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " and "
  strSQL = strSQL & "{vPosReparacionCiclos.ESTADO} = '" & Mid(cboDetalle.Text, 1, 1) & "'"
  vSubTitulo = vSubTitulo & "[Estado: " & cboDetalle.Text & "] [Base en Fechas y Usuarios : " & cboDetalle.Text & "] "

  Select Case Mid(cboDetalle.Text, 1, 1)
    Case "S"
        If cboUsuario.Text <> "Todos" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.USUARIO_ENTRADA} = '" & cboUsuario.Text & "'"
        End If
        vSubTitulo = "[US.:" & cboUsuario.Text & "] "
    
        If cboBase.Text <> "Todas" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.FECHA_ENTRADA}" & fxRepFechas
          vSubTitulo = vSubTitulo & "[Fecha : i." & Format(dtpInicio.Value, "dd/mm/yyyy") _
                     & " c." & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
                     
        Else
          vSubTitulo = vSubTitulo & "[Fecha : Todas]"
        End If
        
    Case "T"
        If cboUsuario.Text <> "Todos" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.USUARIO_TRASLADO} = '" & cboUsuario.Text & "'"
        End If
        vSubTitulo = "[US.:" & cboUsuario.Text & "] "
    
        If cboBase.Text <> "Todas" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.FECHA_TRASLADO}" & fxRepFechas
          vSubTitulo = vSubTitulo & "[Fecha : i." & Format(dtpInicio.Value, "dd/mm/yyyy") _
                     & " c." & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
                     
        Else
          vSubTitulo = vSubTitulo & "[Fecha : Todas]"
        End If
    
    
    Case "R"
        If cboUsuario.Text <> "Todos" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.USUARIO_RECIBE} = '" & cboUsuario.Text & "'"
        End If
        vSubTitulo = "[US.:" & cboUsuario.Text & "] "
    
        If cboBase.Text <> "Todas" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.FECHA_RECIBE}" & fxRepFechas
          vSubTitulo = vSubTitulo & "[Fecha : i." & Format(dtpInicio.Value, "dd/mm/yyyy") _
                     & " c." & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
                     
        Else
          vSubTitulo = vSubTitulo & "[Fecha : Todas]"
        End If
    
    Case "E"
        If cboUsuario.Text <> "Todos" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.USUARIO_ENTREGA} = '" & cboUsuario.Text & "'"
        End If
        vSubTitulo = "[US.:" & cboUsuario.Text & "] "
    
        If cboBase.Text <> "Todas" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.FECHA_ENTREGA}" & fxRepFechas
          vSubTitulo = vSubTitulo & "[Fecha : i." & Format(dtpInicio.Value, "dd/mm/yyyy") _
                     & " c." & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
                     
        Else
          vSubTitulo = vSubTitulo & "[Fecha : Todas]"
        End If
    
    
  End Select

Else
  vSubTitulo = vSubTitulo & "[Estado: " & cboDetalle.Text & "]"

End If

If txtCedula.Text <> "" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " and "
  strSQL = strSQL & "{POS_REPARACION.CEDULA} = '" & txtCedula.Text & "'"
  vSubTitulo = vSubTitulo & "[Cliente: " & txtNombre.Text & "]"
End If

If cboProveedor.Text <> "Todos" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " and "
  strSQL = strSQL & "{vPosReparacionCiclos.COD_PROVEEDOR} = " & cboProveedor.ItemData(cboProveedor.ListIndex)
End If
vSubTitulo = vSubTitulo & "[Proveedor: " & cboProveedor.Text & "]"



Me.MousePointer = vbHourglass

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
 .Formulas(3) = "fxTitulo = '" & UCase(vTitulo) & "'"
 .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"

Select Case lblReporte.Tag
    Case "x00i" 'Listado General"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoInventario.rpt")
    Case "x01i" 'Listado x Estado"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoInventarioEstado.rpt")
    Case "x02i" 'Listado Agrupado Proveedor-Estado"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoInventarioProvEst.rpt")
    Case "x04i" 'Listado Agrupado Estado-Proveedor"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoInventarioEstProv.rpt")
    Case "x05i" 'Listado Articulos Reenviados"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoInventarioReenvios.rpt")
    Case "x06i" 'Listado Articulos x Cliente"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoInventarioCliente.rpt")
End Select

 .SelectionFormula = strSQL
 .PrintReport
End With

Me.MousePointer = vbDefault



End Sub

Private Sub sbReporteTiempos()
Dim strSQL As String
Dim vTitulo As String, vSubTitulo As String
vTitulo = lblReporte.Caption
strSQL = ""



If cboDetalle.Text <> "Todos" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " and "
  strSQL = strSQL & "{vPosReparacionCiclos.ESTADO} = '" & Mid(cboDetalle.Text, 1, 1) & "'"
  vSubTitulo = vSubTitulo & "[Estado: " & cboDetalle.Text & "] [Base en Fechas y Usuarios : " & cboDetalle.Text & "] "

  Select Case Mid(cboDetalle.Text, 1, 1)
    Case "S"
        If cboUsuario.Text <> "Todos" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.USUARIO_ENTRADA} = '" & cboUsuario.Text & "'"
        End If
        vSubTitulo = "[US.:" & cboUsuario.Text & "] "
    
        If cboBase.Text <> "Todas" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.FECHA_ENTRADA}" & fxRepFechas
          vSubTitulo = vSubTitulo & "[Fecha : i." & Format(dtpInicio.Value, "dd/mm/yyyy") _
                     & " c." & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
                     
        Else
          vSubTitulo = vSubTitulo & "[Fecha : Todas]"
        End If
        
    Case "T"
        If cboUsuario.Text <> "Todos" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.USUARIO_TRASLADO} = '" & cboUsuario.Text & "'"
        End If
        vSubTitulo = "[US.:" & cboUsuario.Text & "] "
    
        If cboBase.Text <> "Todas" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.FECHA_TRASLADO}" & fxRepFechas
          vSubTitulo = vSubTitulo & "[Fecha : i." & Format(dtpInicio.Value, "dd/mm/yyyy") _
                     & " c." & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
                     
        Else
          vSubTitulo = vSubTitulo & "[Fecha : Todas]"
        End If
    
    
    Case "R"
        If cboUsuario.Text <> "Todos" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.USUARIO_RECIBE} = '" & cboUsuario.Text & "'"
        End If
        vSubTitulo = "[US.:" & cboUsuario.Text & "] "
    
        If cboBase.Text <> "Todas" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.FECHA_RECIBE}" & fxRepFechas
          vSubTitulo = vSubTitulo & "[Fecha : i." & Format(dtpInicio.Value, "dd/mm/yyyy") _
                     & " c." & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
                     
        Else
          vSubTitulo = vSubTitulo & "[Fecha : Todas]"
        End If
    
    Case "E"
        If cboUsuario.Text <> "Todos" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.USUARIO_ENTREGA} = '" & cboUsuario.Text & "'"
        End If
        vSubTitulo = "[US.:" & cboUsuario.Text & "] "
    
        If cboBase.Text <> "Todas" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " and "
          strSQL = strSQL & "{vPosReparacionCiclos.FECHA_ENTREGA}" & fxRepFechas
          vSubTitulo = vSubTitulo & "[Fecha : i." & Format(dtpInicio.Value, "dd/mm/yyyy") _
                     & " c." & Format(dtpCorte.Value, "dd/mm/yyyy") & "]"
                     
        Else
          vSubTitulo = vSubTitulo & "[Fecha : Todas]"
        End If
    
    
  End Select

Else
  vSubTitulo = vSubTitulo & "[Estado: " & cboDetalle.Text & "]"

End If

If txtCedula.Text <> "" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " and "
  strSQL = strSQL & "{POS_REPARACION.CEDULA} = '" & txtCedula.Text & "'"
  vSubTitulo = vSubTitulo & "[Cliente: " & txtNombre.Text & "]"
End If

If cboProveedor.Text <> "Todos" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " and "
  strSQL = strSQL & "{vPosReparacionCiclos.COD_PROVEEDOR} = " & cboProveedor.ItemData(cboProveedor.ListIndex)
End If
vSubTitulo = vSubTitulo & "[Proveedor: " & cboProveedor.Text & "]"


Me.MousePointer = vbHourglass

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
 .Formulas(3) = "fxTitulo = '" & UCase(vTitulo) & "'"
 .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"

Select Case lblReporte.Tag
    Case "x00t" 'Ciclo de Solución"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoCiclos.rpt")
    Case "x01t" 'Tiempos en Taller"
         .ReportFileName = SIFGlobal.fxPathReportes("POS_ReparacionListadoCiclosTaller.rpt")
End Select

 .SelectionFormula = strSQL
 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub


Private Sub cboBase_Click()

If Mid(cboBase.Text, 1, 1) = "T" Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If

End Sub

Private Sub cmdReporte_Click()

Select Case True
  Case opt.Item(0).Value 'Boletas
     Call sbReporteBoletas
  Case opt.Item(1).Value 'Inventario
     Call sbReporteInventario
  Case opt.Item(2).Value 'Tiempos
     Call sbReporteTiempos
End Select

End Sub

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub sbCargaCombos()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

cboProveedor.Clear
cboProveedor.AddItem "Todos"
cboProveedor.ItemData(cboProveedor.NewIndex) = 0
cboProveedor.Text = "Todos"

strSQL = "select P.cod_proveedor,P.descripcion" _
       & " from POS_REPARACION_DETALLE R inner join CXP_Proveedores P on R.cod_Proveedor = P.cod_Proveedor" _
       & " group by P.cod_proveedor,P.descripcion"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  cboProveedor.AddItem rs!Descripcion
  cboProveedor.ItemData(cboProveedor.NewIndex) = rs!cod_proveedor
 rs.MoveNext
Loop
rs.Close


cboUsuario.Clear
cboUsuario.AddItem "Todos"
cboUsuario.Text = "Todos"

strSQL = "select Usuario_Entrada as Usuario From POS_REPARACION_DETALLE" _
       & " group by Usuario_Entrada Union" _
       & " select Usuario_Traslado as Usuario From POS_REPARACION_DETALLE" _
       & " group by Usuario_Traslado Union" _
       & " select USUARIO_RECIBO as Usuario From POS_REPARACION_DETALLE" _
       & " group by USUARIO_RECIBO Union" _
       & " select USUARIO_entrega as Usuario From POS_REPARACION_DETALLE" _
       & " group by USUARIO_entrega"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboUsuario.AddItem rs!Usuario
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub


Private Sub Form_Load()


vModulo = 33

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

cboBoleta.Clear
cboBoleta.AddItem "Todas"
cboBoleta.AddItem "Solicitada"
cboBoleta.AddItem "Trasladada"
cboBoleta.AddItem "Trasladada (Parcial)"
cboBoleta.AddItem "Recibida"
cboBoleta.AddItem "Recibida (Parcial)"
cboBoleta.AddItem "Entregada"
cboBoleta.AddItem "Entrega (Parcial)"
cboBoleta.Text = "Todas"

cboDetalle.Clear
cboDetalle.AddItem "Solicitada"
cboDetalle.AddItem "Trasladada"
cboDetalle.AddItem "Recibida"
cboDetalle.AddItem "Entregada"
cboDetalle.Text = "Solicitada"

cboBase.Clear
cboBase.AddItem "Todas"
cboBase.AddItem "Rango"
cboBase.Text = "Todas"

Call opt_Click(0)

Call Formularios(Me)
Call RefrescaTags(Me)

Timer1.Interval = 20

End Sub


Private Sub lsw_Click()

lblReporte.Tag = lsw.SelectedItem.Key
lblReporte.Caption = lsw.SelectedItem

End Sub

Private Sub opt_Click(Index As Integer)

lsw.ListItems.Clear

cboProveedor.Enabled = True
cboBase.Enabled = True
cboDetalle.Enabled = True
cboBoleta.Enabled = True

Select Case Index
  Case 0 'Boletas
    lsw.ListItems.Add , "x00b", "Listado de Boletas"
    lsw.ListItems.Add , "x01b", "Resumen de Boletas x Usuario"
    lsw.ListItems.Add , "x02b", "Resumen de Boletas x Cliente"
    lsw.ListItems.Add , "x03b", "Listado de Boletas x Estado"
    lsw.ListItems.Add , "x04b", "Resumen de Boletas x Estado"
      
    lblReporte.Tag = "x00b"
    lblReporte.Caption = "Listado de Boletas"
      
    cboProveedor.Enabled = False
    cboDetalle.Enabled = False
  
  Case 1 'Inventario
    lsw.ListItems.Add , "x00i", "Listado General"
    lsw.ListItems.Add , "x01i", "Listado x Estado"
    lsw.ListItems.Add , "x02i", "Listado Agrupado Proveedor-Estado"
    lsw.ListItems.Add , "x04i", "Listado Agrupado Estado-Proveedor"
    lsw.ListItems.Add , "x05i", "Listado Articulos Reenviados"
    lsw.ListItems.Add , "x06i", "Listado Articulos x Cliente"
    
    lblReporte.Tag = "x00i"
    lblReporte.Caption = "Listado General"
    
    cboProveedor.Enabled = True
    cboDetalle.Enabled = True
    cboBoleta.Enabled = False
    
  Case 2 'Tiempos
    lsw.ListItems.Add , "x00t", "Ciclo de Solución"
    lsw.ListItems.Add , "x01t", "Tiempos en Taller"
    
    cboProveedor.Enabled = True
    cboDetalle.Enabled = False
    cboBoleta.Enabled = False
    
    lblReporte.Tag = "x00t"
    lblReporte.Caption = "Ciclo de Solución"
    

End Select

End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0
Call sbCargaCombos
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = " and cedula in(select cedula from Pos_Reparacion)"
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
  gBusquedas.Filtro = " and cedula in(select cedula from Pos_Reparacion)"
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If
End Sub


