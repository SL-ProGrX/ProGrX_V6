VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmInvReporteInventarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de Inventarios"
   ClientHeight    =   7356
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7356
   ScaleWidth      =   7560
   Begin XtremeSuiteControls.CheckBox chkConBodegas 
      Height          =   372
      Left            =   2400
      TabIndex        =   26
      Top             =   4680
      Width           =   4452
      _Version        =   1245187
      _ExtentX        =   7853
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Consolidar Inventarios de Bodegas   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Alignment       =   1
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   6960
      Top             =   600
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   372
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   1320
      Width           =   2052
      _Version        =   1245187
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Movimientos"
      BackColor       =   -2147483633
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
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1452
      Left            =   0
      TabIndex        =   8
      Top             =   5760
      Width           =   8172
      _Version        =   1245187
      _ExtentX        =   14414
      _ExtentY        =   2561
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   492
         Left            =   5280
         TabIndex        =   9
         Top             =   360
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
         Picture         =   "frmInvReporteInventarios.frx":0000
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   732
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   4212
      End
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   372
      Index           =   1
      Left            =   2760
      TabIndex        =   12
      Top             =   1320
      Width           =   2052
      _Version        =   1245187
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Cierre de Inventarios"
      BackColor       =   -2147483633
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
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   372
      Index           =   2
      Left            =   4800
      TabIndex        =   13
      Top             =   1320
      Width           =   2172
      _Version        =   1245187
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Inventario en Proceso"
      BackColor       =   -2147483633
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
   End
   Begin XtremeSuiteControls.ComboBox cboBodega 
      Height          =   312
      Left            =   2160
      TabIndex        =   14
      Top             =   2760
      Width           =   4692
      _Version        =   1245187
      _ExtentX        =   8276
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboReporte 
      Height          =   312
      Left            =   2160
      TabIndex        =   15
      Top             =   2280
      Width           =   4692
      _Version        =   1245187
      _ExtentX        =   8276
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   2160
      TabIndex        =   16
      Top             =   3960
      Width           =   4692
      _Version        =   1245187
      _ExtentX        =   8276
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboProExistencia 
      Height          =   312
      Left            =   2160
      TabIndex        =   17
      Top             =   1920
      Width           =   2412
      _Version        =   1245187
      _ExtentX        =   4255
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   3360
      TabIndex        =   18
      Top             =   4320
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
      Left            =   5640
      TabIndex        =   19
      Top             =   4320
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorteInv 
      Height          =   312
      Left            =   5640
      TabIndex        =   20
      Top             =   1920
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
   Begin XtremeSuiteControls.ComboBox cboLinea 
      Height          =   312
      Left            =   2160
      TabIndex        =   22
      Top             =   3120
      Width           =   4692
      _Version        =   1245187
      _ExtentX        =   8276
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboLineaSub 
      Height          =   312
      Left            =   2160
      TabIndex        =   24
      Top             =   3480
      Width           =   4692
      _Version        =   1245187
      _ExtentX        =   8276
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkArticulosSinMov 
      Height          =   372
      Left            =   2400
      TabIndex        =   27
      Top             =   5040
      Width           =   4452
      _Version        =   1245187
      _ExtentX        =   7853
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Mostrar Articulos sin Movimientos  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkCostos 
      Height          =   372
      Left            =   2400
      TabIndex        =   28
      Top             =   5400
      Width           =   4452
      _Version        =   1245187
      _ExtentX        =   7853
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Mostrar Costos de Articulos al Corte   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   732
      Left            =   2160
      TabIndex        =   25
      Top             =   240
      Width           =   4932
      _Version        =   1245187
      _ExtentX        =   8700
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Informes Inventarios (Existencias)"
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
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sub Familia"
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
      Index           =   9
      Left            =   720
      TabIndex        =   23
      Top             =   3480
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Línea / Familia"
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
      Index           =   8
      Left            =   720
      TabIndex        =   21
      Top             =   3120
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Reporte"
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
      Index           =   7
      Left            =   720
      TabIndex        =   7
      Top             =   2280
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tipo Mov."
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
      Index           =   6
      Left            =   720
      TabIndex        =   6
      Top             =   3960
      Width           =   1452
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
      Height          =   312
      Index           =   5
      Left            =   2160
      TabIndex        =   5
      Top             =   4320
      Width           =   1272
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
      Height          =   312
      Index           =   4
      Left            =   4560
      TabIndex        =   4
      Top             =   4320
      Width           =   1116
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Al Corte "
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
      Height          =   312
      Index           =   1
      Left            =   4560
      TabIndex        =   3
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fechas"
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
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   4320
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Existencia"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Bodega"
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
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmInvReporteInventarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSubTitulo As String
Dim vPaso As Boolean


Private Sub sbLlenaCbos()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
dtpCorteInv.Value = dtpInicio.Value

'Carga Bodegas
Call sbPosCombosCarga("Bodegas", cboBodega)
cboBodega.AddItem "[TODOS]"
cboBodega.Text = "[TODOS]"

'Carga Reportes Formatos
cboReporte.Clear
cboReporte.AddItem "01 - General x Bodegas"
cboReporte.AddItem "02 - Movimientos Un Solo Articulo"
cboReporte.AddItem "03 - Bodegas Agrupado por Articulos"
cboReporte.Text = "01 - General x Bodegas"


'Lineas
vPaso = True
    strSQL = "select cod_prodclas as 'IdX',rtrim(descripcion) as 'ItmX' from pv_prod_clasifica"
    Call sbCbo_Llena_New(cboLinea, strSQL, True, True)
vPaso = False

'Call
'Carga Existencias Rangos
Call sbInvExistenciaCargaCbo(cboProExistencia)

'Carga tipos de origenes
Call sbInvOrigenCargaCbo(cboTipo)

chkArticulosSinMov.Value = vbChecked
chkConBodegas.Value = vbUnchecked
chkConBodegas.Enabled = False
chkArticulosSinMov.Enabled = False


Call btnOpcion_Click(0)
Call cboLinea_Click

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub


Private Function fxFechaReportes(vTipo As Integer) As String

fxFechaReportes = " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")" _
                & " to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

End Function


Private Function fxSQL(i As Integer) As String
Dim vSQL As String

vSQL = ""
vSubTitulo = ""

Select Case i
  Case 0 'Movimientos
    Select Case cboTipo.Text
        Case "[TODOS]"
           'Nada
        Case "[SOLO ENTRADAS]"
           If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
           vSQL = vSQL & " {PV_INVENTARIO_MOV.TIPO} = 'E'"
           vSubTitulo = vSubTitulo & " ORIGEN: " & UCase(cboTipo.Text)
        
        Case "[SOLO SALIDAS]"
           If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
           vSQL = vSQL & " {PV_INVENTARIO_MOV.TIPO} = 'S'"
           vSubTitulo = vSubTitulo & " ORIGEN: " & UCase(cboTipo.Text)
        
        Case Else
           If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
           vSQL = vSQL & " {PV_INVENTARIO_MOV.ORIGEN} = '" & cboTipo.Text & "'"
           vSubTitulo = vSubTitulo & " ORIGEN: " & UCase(cboTipo.Text)
    End Select
     
    If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
    vSQL = vSQL & "CDATE({PV_INVENTARIO_MOV.FECHA}) " & fxFechaReportes(1)
    vSubTitulo = vSubTitulo & " INICIO:" & Format(dtpInicio.Value, "dd/mm/yyyy") _
              & " CORTE: " & Format(dtpCorte.Value, "dd/mm/yyyy")
   
    
    Select Case Mid(cboReporte.Text, 1, 2)
      Case "01" ' General x Bodegas
      Case "02" ' Movimientos Un Solo Articulo
             If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
             vSQL = vSQL & " {PV_INVENTARIO_MOV.COD_PRODUCTO} = '"
             vSQL = vSQL & InputBox("Código del Producto?", "Reportes de Inventarios")
             vSQL = vSQL & "'"
      Case "03" ' Bodegas Agrupado por Articulos
    End Select
    
  
  Case 1 'Cierre de Inventarios
     If cboProExistencia.Text <> "[TODOS]" Then
        If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
        Select Case Mid(cboProExistencia.Text, 1, 2)
          Case "00" 'Agotados
             vSQL = vSQL & " {PV_INVENTARIO.EXISTENCIA} = 0"
          Case "01" 'Minima
             vSQL = vSQL & " {PV_INVENTARIO.EXISTENCIA} <= {PV_PRODUCTOS.INVENTARIO_MINIMO}"
          Case "02" 'Maxima
             vSQL = vSQL & " {PV_INVENTARIO.EXISTENCIA} >= {PV_PRODUCTOS.INVENTARIO_MAXIMO}"
          Case "03" 'Inv (-) Reposición
             vSQL = vSQL & " {PV_INVENTARIO.EXISTENCIA} < 0"
          Case "04" 'Mayor Igual xx
             vSQL = vSQL & " {PV_INVENTARIO.EXISTENCIA} >= "
             vSQL = vSQL & InputBox("Existencia Mayor / Igual a ?", "Reportes de Inventarios")
         End Select
     End If
     
     If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
     vSQL = vSQL & " {PV_INVENTARIO.ANIO} = " & Year(dtpCorteInv.Value)
     vSQL = vSQL & " AND {PV_INVENTARIO.MES} = " & Month(dtpCorteInv.Value)
     vSubTitulo = vSubTitulo & " PERIODO: " & Year(dtpCorteInv.Value) & "-" & Format(Month(dtpCorteInv.Value), "00")
     
     'Si consolida las bodegas no permitir la asignacion del filtro de bodegas
     If chkConBodegas.Value = vbChecked Then
        fxSQL = vSQL
        Exit Function
     End If
    
     
     
  Case 2 'Inventario en Proceso
     
     If cboProExistencia.Text <> "[TODOS]" Then
        If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
        Select Case Mid(cboProExistencia.Text, 1, 2)
          Case "00" 'Agotados
             vSQL = vSQL & " ({PV_INVENTARIO_PROCESO.EXISTENCIA_INICIAL} + {PV_INVENTARIO_PROCESO.ENTRADAS} - {PV_INVENTARIO_PROCESO.SALIDAS}) = 0"
          Case "01" 'Minima
             vSQL = vSQL & " ({PV_INVENTARIO_PROCESO.EXISTENCIA_INICIAL} + {PV_INVENTARIO_PROCESO.ENTRADAS} - {PV_INVENTARIO_PROCESO.SALIDAS}) <= {PV_PRODUCTOS.INVENTARIO_MINIMO}"
          Case "02" 'Maxima
             vSQL = vSQL & " ({PV_INVENTARIO_PROCESO.EXISTENCIA_INICIAL} + {PV_INVENTARIO_PROCESO.ENTRADAS} - {PV_INVENTARIO_PROCESO.SALIDAS}) >= {PV_PRODUCTOS.INVENTARIO_MAXIMO}"
          Case "03" 'Inv (-) Reposición
             vSQL = vSQL & " ({PV_INVENTARIO_PROCESO.EXISTENCIA_INICIAL} + {PV_INVENTARIO_PROCESO.ENTRADAS} - {PV_INVENTARIO_PROCESO.SALIDAS}) < 0"
          Case "04" 'Mayor Igual xx
             vSQL = vSQL & " ({PV_INVENTARIO_PROCESO.EXISTENCIA_INICIAL} + {PV_INVENTARIO_PROCESO.ENTRADAS} - {PV_INVENTARIO_PROCESO.SALIDAS}) >= "
             vSQL = vSQL & InputBox("Existencia Mayor / Igual a ?", "Reportes de Inventarios")
         End Select
     End If
     
     If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
     vSQL = vSQL & " {PV_INVENTARIO_PROCESO.USUARIO} = '" & glogon.Usuario & "'"
     vSubTitulo = vSubTitulo & " AL CORTE: " & Format(dtpCorteInv.Value, "dd/mm/yyyy")
      
     'Si consolida las bodegas no permitir la asignacion del filtro de bodegas
     If chkConBodegas.Value = vbChecked Then
        fxSQL = vSQL
        Exit Function
     End If

End Select

If cboBodega.Text <> "[TODOS]" Then
   If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
   vSQL = vSQL & " {PV_BODEGAS.COD_BODEGA} = '" & cboBodega.ItemData(cboBodega.ListIndex) & "'"
   vSubTitulo = vSubTitulo & " BODEGA: " & cboBodega.Text
End If

If cboLinea.Text <> "TODOS" Then
   If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
   vSQL = vSQL & " {PV_PRODUCTOS.COD_PRODCLAS} = " & cboLinea.ItemData(cboLinea.ListIndex) '& "'"
   vSubTitulo = vSubTitulo & " LINEA: " & cboLinea.Text
End If

If cboLineaSub.Text <> "TODOS" And cboLineaSub.ListCount > 0 Then
   If Len(vSQL) > 0 Then vSQL = vSQL & " AND "
   vSQL = vSQL & " {PV_PRODUCTOS.COD_LINEA_SUB} = '" & cboLineaSub.ItemData(cboLineaSub.ListIndex) & "'"
   vSubTitulo = vSubTitulo & " SUB LINEA: " & cboLineaSub.Text
End If

fxSQL = vSQL

End Function

Private Sub btnOpcion_Click(Index As Integer)

btnOpcion.Item(0).Checked = False
btnOpcion.Item(1).Checked = False
btnOpcion.Item(2).Checked = False

btnOpcion.Item(Index).Checked = True

Select Case Index
  Case 0 'Movimientos inventarios
    dtpCorteInv.Enabled = False
    dtpInicio.Enabled = True
    dtpCorte.Enabled = True
    cboReporte.Enabled = True
    cboTipo.Enabled = True
    chkConBodegas.Enabled = False
    chkArticulosSinMov.Enabled = False
    chkCostos.Enabled = False
  Case 1, 2 'Movimientos inventarios
    dtpCorteInv.Enabled = True
    dtpInicio.Enabled = False
    dtpCorte.Enabled = False
    cboReporte.Enabled = False
    cboTipo.Enabled = False
    chkConBodegas.Enabled = True
    chkArticulosSinMov.Enabled = True
    chkCostos.Enabled = True
End Select

End Sub

Private Sub btnReporte_Click()
Dim vSQL As String
Dim vProdMov As Boolean, vConBodegas As Boolean

Dim pLinea As String, pLineaSub As String

Me.MousePointer = vbHourglass


vProdMov = IIf((chkArticulosSinMov.Value = vbChecked), True, False)
vConBodegas = IIf((chkConBodegas.Value = vbChecked), True, False)

If cboLinea.Text = "TODOS" Then
   pLinea = "Null"
   pLineaSub = "Null"
Else
   pLinea = cboLinea.ItemData(cboLinea.ListIndex)
   If cboLineaSub.Text = "0" Or cboLineaSub.Text = "TODOS" Or cboLineaSub.ListCount = 0 Then
        pLineaSub = "Null"
   Else
        pLineaSub = cboLineaSub.ItemData(cboLineaSub.ListIndex)
   End If
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
 .WindowTitle = "Reportes del Módulo de Invertarios"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 

 Select Case True
    Case btnOpcion.Item(0).Checked  'Movimientos
         vSQL = fxSQL(0)
         .Formulas(3) = "fxTitulo = 'MOVIMIENTOS DE INVENTARIOS'"
         .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
         
            Select Case Mid(cboReporte.Text, 1, 2)
              Case "01" ' General x Bodegas
                .ReportFileName = SIFGlobal.fxPathReportes("Inventario_Movimientos.rpt")
              Case "02" ' Movimientos Un Solo Articulo
                .ReportFileName = SIFGlobal.fxPathReportes("Inventario_Movimientos.rpt")
              Case "03" ' Bodegas Agrupado por Articulos
                .ReportFileName = SIFGlobal.fxPathReportes("Inventario_MovProdGrp.rpt")
            End Select
         
         
         .SelectionFormula = vSQL
    
    Case btnOpcion.Item(1).Checked 'Cierre de Inventarios
         
         vSQL = fxSQL(1)
         If Not vConBodegas Then
            .Formulas(3) = "fxTitulo = 'CIERRE DE INVENTARIOS'"
            .ReportFileName = SIFGlobal.fxPathReportes("Inventario_CierreInv.rpt")
         Else
            .Formulas(3) = "fxTitulo = 'CIERRE DE INVENTARIOS CONSOLIDADOS'"
            .ReportFileName = SIFGlobal.fxPathReportes("Inventario_CierreInvCon.rpt")
         End If
         .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
         .SelectionFormula = vSQL
    
    Case btnOpcion.Item(2).Checked 'Inventarios en Proceso
         
         If Not vConBodegas Then
            If cboBodega.Text = "[TODOS]" Then
             Me.MousePointer = vbDefault
             MsgBox "Debe de Seleccionar una bodega...", vbExclamation
             Exit Sub
            End If
         End If
         
         lbl.Caption = "**Procesando Inventario en Proceso(Espere)**"
         lbl.Refresh
         
         
         
         Call sbInvInventarioProceso(dtpCorteInv.Value, cboBodega.ItemData(cboBodega.ListIndex), vProdMov, vConBodegas _
                                , "", pLinea, pLineaSub)
         
         lbl.Caption = ""
         
         vSQL = fxSQL(2)
         If Not vConBodegas Then
            .Formulas(3) = "fxTitulo = 'INVENTARIOS EN PROCESO'"
            .ReportFileName = SIFGlobal.fxPathReportes("Inventario_EnProcesoInv.rpt")
         Else
            .Formulas(3) = "fxTitulo = 'INVENTARIOS EN PROCESO CONSOLIDADO'"
            .ReportFileName = SIFGlobal.fxPathReportes("Inventario_EnProcesoInvCon.rpt")
         End If
         .Formulas(4) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
         .SelectionFormula = vSQL
    
    Case Else
 End Select
 
 .PrintReport

End With

Me.MousePointer = vbDefault

End Sub

Private Sub cboBodega_Click()
If cboBodega.Text = "[TODOS]" Then
 If btnOpcion.Item(0).Checked = True Then
   chkConBodegas.Enabled = False
 Else
   chkConBodegas.Enabled = True
 End If
Else
chkConBodegas.Enabled = False
End If
End Sub




Private Sub cboLinea_Click()
If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


If cboLinea.Text = "TODOS" Then
    strSQL = "select COD_LINEA_SUB as 'IdX',  DESCRIPCION as 'ItmX'" _
        & " From PV_PROD_CLASIFICA_SUB where COD_PRODCLAS = '0'"

Else
    strSQL = "select COD_LINEA_SUB as 'IdX',  DESCRIPCION as 'ItmX'" _
        & " From PV_PROD_CLASIFICA_SUB where COD_PRODCLAS = " & cboLinea.ItemData(cboLinea.ListIndex)
End If

Call sbCbo_Llena_New(cboLineaSub, strSQL, False, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture


End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbLlenaCbos

End Sub
