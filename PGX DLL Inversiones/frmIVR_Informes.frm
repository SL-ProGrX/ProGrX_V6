VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.ShortcutBar.v20.0.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIVR_Informes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SGCI Informes de Inversiones"
   ClientHeight    =   7956
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10584
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7956
   ScaleWidth      =   10584
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   972
      Left            =   120
      TabIndex        =   1
      Top             =   6840
      Width           =   10332
      _Version        =   1310720
      _ExtentX        =   18224
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   732
         Left            =   7680
         TabIndex        =   2
         Top             =   240
         Width           =   2172
         _Version        =   1310720
         _ExtentX        =   3831
         _ExtentY        =   1291
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
         Picture         =   "frmIVR_Informes.frx":0000
      End
      Begin XtremeSuiteControls.CheckBox chkInformeResumen 
         Height          =   312
         Left            =   4680
         TabIndex        =   3
         Top             =   240
         Width           =   2652
         _Version        =   1310720
         _ExtentX        =   4678
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Informe Resumen    "
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   2772
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   4890
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
         Object.Width           =   4834
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpHistorico 
      Height          =   312
      Left            =   6000
      TabIndex        =   5
      Top             =   5160
      Width           =   1332
      _Version        =   1310720
      _ExtentX        =   2350
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   5880
      TabIndex        =   6
      Top             =   4320
      Width           =   1332
      _Version        =   1310720
      _ExtentX        =   2350
      _ExtentY        =   550
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
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   7200
      TabIndex        =   7
      Top             =   4320
      Width           =   1332
      _Version        =   1310720
      _ExtentX        =   2350
      _ExtentY        =   550
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
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   5880
      TabIndex        =   8
      Top             =   3960
      Width           =   2652
      _Version        =   1310720
      _ExtentX        =   4678
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
   Begin XtremeSuiteControls.ComboBox cboInstrumento 
      Height          =   312
      Left            =   2640
      TabIndex        =   9
      Top             =   1320
      Width           =   5892
      _Version        =   1310720
      _ExtentX        =   10393
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
   Begin XtremeSuiteControls.ComboBox cboAdministrador 
      Height          =   312
      Left            =   2640
      TabIndex        =   10
      Top             =   2640
      Width           =   5892
      _Version        =   1310720
      _ExtentX        =   10393
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
   Begin XtremeSuiteControls.ComboBox cboPortafolio 
      Height          =   312
      Left            =   2640
      TabIndex        =   11
      Top             =   3000
      Width           =   5892
      _Version        =   1310720
      _ExtentX        =   10393
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
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   312
      Left            =   8640
      TabIndex        =   12
      Top             =   4320
      Width           =   1812
      _Version        =   1310720
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
   Begin XtremeSuiteControls.ComboBox cboClasificacion 
      Height          =   312
      Left            =   2640
      TabIndex        =   22
      Top             =   1680
      Width           =   5892
      _Version        =   1310720
      _ExtentX        =   10393
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
   Begin XtremeSuiteControls.ComboBox cboReserva 
      Height          =   312
      Left            =   6000
      TabIndex        =   24
      Top             =   5640
      Width           =   2652
      _Version        =   1310720
      _ExtentX        =   4678
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
   Begin XtremeSuiteControls.ComboBox cboDivisa 
      Height          =   312
      Left            =   6000
      TabIndex        =   26
      Top             =   6000
      Width           =   2652
      _Version        =   1310720
      _ExtentX        =   4678
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
   Begin XtremeSuiteControls.ComboBox cboEmisor 
      Height          =   312
      Left            =   2640
      TabIndex        =   28
      Top             =   2160
      Width           =   5892
      _Version        =   1310720
      _ExtentX        =   10393
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Emisor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   5
      Left            =   1080
      TabIndex        =   29
      Top             =   2160
      Width           =   1572
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4320
      TabIndex        =   27
      Top             =   6000
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Reserva"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4320
      TabIndex        =   25
      Top             =   5640
      Width           =   1692
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Clasificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   1080
      TabIndex        =   23
      Top             =   1680
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Instrumento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   1080
      TabIndex        =   21
      Top             =   1320
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Administrador"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   1080
      TabIndex        =   20
      Top             =   2640
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Portafolio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   1080
      TabIndex        =   19
      Top             =   3000
      Width           =   1572
   End
   Begin VB.Label lblx02 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4200
      TabIndex        =   18
      Top             =   4320
      Width           =   1692
   End
   Begin VB.Label lblx01 
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
      Height          =   252
      Left            =   4200
      TabIndex        =   17
      Top             =   3960
      Width           =   1692
   End
   Begin VB.Label lblx04 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4320
      TabIndex        =   16
      Top             =   5160
      Width           =   1692
   End
   Begin VB.Label lblHistorico 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Height          =   252
      Left            =   7440
      TabIndex        =   15
      Top             =   5160
      Width           =   2412
   End
   Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
      Height          =   372
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   3480
      Width           =   4092
      _Version        =   1310720
      _ExtentX        =   7218
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Informes:"
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
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
      Height          =   372
      Index           =   1
      Left            =   4080
      TabIndex        =   13
      Top             =   3480
      Width           =   6492
      _Version        =   1310720
      _ExtentX        =   11451
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Filtros:"
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
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Inversiones"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   3
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   7212
   End
End
Attribute VB_Name = "frmIVR_Informes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset

Dim vPaso As Boolean, mReportKey As String
Dim vTitulo As String, vSubTitulo As String


Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
 dtpInicio.Enabled = False
 dtpCorte.Enabled = False
Else
 dtpInicio.Enabled = True
 dtpCorte.Enabled = True
End If

End Sub



Private Function fxSQL(pTipo As String) As String
Dim vCadena As String

vCadena = ""
vSubTitulo = ""
vTitulo = ""


'
'
'Select Case pTipo
'  Case "LA" 'Lista de Activos (Información Actual)
'    If chkTipoActivo.Value = vbUnchecked Then
'       vCadena = "{Activos_Principal.TIPO_ACTIVO} = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
'       vSubTitulo = vSubTitulo & " ¦ TIPO: " & cboTipo.ItemData(cboTipo.ListIndex)
'    End If
'
'    If chkDepartamentos.Value = vbUnchecked Then
'       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'       vCadena = vCadena & "{Activos_Principal.COD_DEPARTAMENTO} = '" & cboDep.ItemData(cboDep.ListIndex) & "'"
'       vSubTitulo = vSubTitulo & " ¦ DEPT: " & cboDep.ItemData(cboDep.ListIndex)
'    End If
'
'    If chkSeccion.Value = vbUnchecked Then
'       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'       vCadena = vCadena & "{Activos_Principal.COD_SECCION} = '" & cboSec.ItemData(cboSec.ListIndex) & "'"
'       vSubTitulo = vSubTitulo & " ¦ SEC: " & cboSec.ItemData(cboSec.ListIndex)
'    End If
'
'
'
'    If chkEstados.Value = vbUnchecked Then
'      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'      Select Case Mid(cboEstado.Text, 1, 2)
'        Case "01" 'Vigentes
'           vCadena = vCadena & "{Activos_Principal.ESTADO} = 'A' AND {Activos_Principal.VALOR_LIBROS_PERIODO} > {Activos_Principal.VALOR_DESECHO}"
'        Case "02" 'Depreciados
'           vCadena = vCadena & "{Activos_Principal.ESTADO} = 'A' AND {Activos_Principal.VALOR_LIBROS_PERIODO} <= {Activos_Principal.VALOR_DESECHO}"
'        Case "03" 'Retirados
'           vCadena = vCadena & "{Activos_Principal.ESTADO} = 'R'"
'      End Select
'      vSubTitulo = vSubTitulo & " ¦ ESTADO: " & cboEstado.Text
'    End If
'
'    If chkRes.Value = vbUnchecked Then
'      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'
'      vCadena = vCadena & "{Activos_Principal.IDENTIFICACION} = '" & txtResponsable.Tag & "'"
'      vSubTitulo = vSubTitulo & " ¦ RESPONSABLE: " & txtResponsable.Tag
'    End If
'
'    If chkFechas.Value = vbUnchecked Then
'      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'        vSubTitulo = vSubTitulo & " ¦ ADQUISICION: " & Format(dtpInicio.Value, "dd/mm/yyyy") _
'                   & " - " & Format(dtpCorte.Value, "dd/mm/yyyy")
'
'        vCadena = vCadena & "{Activos_Principal.FECHA_ADQUISICION} in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
'                & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
'    End If
'
'
'  '-------------------------------------------------------------------------------------------------------------------
'
'  Case "LH" 'Auxiliar: Lista de Activos
'       Call dtpHistorico_Change
'
'       vCadena = "{vActivos_AuxiliarConsolidado.ANIO} = " & Year(dtpHistorico.Value) _
'                & " AND {vActivos_AuxiliarConsolidado.MES} = " & Month(dtpHistorico.Value)
'       vSubTitulo = " PERIODO: " & lblHistorico.Caption & " ¦ " & fxActivos_PeriodoEstado(dtpHistorico.Value)
'
'    If chkTipoActivo.Value = vbUnchecked Then
'       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'       vCadena = vCadena & "{vActivos_AuxiliarConsolidado.TIPO_ACTIVO} = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
'       vSubTitulo = vSubTitulo & " ¦ TIPO: " & cboTipo.ItemData(cboTipo.ListIndex)
'    End If
'
'    If chkDepartamentos.Value = vbUnchecked Then
'       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'       vCadena = vCadena & "{vActivos_AuxiliarConsolidado.COD_DEPARTAMENTO} = '" & cboDep.ItemData(cboDep.ListIndex) & "'"
'       vSubTitulo = vSubTitulo & " ¦ DEPT: " & cboDep.ItemData(cboDep.ListIndex)
'    End If
'
'    If chkSeccion.Value = vbUnchecked Then
'       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'       vCadena = vCadena & "{vActivos_AuxiliarConsolidado.COD_SECCION} = '" & cboSec.ItemData(cboSec.ListIndex) & "'"
'       vSubTitulo = vSubTitulo & " ¦ SEC: " & cboSec.ItemData(cboSec.ListIndex)
'    End If
'
'    If chkEstados.Value = vbUnchecked Then
'      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'      Select Case Mid(cboEstado.Text, 1, 2)
'        Case "01" 'Vigentes
'           vCadena = vCadena & "{vActivos_AuxiliarConsolidado.VALOR_LIBROS_CONSOLIDADO} > {vActivos_AuxiliarConsolidado.VALOR_DESECHO}"
'        Case "02" 'Depreciados
'           vCadena = vCadena & "{vActivos_AuxiliarConsolidado.VALOR_LIBROS_CONSOLIDADO} <= {vActivos_AuxiliarConsolidado.VALOR_DESECHO}"
'      End Select
'      vSubTitulo = vSubTitulo & " ¦ ESTADO: " & cboEstado.Text
'    End If
'
'    If chkRes.Value = vbUnchecked Then
'      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'
'      vCadena = vCadena & "{vActivos_AuxiliarConsolidado.IDENTIFICACION} = '" & txtResponsable.Tag & "'"
'      vSubTitulo = vSubTitulo & " ¦ RESPONSABLE: " & txtResponsable.Tag
'    End If
'
'    If chkFechas.Value = vbUnchecked Then
'      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
'        vSubTitulo = vSubTitulo & " ¦ ADQUISICION: " & Format(dtpInicio.Value, "dd/mm/yyyy") _
'                   & " - " & Format(dtpCorte.Value, "dd/mm/yyyy")
'
'        vCadena = vCadena & "{vActivos_AuxiliarConsolidado.FECHA_ADQUISICION} in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
'                & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
'    End If
'
'
'End Select

fxSQL = vCadena

End Function

Private Sub cmdReporte_Click()
Dim vSQL As String


With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Inversiones"
 .Connect = glogon.ConectRPT
 
' .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
' .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
' .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
'
'  vSQL = fxSQL(Mid(mReportKey, 1, 2))
'
'Select Case mReportKey
'  Case "LA001"   'Lista General
'    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
'
'    If chkInformeResumen.Value = vbUnchecked Then
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneral.rpt")
'    Else
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralRsm.rpt")
'    End If
'
'  Case "LA002"   'Lista General x Departamento
'    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
'
'    If chkInformeResumen.Value = vbUnchecked Then
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralDeptTipo.rpt")
'    Else
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralDeptTipoRsm.rpt")
'    End If
'
'  Case "LA003"   'Lista General x Persona
'    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
'
'    If chkInformeResumen.Value = vbUnchecked Then
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralPersona.rpt")
'    Else
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralPersonaRsm.rpt")
'    End If
'
'
'  Case "LA004"   'Informe Contable
'    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
'
'    .ReportFileName = SIFGlobal.fxPathReportes("Activos_InformeContable.rpt")
'
'  Case "LH001"   'Auxiliar: Lista General
'    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
'
'    If chkInformeResumen.Value = vbUnchecked Then
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneral.rpt")
'    Else
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneralRsm.rpt")
'    End If
'
'  Case "LH002"   'Auxiliar: Lista x Departamento
'    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
'
'    If chkInformeResumen.Value = vbUnchecked Then
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneralDept.rpt")
'    Else
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneralDeptRsm.rpt")
'    End If
'
'  Case "LH003"   'Auxiliar: Lista x Persona
'    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
'
'    If chkInformeResumen.Value = vbUnchecked Then
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneralPersona.rpt")
'    Else
'       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneralPersonaRsm.rpt")
'    End If
'
'  Case "LH004"   'Auxiliar: Informe Contable
'    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
'
'    .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxInformeContable.rpt")
'
'
'  Case 4 'Depreciacion Historica
'  Case 5 'Boletas
'  Case 6 'Lista de Modificaciones
'  Case 7 'Asignacion de Polizas
'  Case 8 'Lista de Responsables
'  Case 9 'Lista de Departamentos y Secciones
'  Case 10 'Lista de Proveedores
'  Case 11 'Lista de Tipos Activos
'End Select
'
' .SelectionFormula = vSQL
' .PrintReport
End With

End Sub

Private Sub dtpHistorico_Change()

Select Case Month(dtpHistorico.Value)
 Case 1
    lblHistorico.Caption = "ENERO DE " & Year(dtpHistorico.Value)
 Case 2
    lblHistorico.Caption = "FEBRERO DE " & Year(dtpHistorico.Value)
 Case 3
    lblHistorico.Caption = "MARZO DE " & Year(dtpHistorico.Value)
 Case 4
    lblHistorico.Caption = "ABRIL DE " & Year(dtpHistorico.Value)
 Case 5
    lblHistorico.Caption = "MAYO DE " & Year(dtpHistorico.Value)
 Case 6
    lblHistorico.Caption = "JUNIO DE " & Year(dtpHistorico.Value)
 Case 7
    lblHistorico.Caption = "JULIO DE " & Year(dtpHistorico.Value)
 Case 8
    lblHistorico.Caption = "AGOSTO DE " & Year(dtpHistorico.Value)
 Case 9
    lblHistorico.Caption = "SETIEMBRE DE " & Year(dtpHistorico.Value)
 Case 10
    lblHistorico.Caption = "OCTUBRE DE " & Year(dtpHistorico.Value)
 Case 11
    lblHistorico.Caption = "NOVIEMBRE DE " & Year(dtpHistorico.Value)
 Case 12
    lblHistorico.Caption = "DICIEMBRE DE " & Year(dtpHistorico.Value)
End Select

End Sub

Private Sub sbListaReportes()

With lsw.ListItems
   .Clear
   .Add , "LA001", "Lista General"
   .Add , "LA002", "Lista x Administrador"
   .Add , "LA003", "Lista x Portafolio"
   .Add , "LA004", "Lista x Instrumento"
   .Add , "LA005", "Lista x Clasificación"
   .Add , "LA006", "Lista x Emisor"
   .Add , "LA007", "Lista x Fuente de Recursos"
   .Add , "LA008", "Lista x Reserva"
   .Add , "LA010", "Informe Contable"
   
   
   .Add , "LH001", "Auxiliar: Lista General"
   .Add , "LH002", "Auxiliar: Lista x Administrador"
   .Add , "LH003", "Auxiliar: Lista x Portafolio"
   .Add , "LH004", "Auxiliar: Lista x Instrumento"
   .Add , "LH005", "Auxiliar: Lista x Clasificación"
   .Add , "LH006", "Auxiliar: Lista x Emisor"
   .Add , "LH007", "Auxiliar: Lista x Fuente de Recursos"
   .Add , "LH008", "Auxiliar: Lista x Reserva"
   .Add , "LH010", "Auxiliar: Informe Contable"
   
End With
lsw.ListItems.Item(1).ForeColor = vbBlue
lsw.ListItems.Item(1).Bold = vbBlue

lsw.ListItems.Item(1).Selected = True

Call sbOpciones(1)



End Sub


Private Sub dtpInicio_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 36


Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vPaso = False
    

cboEstado.Clear
cboEstado.AddItem "Activos"
cboEstado.AddItem "Liquidados"
cboEstado.AddItem "Todos"
cboEstado.Text = "Activos"

'strSQL = "select rtrim(tipo_activo) as 'IdX' , rtrim(descripcion) as 'ItmX'" _
'       & " from Activos_tipo_activo order by tipo_activo"
'Call sbCbo_Llena_New(cboTipo, strSQL, False, True)
'
'strSQL = "select rtrim(cod_departamento) as 'IdX',  rtrim(descripcion) as 'ItmX'" _
'       & " from Activos_departamentos order by cod_departamento"
'Call sbCbo_Llena_New(cboDep, strSQL, False, True)

vPaso = True

Call sbListaReportes

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
'dtpHistorico.Value = gActivos.Periodo

Call dtpHistorico_Change




Call Formularios(Me)
Call RefrescaTags(Me)
End Sub




Private Sub lsw_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).ForeColor = vbBlack
  lsw.ListItems.Item(i).Bold = False
Next

lsw.SelectedItem.Bold = True
lsw.SelectedItem.ForeColor = vbBlue

mReportKey = lsw.SelectedItem.Key

Call sbOpciones(Mid(lsw.SelectedItem.Key, 1, 2))

End Sub

Private Sub sbOpciones(pTipo As String)



End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


strSQL = "select  rtrim(COD_INSTRUMENTO) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & "  From IVR_INSTRUMENTOS" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboInstrumento, strSQL, True, True)

strSQL = "select  rtrim(COD_EMISOR) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_EMISORES" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboEmisor, strSQL, True, True)


strSQL = "select  rtrim(COD_ADMINISTRADOR) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_ADMINISTRADOR" _
       & " Where ESTADO = 'A'" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboAdministrador, strSQL, True, True)

strSQL = "select  rtrim(COD_PORTAFOLIO) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_PORTAFOLIOS" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboPortafolio, strSQL, True, True)


strSQL = "select  rtrim(COD_CATEGORIA) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_CATEGORIA_TIPOS" _
       & " Where ACTIVO = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboClasificacion, strSQL, True, True)


strSQL = "select  rtrim(COD_RESERVA) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_RESERVAS" _
       & " Where ACTIVA = 1" _
       & " order by DESCRIPCION"
Call sbCbo_Llena_New(cboReserva, strSQL, True, True)


strSQL = "select  rtrim(COD_PERIODICIDAD) AS 'IdX', rtrim(DESCRIPCION) as 'itmX'" _
       & " From IVR_PERIODICIDAD" _
       & " Where ACTIVA = 1" _
       & " order by dias"
'Call sbCbo_Llena_New(cboPeriodicidad, strSQL, True, True)


strSQL = "select rtrim(COD_DIVISA) AS 'Idx', rtrim(DESCRIPCION) as 'ItmX'" _
       & " From vSys_Divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)


End Sub
