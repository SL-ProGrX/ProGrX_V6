VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAH_PlanillaDirecta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patrimonio: Deducciones de Planillas (Directa)"
   ClientHeight    =   8340
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11055
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4935
      Left            =   0
      TabIndex        =   24
      Top             =   2520
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   8705
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Cargados"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Inconsistencias"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "vGridInco"
      Item(1).Control(1)=   "txtIncoNDoc"
      Item(1).Control(2)=   "Label2(3)"
      Item(1).Control(3)=   "btnArchivo(3)"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4335
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   7646
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmAH_PlanillaDirecta.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridInco 
         Height          =   3855
         Left            =   -70000
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   6800
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "frmAH_PlanillaDirecta.frx":1185
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIncoNDoc 
         Height          =   315
         Left            =   -67960
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   550
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   375
         Index           =   3
         Left            =   -65080
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_PlanillaDirecta.frx":1832
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Comprobante"
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
         Left            =   -69760
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin XtremeSuiteControls.CheckBox chkConfirmar 
      Height          =   495
      Left            =   4680
      TabIndex        =   23
      Top             =   7680
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Confirmar No Casos y Monto"
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
   End
   Begin XtremeSuiteControls.CheckBox chkExcel 
      Height          =   255
      Left            =   7920
      TabIndex        =   22
      Top             =   1800
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Archivo Excel"
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
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
      Height          =   312
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   312
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.FlatEdit txtComprobante 
      Height          =   312
      Left            =   6600
      TabIndex        =   8
      Top             =   2160
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
      _ExtentY        =   550
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   432
      Left            =   2520
      TabIndex        =   9
      Top             =   1320
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12086
      _ExtentY        =   762
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
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   0
      Left            =   9480
      TabIndex        =   10
      Top             =   1320
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAH_PlanillaDirecta.frx":1F32
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   1
      Left            =   9960
      TabIndex        =   11
      Top             =   1320
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAH_PlanillaDirecta.frx":2632
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   372
      Index           =   2
      Left            =   10440
      TabIndex        =   12
      Top             =   1320
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAH_PlanillaDirecta.frx":2D4B
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   1080
      TabIndex        =   13
      Top             =   7800
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCasos 
      Height          =   315
      Left            =   2640
      TabIndex        =   14
      Top             =   7800
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1714
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtSocios 
      Height          =   315
      Left            =   3600
      TabIndex        =   15
      Top             =   7800
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1714
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   495
      Left            =   6840
      TabIndex        =   19
      Top             =   7560
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
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
      Picture         =   "frmAH_PlanillaDirecta.frx":3464
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   495
      Left            =   8880
      TabIndex        =   20
      Top             =   7560
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cancelar"
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
      Picture         =   "frmAH_PlanillaDirecta.frx":3B8B
   End
   Begin XtremeSuiteControls.PushButton btnBitacora 
      Height          =   495
      Left            =   8160
      TabIndex        =   21
      Top             =   7560
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   868
      _StockProps     =   79
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
      Picture         =   "frmAH_PlanillaDirecta.frx":428B
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   10560
      TabIndex        =   26
      ToolTipText     =   "Exportar a Excel"
      Top             =   2160
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   16777215
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAH_PlanillaDirecta.frx":4B37
   End
   Begin XtremeShortcutBar.ShortcutCaption scProcess 
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Cargando Archivo Espere!"
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Existe ?"
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
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   18
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Casos"
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
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   17
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Totales"
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
      Left            =   240
      TabIndex        =   16
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   1200
      TabIndex        =   7
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Comprobante"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   4800
      TabIndex        =   4
      Top             =   2160
      Width           =   1572
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Aporte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   1200
      TabIndex        =   3
      Top             =   1800
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11172
   End
End
Attribute VB_Name = "frmAH_PlanillaDirecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim mFecha As Date, vPaso As Boolean, mContrato As Long


Private Sub sbLimpia()

On Error GoTo vError

strSQL = "select dbo.fxPat_Planillas_Comprobante(" & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
        & ", " & cboProceso.Text & ", '" & UCase(Mid(cboTipo.Text, 1, 3)) & "') as 'NumDoc'"
Call OpenRecordSet(rs, strSQL)
  txtComprobante.Text = rs!NumDoc
rs.Close

    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtSocios.Text = 0
    txtArchivo.Text = ""
vError:
End Sub


Private Sub btnAplicar_Click()
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen deducciones cargadas...[verifique!]", vbExclamation
       Exit Sub
    End If
 
 If chkConfirmar.Value = xtpChecked Then
    Call sbProcesar
 End If
End Sub


Private Sub sbInconsistencias_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

vGridInco.MaxRows = 0

strSQL = "exec spPAT_PlanillaDirecta_Inconsistencias '" & txtIncoNDoc.Text & "'"
Call sbCargaGrid(vGridInco, 5, strSQL)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String

Select Case Index
  
  Case 0 'buscar
        txtArchivo.Text = ""
       Call sbBuscaArchivo(1)
  
  Case 1 'cargar
       Call sbCargaDeducciones(1)

  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: IMPORT" & vbCrLf _
              & " 3. Columnas.: CEDULA, NOMBRE, MONTO"
     
     MsgBox vMensaje, vbInformation


 Case 3 'Inconsistencias
    Call sbInconsistencias_Load
End Select

If Index = 3 Then
    tcMain.Item(1).Selected = True
Else
    tcMain.Item(0).Selected = True
End If

End Sub

Private Sub btnBitacora_Click()
 MsgBox "Bitácoras no activada!", vbExclamation
End Sub

Private Sub btnCancelar_Click()
    vGrid.MaxRows = 0
    txtArchivo.Text = ""
    
    txtCasos.Text = "0"
    txtMonto.Text = "0.00"
    txtSocios.Text = "0"
    
    chkConfirmar.Value = xtpUnchecked
End Sub

Private Sub btnExport_Click()
Dim vHeaders As vGridHeaders

Select Case tcMain.SelectedItem

Case 0 'Cargados
    vHeaders.Columnas = vGrid.MaxCols
    vHeaders.Headers(1) = "Cedula"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Existe?"
    vHeaders.Headers(4) = "Monto"
      
    Call sbSIFGridExportar(vGrid, vHeaders, "Patrimonio_" & txtComprobante.Text & "_Cargados")
Case 1 'Inconsistencias
    vHeaders.Columnas = vGrid.MaxCols
    vHeaders.Headers(1) = "Cedula"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Monto"
    vHeaders.Headers(4) = "Aplica?"
    vHeaders.Headers(5) = "Inconsistencia"
      
    Call sbSIFGridExportar(vGridInco, vHeaders, "Patrimonio_" & txtIncoNDoc.Text & "_Inconsistencias")

End Select

End Sub

Private Sub cboInstitucion_Click()

If vPaso Or cboInstitucion.ListCount = 0 Then Exit Sub
 
Call sbLimpia
 
End Sub


Private Sub cboProceso_Click()
 Call sbLimpia
End Sub

Private Sub cboTipo_Click()
Call cboInstitucion_Click
End Sub

Private Sub chkExcel_Click()
 Call sbLimpia
End Sub

Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()
Dim strSQL As String, i As Integer
Dim vProceso As Currency

vModulo = 18

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mFecha = fxFechaServidor

vPaso = True
    strSQL = "select cod_institucion as IdX,descripcion as ItmX from instituciones where activa = 1"
    Call sbCbo_Llena_New(cboInstitucion, strSQL, False, True)


cboTipo.AddItem "Obrero"
cboTipo.AddItem "Patronal"
cboTipo.AddItem "Capitalización"
cboTipo.Text = "Obrero"


tcMain.Item(0).Selected = True

txtArchivo.Text = ""
txtComprobante.Text = ""

vGrid.MaxCols = 4
vGrid.MaxRows = 0

vGridInco.MaxCols = 5
vGridInco.MaxRows = 0



vProceso = GLOBALES.glngFechaCR

For i = 1 To 6
  vProceso = fxFechaProcesoAnterior(vProceso)
  cboProceso.AddItem CStr(vProceso)
Next i

vProceso = GLOBALES.glngFechaCR
cboProceso.AddItem CStr(vProceso)

For i = 1 To 6
  vProceso = fxFechaProcesoSiguiente(vProceso)
  cboProceso.AddItem CStr(vProceso)
Next i
cboProceso.Text = CStr(GLOBALES.glngFechaCR)

vPaso = False

Call cboInstitucion_Click


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbCargaDeducciones(vTipo As Integer)

Dim pCedula As String, pNombre As String, pMonto As Currency
Dim pInstitucion As Long, pLinea As Long, pArchivo As String

Dim strCadena As String, curMonto As Currency
Dim fn As Long, lCasos As Long
Dim strMonto  As String
Dim strCedula As String
Dim strNombre As String
Dim i As Integer, vCampos As Boolean



On Error GoTo vError

vGrid.MaxRows = 0


If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboInstitucion.ListCount <= 0 Then
    MsgBox "No existe ninguna Institución, no se puede procesar el archivo...", vbCritical
    Exit Sub
End If


If fxAplicada Then
   MsgBox "Ya se aplico una planilla con esta fecha de proceso para la institución y el plan elegidos"
   Exit Sub
End If




pInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)

txtSocios.Text = 0
txtMonto.Text = 0
txtCasos.Text = 0

chkConfirmar.Value = xtpUnchecked

curMonto = 0
lCasos = 0 'Total
pArchivo = Dir(txtArchivo.Text, vbArchive)

Set rs = Excel_Load(txtArchivo.Text, "IMPORT")
    
'Validaciónn del Archivo
vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "CEDULA" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
         "Los campos son Cedula, Nombre, Monto ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "NOMBRE" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son Cedula, Nombre, Monto ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "MONTO" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son Cedula, Nombre, Monto ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If

'FIN: Validación del Archivo



scProcess.Visible = True
scProcess.Caption = "Subiendo Archivo, Espere!"
DoEvents

Me.MousePointer = vbHourglass


'Sube, Revisa y Carga
With vGrid
    
    pLinea = 0
    strSQL = ""
    
    Do While Not rs.EOF
      If Trim(rs!Cedula) <> "" Then
        pCedula = rs!Cedula
        pNombre = rs!Nombre & ""
        pMonto = rs!Monto
        pLinea = pLinea + 1
        
        
        If pLinea = 1 Then
            strSQL = strSQL & Space(10) & "exec spPAT_PlanillaDirecta_Sube " & pInstitucion & ",'" & txtComprobante.Text & "'," & cboProceso.Text _
                   & ",'" & Mid(cboTipo.Text, 1, 1) & "','" & pCedula & "'," & pMonto & ",'" & glogon.Usuario _
                   & "'," & pLinea & ", 1, '" & pNombre & "'"
           Call ConectionExecute(strSQL)
           strSQL = ""
        Else
            strSQL = strSQL & Space(10) & "exec spPAT_PlanillaDirecta_Sube " & pInstitucion & ",'" & txtComprobante.Text & "'," & cboProceso.Text _
                   & ",'" & Mid(cboTipo.Text, 1, 1) & "','" & pCedula & "'," & pMonto & ",'" & glogon.Usuario _
                   & "'," & pLinea & ", 0, '" & pNombre & "'"
        End If
        
        If Len(strSQL) > 40000 Then
           Call ConectionExecute(strSQL)
           If glogon.error Then
              Exit Sub
           End If
           
           scProcess.Caption = "Subiendo Archivo, Registros Procesados:  " & pLinea & ", Espere!"
           DoEvents
           strSQL = ""
        End If
        
      End If
      rs.MoveNext
    Loop
    rs.Close

'Procesa Ultimo Bloque

If Len(strSQL) > 0 Then
   scProcess.Caption = "Subiendo Archivo, Registros Procesados:  " & pLinea & ", Espere!"
   DoEvents
   Call ConectionExecute(strSQL)
   If glogon.error Then
      Exit Sub
   End If
   strSQL = ""
End If



scProcess.Caption = "Revisando Registros e Inconsistencias"
DoEvents

Me.MousePointer = vbHourglass

'Revisa Lote y lo Carga
strSQL = "exec spPAT_PlanillaDirecta_Consulta " & pInstitucion & ",'" & txtComprobante & "'," & cboProceso.Text _
                   & ",'" & Mid(cboTipo.Text, 1, 1) & "','" & glogon.Usuario & "', 1, '" & pArchivo & "'"
                   
Call OpenRecordSet(rs, strSQL)
If glogon.error Then
   Exit Sub
End If

    Do While Not rs.EOF
            pCedula = rs!Llave_01
            pNombre = rs!ref_01
            pMonto = rs!Monto_01
      
      
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .col = 1
            .Text = pCedula
            .col = 2
            .Text = pNombre
            .col = 3
            .Value = IIf((rs!Detalle = "-1"), 0, 1)
            
            .col = 4
            .Text = Format(pMonto, "Standard")
            
            If rs!Detalle = "-1" Then
               txtSocios.Text = CInt(txtSocios.Text) + 1
            End If
            
            curMonto = curMonto + pMonto
            txtMonto.Text = Format(curMonto, "Standard")
            
            txtCasos.Text = CLng(txtCasos.Text) + 1

      rs.MoveNext
    Loop
    rs.Close


End With 'vGrid


'Totales
txtMonto.Text = Format(curMonto, "Standard")

scProcess.Caption = "Cargando Inconsistencias!"
DoEvents
Call sbInconsistencias_Load

Me.MousePointer = vbDefault
scProcess.Visible = False

MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    scProcess.Visible = False
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia
End Sub

Private Sub sbProcesar()
Dim vTipoDoc As String, vNumDoc As String
Dim vCuenta  As String, vInstitucion As Long, vOperadora As Long, vPlan As String

On Error GoTo vError


scProcess.Visible = True
scProcess.Caption = "Procesando Planilla de Patrimonio, Espere!"
DoEvents

Me.MousePointer = vbHourglass

vTipoDoc = "PLA"
vNumDoc = txtComprobante.Text

vInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
 

strSQL = "update  SIF_DOCUMENTOS set CONSECUTIVO = ISNULL(CONSECUTIVO,0) + 1" _
       & "where TIPO_DOCUMENTO = 'PLA'"
Call ConectionExecute(strSQL)

scProcess.Visible = False


strSQL = "exec spPAT_PlanillaDirecta_Procesa " & vInstitucion & "," & cboProceso.Text & ",'" & Mid(cboTipo.Text, 1, 1) _
       & "','" & vNumDoc & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If glogon.error Then
   Exit Sub
End If
 
 
vNumDoc = rs!NumDoc
vTipoDoc = rs!TipoDoc
 
rs.Close

scProcess.Visible = False
Me.MousePointer = vbDefault
MsgBox "Proceso Aplicado Satisfactoriamente... Registros Procesados :" & vGrid.MaxRows

Call sbLimpia
Call sbImprimeRecibo(vNumDoc, vTipoDoc)

Exit Sub

vError:
    scProcess.Visible = False
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia
End Sub

Private Sub tlbProceso_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Aplicar"
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen deducciones cargadas...[verifique!]", vbExclamation
       Exit Sub
    End If
   
    Call sbProcesar
  
  Case "cancelar"
    vGrid.MaxRows = 0
    txtArchivo.Text = ""

End Select

End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)


End Sub


Private Sub sbBuscaArchivo(vTipo As Integer)


With frmContenedor.CD
    If vTipo = 1 Or chkExcel.Value = vbChecked Then
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]..."
        .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
            'Ok
        Else
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        
        txtArchivo.Text = .FileName
    
    Else
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Deducciones [Texto]..."
        .Filter = "*.txt"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If
        
        If UCase(Right(.FileName, 3)) = "XLS" Then
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        
        'If UCase(Right(.FileName, 3)) <> "TXT" Or UCase(Right(.FileName, 3)) <> "DAT" Then
         '   MsgBox "La Extensión del Archivo no es válido...", vbExclamation
         '   Exit Sub
        'End If

        txtArchivo.Text = .FileName

End If
End With

End Sub

Private Function fxAplicada() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select isnull(count(*),0) as Cantidad from sif_transacciones" _
       & " where Tipo_Documento = 'PLA'" _
       & " and Documento = '" & txtComprobante.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Cantidad > 0 Then
   fxAplicada = True
Else
   fxAplicada = False
End If
rs.Close

End Function


Private Sub txtComprobante_Change()
  txtIncoNDoc.Text = txtComprobante.Text
End Sub

Private Sub txtIncoNDoc_Change()
vGridInco.MaxCols = 5
vGridInco.MaxRows = 0

End Sub
