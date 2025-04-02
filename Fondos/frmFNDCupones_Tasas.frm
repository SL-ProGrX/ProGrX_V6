VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDCupones_Tasas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Configuración de Tasas para Certificados a Plazo"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18018
      _ExtentY        =   12303
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
      ItemCount       =   3
      Item(0).Caption =   "Catálogos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Tasas"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "Label3(0)"
      Item(1).Control(1)=   "cboT_Modelo"
      Item(1).Control(2)=   "GroupBox2"
      Item(1).Control(3)=   "lswT"
      Item(2).Caption =   "Planes"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "Label3(6)"
      Item(2).Control(1)=   "cboP_Modelo"
      Item(2).Control(2)=   "lswP"
      Begin XtremeSuiteControls.ListView lswT 
         Height          =   4815
         Left            =   -70000
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   10170
         _Version        =   1572864
         _ExtentX        =   17939
         _ExtentY        =   8493
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswP 
         Height          =   5895
         Left            =   -70000
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   10170
         _Version        =   1572864
         _ExtentX        =   17939
         _ExtentY        =   10398
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
         UseVisualStyle  =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   9615
         _Version        =   524288
         _ExtentX        =   16960
         _ExtentY        =   11033
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
         SpreadDesigner  =   "frmFNDCupones_Tasas.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboT_Modelo 
         Height          =   330
         Left            =   -68680
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   5295
         _Version        =   1572864
         _ExtentX        =   9340
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboP_Modelo 
         Height          =   330
         Left            =   -68680
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   5295
         _Version        =   1572864
         _ExtentX        =   9340
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   975
         Left            =   -70000
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   10215
         _Version        =   1572864
         _ExtentX        =   18018
         _ExtentY        =   1720
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.ComboBox cboT_Cupon 
            Height          =   330
            Left            =   2160
            TabIndex        =   10
            Top             =   360
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton btnT_Accion 
            Height          =   375
            Index           =   0
            Left            =   6480
            TabIndex        =   11
            Top             =   360
            Width           =   615
            _Version        =   1572864
            _ExtentX        =   1085
            _ExtentY        =   661
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
            Picture         =   "frmFNDCupones_Tasas.frx":080C
         End
         Begin XtremeSuiteControls.PushButton btnT_Exportar 
            Height          =   375
            Left            =   8760
            TabIndex        =   12
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Exportar"
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
            Picture         =   "frmFNDCupones_Tasas.frx":0F33
         End
         Begin XtremeSuiteControls.PushButton btnT_Accion 
            Height          =   375
            Index           =   1
            Left            =   7080
            TabIndex        =   14
            Top             =   360
            Width           =   615
            _Version        =   1572864
            _ExtentX        =   1085
            _ExtentY        =   661
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
            Picture         =   "frmFNDCupones_Tasas.frx":109D
         End
         Begin XtremeSuiteControls.ComboBox cboT_Plazo 
            Height          =   330
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtT_Tasa 
            Height          =   330
            Left            =   4320
            TabIndex        =   18
            Top             =   360
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   582
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
            Text            =   "0"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   2
            Left            =   4320
            TabIndex        =   17
            Top             =   120
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tasa"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Plazo"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   5
            Left            =   2160
            TabIndex        =   13
            Top             =   120
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cupón / Frecuencia Pago"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   6
         Left            =   -69760
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Modelo"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   0
         Left            =   -69760
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Modelo"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración de Tasas para CDPs"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   7335
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmFNDCupones_Tasas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()

On Error GoTo vError

vModulo = 18
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

'Inicializa Parametros
'strSQL = "exec spFndParametros_Liquida_Auto"
'Call ConectionExecute(strSQL)

tcMain.Item(0).Selected = True

With lswT.ColumnHeaders
    .Clear
    .Add , , "[Id]", 600
    .Add , , "Plazo", 1600, vbCenter
    .Add , , "Cupón/FP", 1600, vbCenter
    .Add , , "Tasa", 1000, vbRightJustify
    .Add , , "R. Fecha", 2100
    .Add , , "R. Usuario", 2100, vbCenter
End With

With lswP.ColumnHeaders
    .Clear
    .Add , , "[Operadora]", 1100
    .Add , , "Plan", 1600, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "R. Fecha", 2100
    .Add , , "R. Usuario", 2100, vbCenter
End With

'strSQL = "select C.CodPlan as 'IdX', P.DESCRIPCION as 'ItmX'" _
'       & " from FND_LIQUIDACION_AUTOMATICA_PLANES C inner join FND_PLANES P on C.Operadora  = P.COD_OPERADORA and C.CodPlan = P.COD_PLAN" _
'       & " order by C.IdRegistro  "
'Call sbCbo_Llena_New(cboR_Planes, strSQL, True, True)
'
'
'strSQL = " select convert(varchar(4),Anio) +  format(Mes, '00') as 'ItmX', convert(varchar(4),Anio) +  format(Mes, '00') as 'IdX'" _
'       & " From FND_LIQUIDACION_AUTOMATICA_RESUMEN" _
'       & "  group by anio, mes" _
'       & " order by Anio desc, mes desc"
'Call sbCbo_Llena_New(cboR_Proceso, strSQL, False, True)
'
'Call sbCbo_Copia(cboR_Proceso, cboT_Cupon)
'
'strSQL = "select rtrim(cod_Operadora) as 'IdX', rtrim(descripcion) as ItmX" _
'         & " from  fnd_Operadoras"
'Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)
'
'Call sbCbo_Copia(cboOperadora, cboCP_Operadora)

Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


