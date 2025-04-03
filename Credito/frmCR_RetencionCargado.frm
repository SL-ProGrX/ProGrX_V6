VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_RetencionCargado 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retenciones: Carga en Lote"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   15465
   Begin XtremeSuiteControls.CheckBox chkExcel 
      Height          =   255
      Left            =   8880
      TabIndex        =   10
      Top             =   2040
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Archivo Excel"
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   372
      Left            =   9480
      TabIndex        =   6
      Top             =   1560
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmCR_RetencionCargado.frx":0000
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   15255
      _Version        =   524288
      _ExtentX        =   26908
      _ExtentY        =   8705
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
      MaxCols         =   496
      SpreadDesigner  =   "frmCR_RetencionCargado.frx":0700
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnCargar 
      Height          =   372
      Left            =   9960
      TabIndex        =   7
      Top             =   1560
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmCR_RetencionCargado.frx":0E70
   End
   Begin XtremeSuiteControls.PushButton btnInfo 
      Height          =   372
      Left            =   10440
      TabIndex        =   8
      Top             =   1560
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmCR_RetencionCargado.frx":1589
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   372
      Left            =   2520
      TabIndex        =   9
      Top             =   1560
      Width           =   6852
      _Version        =   1572864
      _ExtentX        =   12086
      _ExtentY        =   656
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkInstitucion 
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Top             =   2280
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Revisar Institución"
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
   Begin XtremeSuiteControls.ComboBox cboPrideduc 
      Height          =   312
      Left            =   4200
      TabIndex        =   12
      Top             =   2040
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   4200
      TabIndex        =   13
      Top             =   2400
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboCliente 
      Height          =   312
      Left            =   2520
      TabIndex        =   14
      Top             =   240
      Width           =   6852
      _Version        =   1572864
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   2520
      TabIndex        =   15
      Top             =   600
      Width           =   6852
      _Version        =   1572864
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboDeductora 
      Height          =   312
      Left            =   2520
      TabIndex        =   16
      Top             =   960
      Width           =   6852
      _Version        =   1572864
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   1080
      TabIndex        =   18
      Top             =   8160
      Width           =   1575
      _Version        =   1572864
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCasos 
      Height          =   315
      Left            =   2640
      TabIndex        =   19
      Top             =   8160
      Width           =   975
      _Version        =   1572864
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtInclusion 
      Height          =   315
      Left            =   3600
      TabIndex        =   22
      Top             =   8160
      Width           =   975
      _Version        =   1572864
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtExclusion 
      Height          =   315
      Left            =   4560
      TabIndex        =   24
      Top             =   8160
      Width           =   975
      _Version        =   1572864
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCambio 
      Height          =   315
      Left            =   5520
      TabIndex        =   26
      Top             =   8160
      Width           =   975
      _Version        =   1572864
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtErr 
      Height          =   315
      Left            =   6480
      TabIndex        =   28
      Top             =   8160
      Width           =   975
      _Version        =   1572864
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboFrecuencia 
      Height          =   315
      Left            =   6120
      TabIndex        =   30
      Top             =   2040
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   495
      Left            =   10320
      TabIndex        =   31
      Top             =   7920
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Aplicar"
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
      Appearance      =   16
      Picture         =   "frmCR_RetencionCargado.frx":1CA2
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   495
      Left            =   11640
      TabIndex        =   32
      Top             =   7920
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cancelar"
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
      Appearance      =   16
      Picture         =   "frmCR_RetencionCargado.frx":23C9
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   11160
      TabIndex        =   33
      Top             =   1560
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
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
      TextAlignment   =   1
      Appearance      =   16
      Picture         =   "frmCR_RetencionCargado.frx":2ADF
      ImageAlignment  =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Errores"
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
      Index           =   11
      Left            =   6480
      TabIndex        =   29
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cambios"
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
      Index           =   10
      Left            =   5520
      TabIndex        =   27
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Exclus."
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
      Index           =   9
      Left            =   4560
      TabIndex        =   25
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Inclus."
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
      Index           =   8
      Left            =   3600
      TabIndex        =   23
      Top             =   7920
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
      TabIndex        =   21
      Top             =   8160
      Width           =   855
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
      TabIndex        =   20
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deductora"
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
      Index           =   3
      Left            =   1080
      TabIndex        =   17
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo deducción"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   2400
      Width           =   1572
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Primer deducción"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   2040
      Width           =   1572
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1452
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmCR_RetencionCargado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mFrecuenciaPago As String


Private Sub sbLimpia()
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtInclusion.Text = 0
    txtExclusion.Text = 0
    txtCambio.Text = 0
    txtErr.Text = 0
End Sub


Private Sub btnAplicar_Click()
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen deducciones cargadas...[verifique!]", vbExclamation
       Exit Sub
    End If
    Call sbProcesar
End Sub

Private Sub btnBuscar_Click()
        txtArchivo.Text = ""
        
        With frmContenedor.CD
         If chkExcel.Value = vbChecked Then
                .InitDir = "C:\"
                .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
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
        
                
         Else
                .InitDir = "C:\"
                .DialogTitle = "Localice Archivo de Deducciones [Texto]..."
                .Filter = "*.txt"
                .ShowOpen
                
                If .FileName = "" Then
                  MsgBox "Archivo no válido...", vbExclamation
                  Exit Sub
                End If
                
                If UCase(Right(.FileName, 3)) <> "TXT" Then
                  MsgBox "La Extensión del Archivo no es válido...", vbExclamation
                  Exit Sub
                End If
         End If
        
         txtArchivo.Text = .FileName
        
        End With

End Sub

Private Sub btnCancelar_Click()
    vGrid.MaxRows = 0
    txtArchivo.Text = ""
End Sub

Private Sub btnCargar_Click()
    Call sbCarga_Listado
End Sub

Private Sub btnExport_Click()

Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 9
    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Monto"
    vHeaders.Headers(4) = "Movimiento"
    vHeaders.Headers(5) = "Inst.Rev?"
    vHeaders.Headers(6) = "Plazo"
    vHeaders.Headers(7) = "Cuota"
    vHeaders.Headers(8) = "No. Operación"
    vHeaders.Headers(9) = "Fec.Formaliza"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Retenciones_Carga_Lote_Resultado")

End Sub

Private Sub btnInfo_Click()
  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
        & " - Columnas: CEDULA, NOMBRE, MONTO, CUOTA, OPERACION, FORMALIZACION, MOVIMIENTO, PLAZO" & vbCrLf _
        & " - Nombre de la Hoja: IMPORT" _
    , vbInformation, "Información del Archivo de Carga"
End Sub


Private Sub cboCliente_Click()
If vPaso Then Exit Sub
 Call sbLimpia
End Sub


Private Sub cboDeductora_Click()
 Call sbLimpia
 
 
If vPaso Then Exit Sub

On Error GoTo vError

Dim strSQL As String, rs As New ADODB.Recordset
Dim vProceso As Currency, pProcesoClean As Long

strSQL = "select rtrim(descripcion) as 'Descripcion', isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones " _
       & " where cod_institucion = " & cboDeductora.ItemData(cboDeductora.ListIndex)
Call OpenRecordSet(rs, strSQL)
    mFrecuenciaPago = rs!Frecuencia_ID
rs.Close

cboFrecuencia.Clear
Select Case mFrecuenciaPago
    Case "M" 'Mensual
        cboFrecuencia.AddItem "Mensual"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "0"
        cboFrecuencia.Text = "Mensual"
    
    Case "Q" 'Quincenal
        cboFrecuencia.AddItem "1er Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "1"
        cboFrecuencia.AddItem "2da Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "2"
End Select
  
  
vProceso = fxPrimerDeduccion(cboCliente.ItemData(cboCliente.ListIndex), cboDeductora.ItemData(cboDeductora.ListIndex))
pProcesoClean = vProceso

'cboMes.Text = fxConvierteMES(Val(Mid(pProcesoClean, 5, 2)))
'txtAnio.Text = Mid(pProcesoClean, 1, 4)
If mFrecuenciaPago = "Q" Then
    If (vProceso - pProcesoClean) = 0.1 Then
        cboFrecuencia.Text = "1er Quincena"
    Else
        cboFrecuencia.Text = "2da Quincena"
    End If
End If
Exit Sub

vError:

 
 
 
 
End Sub

Private Sub cboInstitucion_Click()
 
If vPaso Or cboInstitucion.ListCount = 0 Then Exit Sub

Call sbLimpia

vPaso = True
    Call sbDeductoras_Load(cboInstitucion.ItemData(cboInstitucion.ListIndex))
vPaso = False


Call sbCboAsignaDato(cboDeductora, cboInstitucion.Text, True, cboInstitucion.ItemData(cboInstitucion.ListIndex))

End Sub



Private Sub cboPrideduc_Click()
 Call sbLimpia
End Sub


Private Sub cboTipo_Click()
 Call sbLimpia
End Sub

Private Sub chkExcel_Click()
 Call sbLimpia
End Sub


Private Sub sbDeductoras_Load(pInstitucion As Long)
Dim strSQL As String

strSQL = "select COD_DEDUCTORA AS 'IdX', DESCRIPCION AS 'ItmX'" _
       & " From vAFI_Deductoras" _
       & " Where cod_institucion = " & pInstitucion

vPaso = True

Call sbCbo_Llena_New(cboDeductora, strSQL, False, True)

vPaso = False

End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()
Dim strSQL As String, i As Integer
Dim vProceso As Currency


vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

txtArchivo.Text = ""

vGrid.MaxCols = 9
vGrid.MaxRows = 0

mFrecuenciaPago = "M"

vPaso = True

    strSQL = "select rtrim(codigo) as 'IdX' , rtrim(descripcion) + '  ['  + rtrim(codigo) + ']' as 'ItmX'" _
           & " from catalogo where retencion = 'S' and activo = 1" _
           & " and codigo not in(select codigo_ase from fnd_planes)"
    Call sbCbo_Llena_New(cboCliente, strSQL, False, False)



    strSQL = "select cod_institucion as IdX,descripcion as ItmX from instituciones where activa = 1"
    Call sbCbo_Llena_New(cboInstitucion, strSQL, False, True)

    
    cboTipo.AddItem "Indefinida"
    cboTipo.AddItem "A Plazo"
    cboTipo.Text = "Indefinida"
    
    vProceso = GLOBALES.glngFechaCR
    cboPrideduc.AddItem vProceso
    
    For i = 1 To 6
      vProceso = fxFechaProcesoSiguiente(vProceso)
      cboPrideduc.AddItem vProceso
    Next i
    cboPrideduc.Text = GLOBALES.glngFechaCR

vPaso = False

Call cboInstitucion_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbCarga_Listado()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset
Dim strCadena As String, curMonto As Currency, iLinea As Long

Dim pCliente As String, pProceso As Long
Dim pCedula As String, pNombre As String, pMonto As Currency, pPlazo As Integer, pCuota As Currency
Dim pInstitucion As Long, pMovimiento As String, pDeductora As Long

Dim pFormaliza As Date, pOperacion As String, pFecha As Date

Dim fn, Casos(4) As Long


If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboCliente.ListCount <= 0 Then Exit Sub
If cboInstitucion.ListCount <= 0 Then Exit Sub


On Error GoTo vError

Me.MousePointer = vbHourglass

pFecha = fxFechaServidor


curMonto = 0
iLinea = 0

Casos(0) = 0 'Total
Casos(1) = 0 'Inclus
Casos(2) = 0 'Exclus
Casos(3) = 0 'Cambios
Casos(4) = 0 'Err

pProceso = cboPrideduc.Text
pInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
pDeductora = cboDeductora.ItemData(cboDeductora.ListIndex)
pCliente = cboCliente.ItemData(cboCliente.ListIndex)

strSQL = "delete CRD_RETENCION_CARGADO_H where codigo = '" & pCliente _
       & "' and PROCESO = '" & pProceso & "' and cod_institucion = " & pInstitucion
Call ConectionExecute(strSQL)

strSQL = "" 'Inicializa Bloque


If chkExcel.Value = vbChecked Then

        Set rsExcel = Excel_Load(txtArchivo.Text, "Import")
        iLinea = 0
            
            Do While Not rsExcel.EOF
              If Not IsNull(rsExcel!Cedula) Then
                        iLinea = iLinea + 1
                        pCedula = Trim(CStr(rsExcel!Cedula))
                        
        '                pCedula = Format(pCedula, "0000000000")
                        pNombre = Trim(CStr(rsExcel!Nombre))
                        pMonto = CCur(IIf(IsNull(rsExcel!Monto), 0, rsExcel!Monto))
 
                        pCuota = rsExcel!Cuota
                        pPlazo = IIf(IsNull(rsExcel!Plazo), 999, rsExcel!Plazo)
                        pOperacion = rsExcel!Operacion & ""
                        pFormaliza = IIf(IsNull(rsExcel!Formalizacion), pFecha, rsExcel!Formalizacion)
                        
                        curMonto = curMonto + CCur(IIf(IsNull(rsExcel!Monto), 0, rsExcel!Monto))
                        
                        Select Case Trim(rsExcel!Movimiento)
                           Case "I", "1"
                              pMovimiento = "Inclusión"
                              Casos(1) = Casos(1) + 1
                           Case "E", "3"
                              pMovimiento = "Exclusión"
                              Casos(2) = Casos(2) + 1
                           Case "C", "2"
                              pMovimiento = "Cambio"
                              Casos(3) = Casos(3) + 1
                           Case Else
                              pMovimiento = "Error"
                              Casos(4) = Casos(4) + 1
                        End Select
                        
'                        If Mid(cboTipo.Text, 1, 1) = "I" Then
'                           pPlazo = 999
'                           pCuota = CCur(IIf(IsNull(rsExcel!Monto), 0, rsExcel!Monto))
'                        Else
'                          'A Plazo
'                           pPlazo = rsExcel!Plazo
'                           pCuota = CCur(IIf(IsNull(rsExcel!Monto), 0, rsExcel!Monto)) / rsExcel!Plazo
'                        End If
                      
                         If pMovimiento <> "Error" Then
                                    strSQL = strSQL & Space(10) & "Insert CRD_RETENCION_CARGADO_H(LINEA,CODIGO,COD_INSTITUCION, COD_DEDUCTORA, PROCESO,CEDULA,MONTO" _
                                            & ",NOMBRE,MOVIMIENTO,TIPO, EXISTE_INST, PLAZO, CUOTA, OPERACION, FORMALIZA)" _
                                            & " VALUES(" & iLinea & ",'" & pCliente & "'," & pInstitucion & "," & pDeductora & "," & pProceso & ",'" & pCedula & "'," & pMonto _
                                            & ",'" & pNombre & "','" & Mid(pMovimiento, 1, 1) & "','I',Null," & pPlazo & "," & pCuota _
                                            & ",'" & pOperacion & "','" & Format(pFormaliza, "yyyy/mm/dd") & "')"
                         End If
                      
                      
                         If Len(strSQL) > 20000 Then
                            Call ConectionExecute(strSQL)
                            strSQL = ""
                         End If
              
              End If 'Null
              
              rsExcel.MoveNext
            Loop

Else 'Archivo Texto
        fn = FreeFile
        iLinea = 0
        
        Open txtArchivo.Text For Input As #fn    ' Lee el archivo.
         Do While Not EOF(fn)
           Input #fn, strCadena
           
           iLinea = iLinea + 1
           If Len(strCadena) >= 79 Then
                       pCedula = Mid(strCadena, 1, 15)
                       pNombre = Mid(strCadena, 16, 50)
                       pMonto = CCur(Mid(strCadena, 67, 12))
                       
                       curMonto = curMonto + CCur(Mid(strCadena, 67, 12))
                       
                       pMovimiento = Mid(strCadena, 79, 1)
                       
                        Select Case pMovimiento
                           Case "I", "1"
                              pMovimiento = "Inclusión"
                              Casos(1) = Casos(1) + 1
                           Case "E", "3"
                              pMovimiento = "Exclusión"
                              Casos(2) = Casos(2) + 1
                           Case "C", "2"
                              pMovimiento = "Cambio"
                              Casos(3) = Casos(3) + 1
                           Case Else
                              pMovimiento = "Error"
                              Casos(4) = Casos(4) + 1
                        End Select
                       
                   
                    If Mid(cboTipo.Text, 1, 1) = "I" Then
                       pPlazo = "999"
                       pCuota = pMonto
                    Else
                      'A Plazo
                       pPlazo = 1 'TODO:
                       pCuota = pMonto / pPlazo
                    End If
                   
                       If pMovimiento <> "Error" Then
                                  strSQL = strSQL & Space(10) & "Insert CRD_RETENCION_CARGADO_H(LINEA,CODIGO,COD_INSTITUCION,COD_DEDUCTORA, PROCESO,CEDULA,MONTO,NOMBRE,MOVIMIENTO,TIPO, EXISTE_INST, PLAZO, CUOTA)" _
                                          & " VALUES(" & iLinea & ",'" & pCliente & "'," & pInstitucion & "," & pDeductora & "," & pProceso & ",'" & pCedula & "'," & pMonto & ",'" & pNombre _
                                          & "','" & Mid(pMovimiento, 1, 1) & "','I',Null," & pPlazo & "," & pCuota & ")"
                       End If
                    
                    
                       If Len(strSQL) > 20000 Then
                          Call ConectionExecute(strSQL)
                          strSQL = ""
                       End If
                   
             End If 'Cadena Válida
         Loop
        Close #fn
        
End If 'Archivo Excel


'Procesa Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


'Procesa Revisión de Institución + Cargado de Datos
strSQL = "exec spCrd_Retenciones_Cargado_Revisado '" & pCliente & "'," & pInstitucion & "," & pProceso
Call OpenRecordSet(rs, strSQL)

With vGrid
    .MaxRows = 0
    'CRD_RETENCION_CARGADO_H: CUOTA, PLAZO, PROCESADO_DATE, PROCESADO_USUARIO
    Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1
        .Text = rs!Cedula
        .Col = 2
        .Text = rs!Nombre
        .Col = 3
        .Text = CStr(rs!Monto)
        .Col = 4
        .Text = rs!Movimiento_Name
        .Col = 5
        .Text = rs!EXISTE_INST
        .Col = 6
        .Text = CStr(rs!Plazo)
        .Col = 7
        .Text = CStr(rs!Cuota)
        .Col = 8
        .Text = rs!Operacion
        .Col = 9
        .Text = CStr(rs!Formaliza)
        
        
        rs.MoveNext
    Loop
    rs.Close
End With


'Totales
txtMonto.Text = Format(curMonto, "Standard")
txtCasos.Text = vGrid.MaxRows

txtInclusion.Text = Casos(1)
txtExclusion.Text = Casos(2)
txtCambio.Text = Casos(3)
txtErr.Text = Casos(4)


Me.MousePointer = vbDefault

If Casos(4) = 0 Then
    MsgBox "Información Cargada Satisfactoriamente", vbInformation
Else
    MsgBox "Información Cargada Pero con errores en algunas líneas (" & Casos(4) & ")", vbInformation
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtInclusion.Text = 0
    txtExclusion.Text = 0
    txtCambio.Text = 0
    txtErr.Text = 0
End Sub

Private Sub sbProcesar()
Dim strSQL As String, pCedula As String, i As Long
Dim pClienteId As String, pInstitucion As Integer, pProceso As Long, pPriDeduc As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass


pClienteId = cboCliente.ItemData(cboCliente.ListIndex)
pInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
pProceso = cboPrideduc.Text
pPriDeduc = cboPrideduc.Text & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex)

strSQL = ""
With vGrid
    'Procesa Cambios de Institucion
    For i = 1 To .MaxRows
        .Row = i
        .Col = 5
        If .Text = "Cambiar" Then
           .Col = 1
           pCedula = .Text
         
           strSQL = strSQL & Space(10) & "update CRD_RETENCION_CARGADO_H set EXISTE_INST = 'Cambiar'" _
                  & " Where codigo = '" & pClienteId & "' and cod_institucion = " & pInstitucion _
                  & " and Proceso = " & pProceso & " and Cedula = '" & pCedula & "'"
        End If
    
        If Len(strSQL) > 20000 Then
           Call ConectionExecute(strSQL)
           strSQL = ""
        End If
    
    Next i
End With

'Procesa Lote Final
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


'Procesa Lote de Retenciones Cargadas
strSQL = "exec spCrd_Retenciones_Cargado_Procesa '" & cboCliente.ItemData(cboCliente.ListIndex) & "'," & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & "," & cboPrideduc.Text & ",'" & glogon.Usuario & "', " & pPriDeduc
Call ConectionExecute(strSQL)


txtArchivo.Text = ""
vGrid.MaxRows = 0

Me.MousePointer = vbDefault

MsgBox "Cargado y Actualización de Retenciones aplicadas satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGuardaBk()
Dim i As Long, vCadena As String, vTempo As String
Dim vFile As String, vArchivo As String, vRuta As String, vFecha As Date
Dim fnFile, vFechaProceso As Long


vFecha = fxFechaServidor
fnFile = FreeFile

vFechaProceso = cboPrideduc.Text


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex)
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\Cargado"
MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\Cargado\" & vFechaProceso


vRuta = SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\Cargado\" & vFechaProceso


vArchivo = vFechaProceso & " [Cargado] " & Format(cboInstitucion.ItemData(cboInstitucion.ListIndex), "00") _
          & " " & cboInstitucion.Text & " [" & glogon.Usuario & "].txt"


vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Kill vTempo
End If


On Error GoTo vError

Dim strSQL As String
Dim vIdCliente As String, vInstitucion As Integer
Dim vCedula As String, vNombre As String
Dim vMonto As Currency, vInstExiste As String, vMovimiento As String



vIdCliente = cboCliente.ItemData(cboCliente.ListIndex)
vInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
vFechaProceso = cboPrideduc.Text

strSQL = "delete CRD_RETENCION_CARGADO_H where codigo = '" & vIdCliente _
       & "' and PROCESO = '" & vFechaProceso & "' and cod_institucion = " & vInstitucion

Call ConectionExecute(strSQL)



Open vTempo For Output As #fnFile  ' Create file name.

For i = 1 To vGrid.MaxRows
 vGrid.Row = i
 vGrid.Col = 1
 vCedula = Trim(vGrid.Text)
 vCadena = SIFGlobal.fxStringRelleno(vGrid.Text, "D", " ", 15)
 
 vGrid.Col = 2
 vNombre = Trim(vGrid.Text)
 vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "D", " ", 50)
 
 vGrid.Col = 3
 vMonto = CCur(vGrid.Text)
 vCadena = vCadena & Format(vGrid.Text, "000000000.00")
 
 vGrid.Col = 4
 vMovimiento = Trim(vGrid.Text)
 vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "I", " ", 10)
 
 vGrid.Col = 5
 vInstExiste = Trim(vGrid.Text)
 vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "I", " ", 10)


 strSQL = "Insert CRD_RETENCION_CARGADO_H(LINEA,CODIGO,COD_INSTITUCION,PROCESO,CEDULA,MONTO,NOMBRE,MOVIMIENTO,TIPO, EXISTE_INST)" _
         & " VALUES(" & i & ",'" & vIdCliente & "'," & vInstitucion & "," & vFechaProceso & ",'" & vCedula & "'," & vMonto & ",'" & vNombre _
         & "','" & Mid(vMovimiento, 1, 1) & "','I','" & vInstExiste & "')"
 
 If vMovimiento <> "Error" Then
    Call ConectionExecute(strSQL)
End If

 Print #fnFile, vCadena
Next i

Close #fnFile
  
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Function fxRevisaInst(pCedula As String) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim Resultado As String

Resultado = "Ignorar"

If chkInstitucion.Value = vbUnchecked Then
    Resultado = "Ok"
Else
    strSQL = "select isnull(cod_institucion,0) as cod_institucion from socios where cedula = '" & pCedula & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs.EOF And rs.BOF Then
        Resultado = "Ok"
    Else
       If rs!cod_institucion <> cboInstitucion.ItemData(cboInstitucion.ListIndex) Then
            Resultado = "Ignorar"
       Else
            Resultado = "Ok"
       End If
    End If
    rs.Close
End If


fxRevisaInst = Resultado
End Function

