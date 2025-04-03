VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_Creditos_Carga_Lote 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Créditos: Carga en Lote"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13935
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   13935
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox gbTools 
      Height          =   1815
      Left            =   0
      TabIndex        =   21
      Top             =   7200
      Width           =   13935
      _Version        =   1441793
      _ExtentX        =   24580
      _ExtentY        =   3201
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtComision 
         Height          =   315
         Left            =   1200
         TabIndex        =   23
         Top             =   600
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNeto 
         Height          =   315
         Left            =   1200
         TabIndex        =   24
         Top             =   960
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   330
         Left            =   4320
         TabIndex        =   28
         Top             =   600
         Width           =   5535
         _Version        =   1441793
         _ExtentX        =   9763
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
      Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
         Height          =   330
         Left            =   6840
         TabIndex        =   29
         Top             =   240
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
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
         Left            =   11040
         TabIndex        =   30
         Top             =   720
         Width           =   1335
         _Version        =   1441793
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
         Picture         =   "frmCR_Creditos_Carga_Lote.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnCancelar 
         Height          =   495
         Left            =   12360
         TabIndex        =   31
         Top             =   720
         Width           =   1335
         _Version        =   1441793
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
         Picture         =   "frmCR_Creditos_Carga_Lote.frx":0727
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtProveedorNombre 
         Height          =   315
         Left            =   5040
         TabIndex        =   32
         Top             =   1320
         Width           =   4815
         _Version        =   1441793
         _ExtentX        =   8493
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtProveedorId 
         Height          =   315
         Left            =   4320
         TabIndex        =   33
         Top             =   1320
         Width           =   750
         _Version        =   1441793
         _ExtentX        =   1323
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboAplicacion 
         Height          =   330
         Left            =   11040
         TabIndex        =   38
         Top             =   240
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   255
         Index           =   15
         Left            =   3240
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Emitir"
         Height          =   255
         Index           =   13
         Left            =   5880
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblProveedor 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   3240
         TabIndex        =   35
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblCuentaTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Se utilizará la cuenta Bancaria Registrada de cada Persona."
         Height          =   255
         Left            =   4320
         TabIndex        =   34
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Neto"
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
         Left            =   0
         TabIndex        =   27
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comisión"
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
         Left            =   0
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Left            =   0
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
   End
   Begin XtremeSuiteControls.ComboBox cboLinea 
      Height          =   312
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboComision 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboConfirma 
      Height          =   312
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   1800
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_Creditos_Carga_Lote.frx":0E3D
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   13695
      _Version        =   524288
      _ExtentX        =   24156
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
      MaxCols         =   10
      SpreadDesigner  =   "frmCR_Creditos_Carga_Lote.frx":153D
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnCargar 
      Height          =   375
      Left            =   9960
      TabIndex        =   8
      Top             =   1800
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_Creditos_Carga_Lote.frx":1D10
   End
   Begin XtremeSuiteControls.PushButton btnInfo 
      Height          =   375
      Left            =   10440
      TabIndex        =   9
      Top             =   1800
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_Creditos_Carga_Lote.frx":2429
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   1800
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboPrideduc 
      Height          =   330
      Left            =   2520
      TabIndex        =   11
      Top             =   2640
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboDeductora 
      Height          =   315
      Left            =   2520
      TabIndex        =   15
      Top             =   2280
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboFrecuencia 
      Height          =   315
      Left            =   4440
      TabIndex        =   16
      Top             =   2640
      Width           =   1815
      _Version        =   1441793
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
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   11160
      TabIndex        =   17
      Top             =   1800
      Width           =   1335
      _Version        =   1441793
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
      Picture         =   "frmCR_Creditos_Carga_Lote.frx":2B42
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnInforme 
      Height          =   315
      Left            =   6360
      TabIndex        =   18
      ToolTipText     =   "Informe"
      Top             =   2640
      Width           =   315
      _Version        =   1441793
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Creditos_Carga_Lote.frx":3413
   End
   Begin XtremeSuiteControls.ComboBox cboDestino 
      Height          =   315
      Left            =   2520
      TabIndex        =   19
      Top             =   960
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destino/ Plan Inversión"
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
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   20
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   4
      Left            =   1440
      TabIndex        =   14
      Top             =   2280
      Width           =   1335
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
      Height          =   375
      Index           =   6
      Left            =   1440
      TabIndex        =   13
      Top             =   2640
      Width           =   855
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
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   12
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto Comisión "
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
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
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
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar"
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
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1695
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmCR_Creditos_Carga_Lote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mAseguradoraId As String
Dim mFrecuenciaPago As String

Private Sub sbLimpia()
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtComision.Text = 0
    txtNeto.Text = 0

End Sub


Private Sub btnAplicar_Click()
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen casos para procesar...[verifique!]", vbExclamation
       Exit Sub
    End If
    
    
    Dim vEmiteTipo As String
    
    vEmiteTipo = fxTipoDocumento(cboTipoDocumento.Text)
    
    If vEmiteTipo = "CP" Then
        If Not IsNumeric(txtProveedorId.Text) Then
            MsgBox "- No se ha indicado a ningún Proveedor para la Cuenta por Pagar", vbExclamation
            Exit Sub
        End If
    End If

    'Procesa Lote
    Call sbProcesar
    
End Sub

Private Sub btnBuscar_Click()
txtArchivo.Text = ""

With frmContenedor.CD
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



        txtArchivo.Text = .FileName

End With

End Sub

Private Sub btnCancelar_Click()
    txtArchivo.Text = ""
    Call sbLimpia
End Sub

Private Sub btnCargar_Click()
    Call sbCargaArchivo
End Sub

Private Sub btnExport_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 10
    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Monto"
    vHeaders.Headers(4) = "Plazo"
    vHeaders.Headers(5) = "Tasa"
    vHeaders.Headers(6) = "Cuota"
    vHeaders.Headers(7) = "Comisión"
    vHeaders.Headers(8) = "Documento"
    vHeaders.Headers(9) = "Notas"
    vHeaders.Headers(10) = "Cta.Bancaria"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Credito_Carga_Lote_Resultado")
End Sub

Private Sub btnInfo_Click()
  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
        & " - Columnas: CEDULA, NOMBRE, MONTO, PLAZO, COMISION, DOCUMENTO, NOTAS" & vbCrLf _
        & " - Nombre de la Hoja: IMPORT" _
    , vbInformation, "Información del Archivo de Carga"
End Sub

Private Sub btnInforme_Click()
Call sbReportes(cboLinea.ItemData(cboLinea.ListIndex), cboPrideduc.Text)
End Sub

Private Sub cboDeductora_Click()
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
  
  
vProceso = fxPrimerDeduccion(cboLinea.ItemData(cboLinea.ListIndex), cboDeductora.ItemData(cboDeductora.ListIndex))
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

Private Sub cboLinea_Click()
If vPaso Or cboLinea.ListCount = 0 Then Exit Sub

Call sbLimpia
 
Me.MousePointer = vbHourglass
    Call sbSTCargaCboDestinos(cboDestino, cboLinea.ItemData(cboLinea.ListIndex))
Me.MousePointer = vbDefault

End Sub


Private Sub cboPrideduc_Click()
If vPaso Or cboPrideduc.ListCount = 0 Then Exit Sub
 Call sbLimpia
End Sub

Private Sub cboTipoDocumento_Click()
If vPaso Then Exit Sub
If cboTipoDocumento.ListCount = 0 Then Exit Sub


Dim pTipo As String

pTipo = fxTipoDocumento(cboTipoDocumento.Text)

lblProveedor.top = lblCuentaTitulo.top

Select Case pTipo
    Case "CP"
        lblCuentaTitulo.Visible = False
        lblProveedor.Visible = True
    Case "CK", "TE"
        lblCuentaTitulo.Visible = True
        lblProveedor.Visible = False
    Case Else
        lblCuentaTitulo.Visible = False
        lblProveedor.Visible = False
End Select

txtProveedorId.top = lblProveedor.top
txtProveedorNombre.top = lblProveedor.top

txtProveedorId.Visible = lblProveedor.Visible
txtProveedorNombre.Visible = lblProveedor.Visible

End Sub

Private Sub chkExcel_Click()
 Call sbLimpia
End Sub

'Function fxFechaProcesoSiguiente(lngFecha As Long) As Long
'Dim strMes As String, strAnio As String, strFecha As String
'Dim iMes As Integer, iAnio As Integer
'strFecha = Trim(CStr(lngFecha))
'     strAnio = Mid(strFecha, 1, 4)
'     strMes = Mid(strFecha, 5, 2)
'     iAnio = CInt(strAnio)
'     iMes = CInt(strMes)
'     If CInt(strMes) = 12 Then
'         iAnio = iAnio + 1
'         strAnio = Trim(str(iAnio))
'         strMes = "01"
'     Else
'       Select Case iMes
'       Case 1, 2, 3, 4, 5, 6, 7, 8
'         iMes = iMes + 1
'         strMes = "0" & Trim(str(iMes))
'       Case 9, 10, 11
'         iMes = iMes + 1
'         strMes = Trim(str(iMes))
'       End Select
'     End If
'     fxFechaProcesoSiguiente = CLng(Trim(strAnio) & Trim(strMes))
'End Function

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim vProceso As Currency


vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mFrecuenciaPago = "M"

vPaso = True


cboLinea.Clear
cboConfirma.Clear


strSQL = "select COD_INSTITUCION AS 'IdX', DESCRIPCION  as 'ItmX'" _
       & "  From INSTITUCIONES Where ACTIVA = 1 And DEDUCCION_PLANILLA = 1"
Call sbCbo_Llena_New(cboDeductora, strSQL, False, True)
       

strSQL = "select rtrim(codigo) as 'IdX' , rtrim(descripcion) + '  ['  + rtrim(codigo) + ']' as 'ItmX'" _
       & " from catalogo where retencion = 'N' and activo = 1" _
       & " and codigo not in(select codigo_ase from fnd_planes)"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 cboLinea.AddItem rs!itmX & ""
 cboLinea.ItemData(cboLinea.ListCount - 1) = CStr(rs!IdX)
 
 cboConfirma.AddItem rs!itmX & ""
 cboConfirma.ItemData(cboConfirma.ListCount - 1) = CStr(rs!IdX)
 
 
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboLinea.Text = rs!itmX & ""
End If
rs.Close


strSQL = "select COD_CONDEB AS 'IdX', DESCRIPCION  as 'ItmX'" _
       & " From CONCEPTO_DESEMB Where ACTIVO = 1 And RETIENE = 1"
Call sbCbo_Llena_New(cboComision, strSQL, False, True)


cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("CP")
cboTipoDocumento.AddItem fxTipoDocumento("ND")
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.Text = fxTipoDocumento("TE")

cboAplicacion.Clear
cboAplicacion.AddItem "Formalización"
cboAplicacion.AddItem "Solicitud"
cboAplicacion.Text = "Formalización"


strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

txtArchivo.Text = ""

vGrid.MaxCols = 10
vGrid.MaxRows = 0

vProceso = GLOBALES.glngFechaCR
cboPrideduc.AddItem vProceso

For i = 1 To 6
  vProceso = fxFechaProcesoSiguiente(vProceso)
  cboPrideduc.AddItem vProceso
Next i
cboPrideduc.Text = GLOBALES.glngFechaCR

vPaso = False

Call cboTipoDocumento_Click
Call cboDeductora_Click

End Sub

Private Sub sbCargaArchivo()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset
Dim strCadena As String, curMonto As Currency, curComision As Currency, iLinea As Long

Dim pCliente As String, pProceso As Long, pComision As Currency, pNeto As Currency
Dim pCedula As String, pNombre As String, pReferencia As String
Dim pMonto As Currency, pPlazo As Integer, pTasa As Currency, pCuota As Currency

Dim pDocumento As String, pNotas As String

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboLinea.ListCount <= 0 Then Exit Sub
If cboComision.ListCount <= 0 Then Exit Sub


On Error GoTo vError

Me.MousePointer = vbHourglass

pReferencia = fxTipoDocumento(cboTipoDocumento.Text)

curMonto = 0
curComision = 0
iLinea = 0

pProceso = cboPrideduc.Text
pCliente = cboLinea.ItemData(cboLinea.ListIndex)

strSQL = "delete CRD_CREDITOS_CARGADO_H where codigo = '" & pCliente _
       & "' and PROCESO = " & pProceso
       
Call ConectionExecute(strSQL)

strSQL = "" 'Inicializa Bloque



Set rsExcel = Excel_Load(txtArchivo.Text, "Import")
iLinea = 0

Do While Not rsExcel.EOF

    iLinea = iLinea + 1
    
    pCedula = Trim(CStr(rsExcel!Cedula & ""))
    
    
     If pCedula <> "" Then
                
            pTasa = 0 'rsExcel!Tasa
            pCuota = 0
                
                
            pNombre = Trim(CStr(rsExcel!Nombre & ""))
            pMonto = CCur(IIf(IsNull(rsExcel!Monto), 0, rsExcel!Monto))
            pComision = CCur(IIf(IsNull(rsExcel!Comision), 0, rsExcel!Comision))
            pNeto = pMonto - pComision
            pDocumento = Trim(CStr(rsExcel!Documento & ""))
            pNotas = Trim(CStr(rsExcel!Notas & ""))
            
            
            pCedula = fxSysCleanTxtInject(pCedula)
            pNombre = fxSysCleanTxtInject(pNombre)
            pDocumento = fxSysCleanTxtInject(pDocumento)
            pNotas = fxSysCleanTxtInject(pNotas)
            
            
            curMonto = curMonto + pMonto
            curComision = curComision + pComision
            pPlazo = rsExcel!Plazo
                
                strSQL = strSQL & Space(10) & "Insert CRD_CREDITOS_CARGADO_H(LINEA,CODIGO,COD_REFERENCIA,PROCESO,CEDULA,MONTO,NOMBRE,TIPO" _
                                & ", PLAZO, TASA, CUOTA, COMISION, DOCUMENTO, NOTAS)" _
                        & " VALUES(" & iLinea & ",'" & pCliente & "','" & pReferencia & "'," & pProceso & ",'" & pCedula & "'," & pMonto & ",'" & pNombre _
                        & "','D'," & pPlazo & "," & pCuota & "," & pTasa & "," & pComision & ",'" & pDocumento & "','" & pNotas & "')"
     End If
  
     
     If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
     End If
  
  rsExcel.MoveNext
Loop



'Procesa Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


'Procesa Revisión de la Carga de Datos
curMonto = 0
curComision = 0

strSQL = "exec spCrd_Creditos_Lote_Cargado_Revisado '" & pCliente & "','" & pReferencia & "'," & pProceso & "," & cboBanco.ItemData(cboBanco.ListIndex) _
       & ", '" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

With vGrid
    .MaxRows = 0
    Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .col = 1
        .Text = rs!Cedula
        .col = 2
        .Text = rs!Nombre
        .col = 3
        .Text = CStr(rs!Monto)
        .col = 4
        .Text = CStr(rs!Plazo)
        .col = 5
        .Text = CStr(rs!Tasa)
        .col = 6
        .Text = CStr(rs!Cuota)
        .col = 7
        .Text = CStr(rs!Comision)
        
        .col = 8
        .Text = CStr(rs!Documento & "")
        .col = 9
        .Text = CStr(rs!Notas & "")
        .col = 10
        .Text = CStr(rs!CTA_BANCOS & "")
        
        
        curMonto = curMonto + rs!Monto
        curComision = curComision + rs!Comision
        
        rs.MoveNext
    Loop
    rs.Close
End With


'Totales
txtMonto.Text = Format(curMonto, "Standard")
txtComision.Text = Format(curComision, "Standard")
txtNeto.Text = Format(curMonto - curComision, "Standard")

Me.MousePointer = vbDefault

MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtComision.Text = 0
    txtNeto.Text = 0
End Sub

Private Function fxMaestroTesoreria(vTipoDocumento As String, vBanco As Integer, vMonto As Currency, vCodigo As String _
                              , vBeneficiario As String, vOP As Long, vDetalle1 As String, vReferencia As Long _
                              , vDetalle2 As String, vCuenta As String, vFecha As Date, vUnidad As String, vConcepto As String) As Long                                  'Regresa el NSOLICITUD
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngSol As Long

strSQL = "insert Tes_Transacciones(cod_concepto,cod_unidad,id_banco,tipo,tipo_Beneficiario,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza,user_solicita,autoriza,user_autoriza,fecha_autorizacion)" _
       & " values('" & vConcepto & "','" & vUnidad & "'," & vBanco & ",'" & vTipoDocumento & "',5,'" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','Pol','C','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S','" & glogon.Usuario & "'"
       
If UCase(vTipoDocumento) = "CK" Then
   strSQL = strSQL & ",'S','" & glogon.Usuario & "',dbo.MyGetdate())"
Else
   strSQL = strSQL & ",'N',null,null)"
End If
Call ConectionExecute(strSQL)

strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones"
Call OpenRecordSet(rsX, strSQL, 0)
 strSQL = "select * from Tes_Transacciones where nsolicitud = " & rsX!solicitud
rsX.Close

lngSol = 0

Call OpenRecordSet(rsX, strSQL, 0)
If Trim(rsX!Codigo) = Trim(vCodigo) Then lngSol = rsX!NSolicitud
rsX.Close

If lngSol = 0 Then
  strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones where codigo ='" & vCodigo _
         & "'"
  rsX.CursorLocation = adUseServer
  Call OpenRecordSet(rsX, strSQL, 0)
  lngSol = rsX!solicitud
  rsX.Close
End If

fxMaestroTesoreria = lngSol

End Function



Private Sub sbCreaDetalle(vSolicitud As Long, vCtaConta As String, vMonto As Currency, vDH As String, vLinea As Integer, vUnidad As String)
Dim strSQL As String

strSQL = "insert Tes_Trans_Asiento(nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad) values(" _
       & vSolicitud & ",'" & Trim(vCtaConta) & "'," & vMonto & ",'" & vDH _
       & "'," & vLinea & ",'" & vUnidad & "')"
Call ConectionExecute(strSQL)

End Sub

Private Function fxCtaBanco(pBanco As Integer) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select CTACONTA from Tes_Bancos where id_banco =" & pBanco
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
 fxCtaBanco = ""
Else
 fxCtaBanco = rsX!ctaConta
End If
rsX.Close
End Function


Private Function fxCtaPuente(pCodigo As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select CtaPuente from Catalogo where codigo  ='" & pCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
     fxCtaPuente = ""
Else
     fxCtaPuente = rsX!CtaPuente
End If

rsX.Close

End Function



Private Sub sbProcesar()
Dim strSQL As String, pCedula As String, i As Long
Dim pClienteId As String, pComisionRef As String, pProceso As Long
Dim pTesoreriaId As Long, vFecha As Date, pConfirma As String

Dim pCuenta As String, pUnidad As String, pConcepto As String, pTipo As String

Dim pPriDeduc As Currency, pProveedor As String, pBanco As Long

On Error GoTo vError



pClienteId = cboLinea.ItemData(cboLinea.ListIndex)
pConfirma = cboConfirma.ItemData(cboConfirma.ListIndex)

If pClienteId <> pConfirma Then
   MsgBox "La confirmación de la línea/cliente ha fallado, revise!", vbExclamation
   Exit Sub
End If

pComisionRef = cboComision.ItemData(cboComision.ListIndex)
pProceso = cboPrideduc.Text

pPriDeduc = cboPrideduc.Text & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex)

pBanco = cboBanco.ItemData(cboBanco.ListIndex)

pUnidad = "OC"
pConcepto = "CAR"

vFecha = fxFechaServidor

pTipo = fxTipoDocumento(cboTipoDocumento.Text)

If pTipo = "CP" Then
    pProveedor = txtProveedorId.Text
Else
    pProveedor = "Null"
End If


Me.MousePointer = vbHourglass

'Procesa Lote
strSQL = "exec spCrd_Creditos_Lote_Procesa '" & cboLinea.ItemData(cboLinea.ListIndex) & "', " & pProceso _
       & ", '" & pTipo & "', " & pPriDeduc & ", " & pBanco & ", " & pProveedor _
       & ", '" & pComisionRef & "','" & glogon.Usuario & "','" & cboDestino.ItemData(cboDestino.ListIndex) _
       & "', '" & Mid(cboAplicacion.Text, 1, 1) & "'"
Call ConectionExecute(strSQL)

Call sbReportes(cboLinea.ItemData(cboLinea.ListIndex), pProceso)

txtArchivo.Text = ""
Call sbLimpia

Me.MousePointer = vbDefault

MsgBox "Cargado y Registro de Solicitud en Bancos realizada satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbReportes(pLinea As String, ByVal pProceso As Currency)

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Crédito"
    
     .Connect = glogon.ConectRPT
    .Formulas(1) = "fxTitulo = 'Carga de Créditos en Lote'"
    .Formulas(2) = "fxSubTitulo= 'Línea: " & pLinea & "'"
    .Formulas(3) = "fxUsuario = '" & glogon.Usuario & "'"
    .Formulas(4) = "fxFecha = '" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(5) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
    .ReportFileName = SIFGlobal.fxPathReportes("Credito_Carga_Lote.rpt")
'    .SelectionFormula = "{CHEQUES.NSOLICITUD} = " & pTesoreria
'
'    .SubreportToChange = "sbDetalle"
'
    .StoredProcParam(0) = pLinea
    .StoredProcParam(1) = pProceso
    
    .Action = 1
'    .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbGuardaBk()
'Dim i As Long, vCadena As String, vTempo As String
'Dim vFile As String, vArchivo As String, vRuta As String, vFecha As Date
'Dim fnFile, vFechaProceso As Long
'
'
'vFecha = fxFechaServidor
'fnFile = FreeFile
'
'vFechaProceso = cboPrideduc.Text
'
'
''Crea Directorios
'
'On Error Resume Next
'
'MkDir SIFGlobal.DirectorioDeResultados
'MkDir SIFGlobal.DirectorioDeResultados & "\" & cboLinea.ItemData(cboLinea.ListIndex)
'MkDir SIFGlobal.DirectorioDeResultados & "\" & cboLinea.ItemData(cboLinea.ListIndex) & "\Cargado"
'MkDir SIFGlobal.DirectorioDeResultados & "\" & cboLinea.ItemData(cboLinea.ListIndex) & "\Cargado\" & vFechaProceso
'
'
'vRuta = SIFGlobal.DirectorioDeResultados & "\" & cboLinea.ItemData(cboLinea.ListIndex) & "\Cargado\" & vFechaProceso
'
'
'vArchivo = vFechaProceso & " [Cargado] " & cboLinea.ItemData(cboLinea.ListIndex) & " - " & cboComision.ItemData(cboComision.ListIndex) _
'          & " [" & glogon.Usuario & "].txt"
'
'
'vTempo = vRuta & "\" & vArchivo
'
'vFile = Dir(vTempo, vbArchive)
'
'If vFile = vArchivo Then  'El archivo existe
' Kill vTempo
'End If
'
'
'On Error GoTo vError
'
'Dim strSQL As String
'Dim vIdCliente As String, vInstitucion As Integer
'Dim vCedula As String, vNombre As String
'Dim vMonto As Currency, vInstExiste As String, vMovimiento As String
'
'
'
'vIdCliente = cboLinea.ItemData(cboLinea.ListIndex)
'vInstitucion = cboComision.ItemData(cboComision.ListIndex)
'vFechaProceso = cboPrideduc.Text
'
'strSQL = "delete CRD_CREDITOS_CARGADO_H where codigo = '" & vIdCliente _
'       & "' and PROCESO = '" & vFechaProceso & "' and cod_institucion = " & vInstitucion
'
'Call ConectionExecute(strSQL)
'
'
'
'Open vTempo For Output As #fnFile  ' Create file name.
'
'For i = 1 To vGrid.MaxRows
' vGrid.Row = i
' vGrid.Col = 1
' vCedula = Trim(vGrid.Text)
' vCadena = SIFGlobal.fxStringRelleno(vGrid.Text, "D", " ", 15)
'
' vGrid.Col = 2
' vNombre = Trim(vGrid.Text)
' vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "D", " ", 50)
'
' vGrid.Col = 3
' vMonto = CCur(vGrid.Text)
' vCadena = vCadena & Format(vGrid.Text, "000000000.00")
'
' vGrid.Col = 4
' vMovimiento = Trim(vGrid.Text)
' vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "I", " ", 10)
'
' vGrid.Col = 5
' vInstExiste = Trim(vGrid.Text)
' vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "I", " ", 10)
'
'
' strSQL = "Insert CRD_CREDITOS_CARGADO_H(LINEA,CODIGO,COD_INSTITUCION,PROCESO,CEDULA,MONTO,NOMBRE,MOVIMIENTO,TIPO, EXISTE_INST)" _
'         & " VALUES(" & i & ",'" & vIdCliente & "'," & vInstitucion & "," & vFechaProceso & ",'" & vCedula & "'," & vMonto & ",'" & vNombre _
'         & "','" & Mid(vMovimiento, 1, 1) & "','I','" & vInstExiste & "')"
'
' If vMovimiento <> "Error" Then
'    Call ConectionExecute(strSQL)
'End If
'
' Print #fnFile, vCadena
'Next i
'
'Close #fnFile
'
'Me.MousePointer = vbDefault
'
'Exit Sub
'
'vError:
'  Me.MousePointer = vbDefault
'  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Function fxRevisaInst(pCedula As String) As String
Dim Resultado As String


Resultado = "Ok"



fxRevisaInst = Resultado
End Function


Private Sub Form_Resize()
On Error Resume Next

Dim pH As Long, pW As Long


pH = 8505
pW = 14055

If Me.Height > pH Then
    pH = Me.Height
End If

If Me.Width > pW Then
    pW = Me.Width
End If

Me.Height = pH
Me.Width = pW

imgBanner.Width = pW

vGrid.Width = pW - (vGrid.Left + 200)

gbTools.Width = vGrid.Width

gbTools.top = pH - (gbTools.Height + 100)
vGrid.Height = gbTools.top - (vGrid.top + 150)


End Sub

Private Sub txtProveedorId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedorNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal

  txtProveedorId.Text = gBusquedas.Resultado
  txtProveedorNombre.Text = gBusquedas.Resultado3
End If

End Sub



Private Sub txtProveedorNombre_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal

  txtProveedorId.Text = gBusquedas.Resultado
  txtProveedorNombre.Text = gBusquedas.Resultado3
End If

End Sub

