VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmTES_CopiaEsquema 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia Esquema de una Solicitud"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   HelpContextID   =   1005
   Icon            =   "frmTES_CopiaEsquema.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   8685
   Begin XtremeSuiteControls.FlatEdit txtBanco 
      Height          =   315
      Left            =   1800
      TabIndex        =   11
      Top             =   1440
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.FlatEdit txtNumeroSolicitud 
      Height          =   432
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.PushButton btnCopiar 
      Height          =   675
      Left            =   6000
      TabIndex        =   0
      Top             =   8040
      Width           =   2295
      _Version        =   1441793
      _ExtentX        =   4048
      _ExtentY        =   1191
      _StockProps     =   79
      Caption         =   "Copiar Solicitud"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmTES_CopiaEsquema.frx":6852
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtTipo 
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Top             =   1800
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.FlatEdit txtBeneficiario 
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Top             =   2280
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1800
      TabIndex        =   14
      Top             =   2640
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      Height          =   315
      Left            =   5880
      TabIndex        =   16
      Top             =   3000
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   1800
      TabIndex        =   15
      Top             =   3000
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUnidad 
      Height          =   315
      Left            =   1800
      TabIndex        =   17
      Top             =   5160
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.FlatEdit txtConcepto 
      Height          =   315
      Left            =   1800
      TabIndex        =   18
      Top             =   5520
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   6480
      Width           =   8415
      _Version        =   1441793
      _ExtentX        =   14843
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Notas: "
      ForeColor       =   4210752
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   975
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   7575
         _Version        =   1441793
         _ExtentX        =   13361
         _ExtentY        =   1720
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtTipoId 
      Height          =   315
      Left            =   5880
      TabIndex        =   21
      Top             =   2640
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCtaIBAN 
      Height          =   315
      Left            =   1800
      TabIndex        =   23
      Top             =   3480
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   315
      Left            =   1800
      TabIndex        =   25
      Top             =   3960
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.FlatEdit txtCuentaOrigen 
      Height          =   315
      Left            =   1800
      TabIndex        =   27
      Top             =   4320
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.FlatEdit txtDivisa 
      Height          =   315
      Left            =   1800
      TabIndex        =   29
      Top             =   4800
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCorreo 
      Height          =   315
      Left            =   4320
      TabIndex        =   31
      Top             =   4800
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7011
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDetalle 
      Height          =   555
      Left            =   1800
      TabIndex        =   33
      Top             =   5880
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
      _ExtentY        =   979
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   14
      Left            =   480
      TabIndex        =   34
      Top             =   5880
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Detalle"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   13
      Left            =   3600
      TabIndex        =   32
      Top             =   4800
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Correo"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   12
      Left            =   480
      TabIndex        =   30
      Top             =   4800
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Divisa"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   11
      Left            =   480
      TabIndex        =   28
      Top             =   4320
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cuenta Origen"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   26
      Top             =   3960
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Origen"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   24
      Top             =   3480
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cuenta IBAN"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   8
      Left            =   4440
      TabIndex        =   22
      Top             =   2640
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo Id"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Unidad"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   9
      Top             =   5520
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Concepto"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fecha Solicitud"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Monto"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Identificación"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Beneficiario"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cuenta"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Solicitud a Copiar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmTES_CopiaEsquema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbLimpiaDatos()

txtBeneficiario = ""
txtCodigo = ""
txtFecha = ""
txtTipo = ""
txtMonto = ""
txtBanco = ""
txtConcepto = ""
txtUnidad = ""
txtBanco.Tag = 0
txtTipo.Tag = ""

txtNotas.Text = ""

End Sub

Private Function fxValidaSolicitud() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Valida que la Solicitud por duplicar contenga identificador de Banco y
'               codigo.
'REFERENCIAS:   Ninguna.
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

fxValidaSolicitud = True

If Len(txtBanco.Tag) = 0 Then fxValidaSolicitud = False
If Len(Trim(txtCodigo)) = 0 Then fxValidaSolicitud = False

End Function

Private Sub btnCopiar_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Duplica una determinada solicitud ya ingresada a Tesoreria. Tambien duplica
'               el detalle de la misma solicitud para la nueva.
'REFERENCIAS:   Bitacora - (Registra movimientos sobre la Base de Datos)
'               sbLimpiaDatos - (Limpia los objetos de entrada de datos)
'               fxValidaSolicitud - (Valida que la Solicitud por duplicar contenga
'               identificador de Banco y codigo)
'               fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, rs As New ADODB.Recordset
Dim lngNewSol As Long, i As Integer, pNotas As String

If Not fxValidaSolicitud Then
  MsgBox "Indique un número de solicitud válido...", vbExclamation
  Exit Sub
End If

i = MsgBox("Esta seguro que desea Duplicar Esta Solicitud", vbYesNo)
If i = vbNo Then Exit Sub
    
On Error GoTo vError

Me.MousePointer = vbHourglass

pNotas = Mid(fxSysCleanTxtInject(txtNotas.Text), 1, 500)
    
strSQL = "exec spTES_Transaccion_Copia " & txtNumeroSolicitud.Text & ", '" & pNotas & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then
    lngNewSol = rs!TesoreriaId
Else
    lngNewSol = 0
End If

Me.MousePointer = vbDefault

'Bitacoras
If lngNewSol > 0 Then
    Call Bitacora("Aplica", "Copia Solicitud : " & txtNumeroSolicitud.Text & " A la Sol : " & lngNewSol)
    MsgBox "Copia Realizada, NUEVA SOLICITUD GENERADA : " & lngNewSol, vbInformation
Else
    MsgBox "No fue posible realizar la Copia de la Solicitud!", vbExclamation
End If


Call sbLimpiaDatos
txtNumeroSolicitud = ""

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
vModulo = 9

 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub txtNumeroSolicitud_Change()
 sbLimpiaDatos
End Sub

Private Sub txtNumeroSolicitud_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then sbConsulta
End Sub

Private Sub sbConsulta()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Despliega en pantalla datos principales del # solicitud suministrado por el
'               usuario.
'REFERENCIAS:   fxDescribeBanco - (Devuelve la descripcion del Banco al que se giro la
'               solicitud)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

strSQL = "select C.codigo,C.Beneficiario,C.Monto,C.Fecha_Solicitud,C.Tipo,C.Id_Banco" _
       & ",C.cod_unidad,C.cod_concepto,U.descripcion as UnidadDesc,X.descripcion as ConceptoDesc" _
       & ",T.descripcion as TDocumento,B.descripcion as BancoDesc" _
       & " from Tes_Transacciones C inner join CntX_unidades U on C.cod_unidad = U.cod_unidad and cod_Contabilidad = " & GLOBALES.gEnlace _
       & " inner join tes_tipos_doc T on C.tipo = T.tipo" _
       & " inner join tes_conceptos X on C.cod_concepto = X.cod_concepto" _
       & " inner join Tes_Bancos B on C.id_banco = B.id_banco" _
       & " where C.nsolicitud = " & txtNumeroSolicitud

Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  sbLimpiaDatos
  MsgBox "No se encontró esta Solicitud...", vbCritical
Else
  txtCodigo = rs!Codigo
  txtBeneficiario = rs!Beneficiario
  txtMonto = Format(rs!Monto, "Standard")
  txtFecha = Format(rs!fecha_solicitud, "dddd, mmm d yyyy")
  
  txtTipo.Tag = rs!Tipo
  txtTipo.Text = rs!TDOCUMENTO
  txtBanco.Tag = rs!Id_Banco
  txtBanco.Text = rs!BancoDesc

  txtUnidad.Tag = rs!Cod_Unidad
  txtUnidad.Text = rs!UnidadDesc
  
  txtConcepto.Tag = rs!COD_CONCEPTO
  txtConcepto.Text = rs!ConceptoDesc
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub
