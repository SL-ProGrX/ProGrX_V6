VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmTES_Auto_Registro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de Transacciones con Auto Registro"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   2775
      Left            =   0
      TabIndex        =   40
      Top             =   6840
      Width           =   10335
      _Version        =   1441793
      _ExtentX        =   18230
      _ExtentY        =   4895
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
      Item(0).Caption =   "Beneficiario"
      Item(0).ControlCount=   8
      Item(0).Control(0)=   "chkInd_Referencia"
      Item(0).Control(1)=   "cboTipos"
      Item(0).Control(2)=   "txtBeneficiarioNombre"
      Item(0).Control(3)=   "txtBeneficiarioId"
      Item(0).Control(4)=   "Label1(12)"
      Item(0).Control(5)=   "Label1(11)"
      Item(0).Control(6)=   "Label1(10)"
      Item(0).Control(7)=   "ShortcutCaption1"
      Item(1).Caption =   "Filtro por Cta Bancaria"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "chkFiltraCtasBancos"
      Item(1).Control(1)=   "txtFiltraCtas"
      Item(1).Control(2)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   1695
         Left            =   -69880
         TabIndex        =   50
         Top             =   840
         Visible         =   0   'False
         Width           =   10095
         _Version        =   1441793
         _ExtentX        =   17806
         _ExtentY        =   2990
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkInd_Referencia 
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         Top             =   960
         Width           =   7935
         _Version        =   1441793
         _ExtentX        =   13996
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Indica Información de Persona de Referencia"
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
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboTipos 
         Height          =   330
         Left            =   2400
         TabIndex        =   42
         Top             =   1320
         Width           =   1935
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
      Begin XtremeSuiteControls.FlatEdit txtBeneficiarioNombre 
         Height          =   315
         Left            =   2400
         TabIndex        =   43
         Top             =   2280
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   556
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBeneficiarioId 
         Height          =   315
         Left            =   2400
         TabIndex        =   44
         Top             =   1800
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkFiltraCtasBancos 
         Height          =   255
         Left            =   -69880
         TabIndex        =   49
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Filtra cuentas bancarias"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltraCtas 
         Height          =   330
         Left            =   -65440
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   5655
         _Version        =   1441793
         _ExtentX        =   9975
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   48
         Top             =   360
         Width           =   10335
         _Version        =   1441793
         _ExtentX        =   18230
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Opcional: Indica el Cliente de Referencia "
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   47
         Top             =   1320
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo de Id"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   46
         Top             =   1800
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificación"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   45
         Top             =   2280
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Beneficiario/Referencia"
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
   Begin XtremeSuiteControls.FlatEdit txtCentroDesc 
      Height          =   330
      Left            =   4320
      TabIndex        =   21
      Top             =   4680
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtUnidadDesc 
      Height          =   330
      Left            =   4320
      TabIndex        =   19
      Top             =   4200
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
      Height          =   330
      Left            =   4320
      TabIndex        =   17
      Top             =   3720
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   330
      Left            =   2400
      TabIndex        =   11
      Top             =   1800
      Width           =   7695
      _Version        =   1441793
      _ExtentX        =   13573
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Activo"
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
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
   Begin XtremeSuiteControls.FlatEdit txtPalabraClave 
      Height          =   330
      Left            =   2400
      TabIndex        =   12
      Top             =   2280
      Width           =   7695
      _Version        =   1441793
      _ExtentX        =   13573
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDetalle 
      Height          =   330
      Left            =   2400
      TabIndex        =   13
      Top             =   2760
      Width           =   7695
      _Version        =   1441793
      _ExtentX        =   13573
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtConceptoDesc 
      Height          =   330
      Left            =   4320
      TabIndex        =   15
      Top             =   3240
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtConcepto 
      Height          =   330
      Left            =   2400
      TabIndex        =   14
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   3240
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtCuenta 
      Height          =   330
      Left            =   2400
      TabIndex        =   16
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   3720
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtUnidad 
      Height          =   330
      Left            =   2400
      TabIndex        =   18
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   4200
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtCentro 
      Height          =   330
      Left            =   2400
      TabIndex        =   20
      Top             =   4680
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtMntInicio 
      Height          =   330
      Left            =   2400
      TabIndex        =   22
      Top             =   5280
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMntCorte 
      Height          =   330
      Left            =   4320
      TabIndex        =   23
      Top             =   5280
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkAPL_CargasDiarias 
      Height          =   255
      Left            =   2400
      TabIndex        =   26
      Top             =   6120
      Width           =   3855
      _Version        =   1441793
      _ExtentX        =   6800
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cargas Diarias de Movimientos en Bancos"
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
   Begin XtremeSuiteControls.CheckBox chkAPL_Conciliacion 
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   6480
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Conciliación Bancaria"
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
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   6840
      TabIndex        =   28
      ToolTipText     =   "Nuevo"
      Top             =   1200
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_Auto_Registro.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   7920
      TabIndex        =   29
      ToolTipText     =   "Editar"
      Top             =   1200
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_Auto_Registro.frx":0632
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   8280
      TabIndex        =   30
      ToolTipText     =   "Eliminar"
      Top             =   1200
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_Auto_Registro.frx":0C2D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   8880
      TabIndex        =   31
      ToolTipText     =   "Guardar"
      Top             =   1200
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_Auto_Registro.frx":11D1
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   9240
      TabIndex        =   32
      ToolTipText     =   "Deshacer"
      Top             =   1200
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_Auto_Registro.frx":1902
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   9720
      TabIndex        =   33
      ToolTipText     =   "Reporte"
      Top             =   1200
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_Auto_Registro.frx":2002
      ImageAlignment  =   6
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   4440
      TabIndex        =   34
      Top             =   1250
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboTipoMov 
      Height          =   330
      Left            =   8400
      TabIndex        =   35
      Top             =   5280
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.ComboBox cboTipoDoc 
      Height          =   330
      Left            =   2400
      TabIndex        =   37
      Top             =   5640
      Width           =   3855
      _Version        =   1441793
      _ExtentX        =   6800
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
   Begin XtremeSuiteControls.CheckBox chkIgnoraRegistro 
      Height          =   255
      Left            =   6480
      TabIndex        =   39
      ToolTipText     =   "Filtra Casos con la Regla para No registra en Bancos"
      Top             =   6120
      Width           =   2895
      _Version        =   1441793
      _ExtentX        =   5106
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ignorar el Registro en Bancos"
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
   Begin XtremeSuiteControls.CheckBox chkInd_Control_Depositos 
      Height          =   255
      Left            =   6480
      TabIndex        =   52
      Top             =   6480
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Control de Depósito?"
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
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   38
      Top             =   5640
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo de Registro"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   13
      Left            =   6480
      TabIndex        =   36
      Top             =   5280
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo de Movimiento"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   9
      Left            =   240
      TabIndex        =   25
      Top             =   6120
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Aplicar Regla de Auto Registro en ?"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   495
      Left            =   2400
      TabIndex        =   24
      Top             =   240
      Width           =   8055
      _Version        =   1441793
      _ExtentX        =   14208
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Configuración de Auto-Registro"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Palabra Clave"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Montos Autorizados"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Centro de Costo"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Unidad"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cuenta "
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Concepto"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Detalle "
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Descripción"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Id Auto Registro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_Auto_Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long
Dim vScroll As Boolean, vPaso As Boolean

Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem



Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub btnBarra_Click(Index As Integer)



Select Case Index
    Case 0 'NUEVO
        vEdita = False
        Call sbLimpiaPantalla
        txtDescripcion.SetFocus

        Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
        vEdita = True
        txtDescripcion.SetFocus
        
        Call sbBarra_Accion("Editar")
      
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
      Call sbBarra_Accion("Editar")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
    
    Case 5 'REPORTES
   
End Select


End Sub




Private Sub cboTipoMov_Click()
If vPaso Then Exit Sub

strSQL = "exec spTes_Tipos_Docs '" & Mid(cboTipoMov.Text, 1, 1) & "'"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)

End Sub

Private Sub FlatScrollBar_Change()
On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 ID_AUTO from vTES_AUTO_REGISTRO"

    If Not IsNumeric(txtCodigo.Text) Then txtCodigo.Text = "0"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where ID_AUTO > " & txtCodigo.Text & " order by ID_AUTO asc"
    Else
       strSQL = strSQL & " where ID_AUTO < " & txtCodigo.Text & " order by ID_AUTO desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!ID_Auto
      Call sbConsulta(txtCodigo.Text)
    End If

End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()


On Error GoTo vError

vModulo = 9
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
vPaso = True
 
cboTipos.Clear
cboTipos.AddItem "Personas"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(1)
cboTipos.AddItem "Bancos"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(2)
cboTipos.AddItem "Proveedores"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(3)
cboTipos.AddItem "Acreedores"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(4)
cboTipos.AddItem "Cuentas"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(5)
cboTipos.AddItem "Empleados"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(6)
cboTipos.AddItem "Directos"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(7)
 
 
cboTipoMov.Clear
cboTipoMov.AddItem "Débitos"
cboTipoMov.ItemData(cboTipos.ListCount - 1) = "D"
cboTipoMov.AddItem "Créditos"
cboTipoMov.ItemData(cboTipos.ListCount - 1) = "C"
 
 
 
vPaso = False
 
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Divisa", 1000, vbCenter
    .Add , , "Cuenta/IBAN", 2100, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Contabilidad", 2100, vbCenter
End With
 
 
 
 vEdita = False

 Call sbBarra_Accion("Nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation


End Sub




Private Sub sbLimpiaPantalla()

vCodigo = 0
txtCodigo.Text = "0"

txtDescripcion.Text = ""
txtPalabraClave.Text = ""
txtDetalle.Text = ""
txtConcepto.Text = ""
txtConceptoDesc.Text = ""
txtCuenta.Text = ""
txtCuentaDesc.Text = ""
txtUnidad.Text = ""
txtUnidadDesc.Text = ""
txtCentro.Text = ""
txtCentroDesc.Text = ""

txtMntInicio.Text = "0"
txtMntCorte.Text = Format(999999999999.99, "Standard")


chkActivo.Value = xtpChecked

chkAPL_CargasDiarias.Value = xtpChecked
chkAPL_Conciliacion.Value = xtpChecked

chkInd_Control_Depositos.Value = xtpUnchecked

chkInd_Referencia.Value = xtpUnchecked
chkIgnoraRegistro.Value = xtpUnchecked
chkFiltraCtasBancos.Value = xtpUnchecked

cboTipos.Text = "Personas"

cboTipoMov.Text = "Débitos"

lsw.ListItems.Clear

tcMain.Item(0).Selected = True

txtBeneficiarioId.Text = ""
txtBeneficiarioNombre.Text = ""

'txtDescripcion.SetFocus

End Sub



Private Sub sbConsulta(xCodigo As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vTES_AUTO_REGISTRO where ID_Auto = " & xCodigo
Call OpenRecordSet(rs, strSQL)

tcMain.Item(0).Selected = True

If Not rs.BOF And Not rs.EOF Then
   Call sbBarra_Accion("activo")
  vEdita = True
  
  vCodigo = rs!ID_Auto
  txtCodigo = rs!ID_Auto
 
  txtDescripcion.Text = rs!Descripcion
  txtDescripcion.SetFocus
    
    

    txtPalabraClave.Text = rs!PALABRAS_CLAVE
    txtDetalle.Text = rs!Detalle
    txtConcepto.Text = rs!cod_Concepto
    txtConceptoDesc.Text = rs!Concepto_Desc
    txtCuenta.Text = rs!Cod_Cuenta_Mask
    txtCuentaDesc.Text = rs!Cuenta_Desc
    txtUnidad.Text = rs!Cod_Unidad
    txtUnidadDesc.Text = rs!Unidad_Desc
    txtCentro.Text = rs!Cod_Centro_Costo
    txtCentroDesc.Text = rs!Centro_Desc
    
    txtMntInicio.Text = Format(rs!Mnt_Inicio, "Standard")
    txtMntCorte.Text = Format(rs!Mnt_Corte, "Standard")
    
    
    chkActivo.Value = rs!Activo
    
    chkAPL_CargasDiarias.Value = rs!APL_CARGA_DIARIA
    chkAPL_Conciliacion.Value = rs!APL_CONCILIACION
    
    
    chkInd_Referencia.Value = rs!IND_INFO_PERSONA
    
    txtBeneficiarioId.Text = rs!BENEFICIARIO_ID
    txtBeneficiarioNombre.Text = rs!BENEFICIARIO_NOMBRE

    cboTipos.ListIndex = IIf(IsNull(rs!Tipo_Beneficiario), 0, rs!Tipo_Beneficiario - 1)
  
    chkIgnoraRegistro.Value = rs!IGNORA_REGISTRO_ID
    chkFiltraCtasBancos.Value = rs!FILTRA_CTA_BANCOS
    
    Call sbCboAsignaDato(cboTipoMov, rs!APL_TIPO_MOV_DESC, True, rs!APL_TIPO_MOV)
    
    Call sbCboAsignaDato(cboTipoDoc, rs!TIPO_DOC_DESC, True, rs!TIPO_DOC_ID)
    
    
    chkInd_Control_Depositos.Value = rs!DP_TRAMITE
Else
  If vEdita Then
      MsgBox "No se encontró registro verifique...", vbInformation
  End If
End If

rs.Close
Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Limpia Inyección
Dim Ctrl As Control

For Each Ctrl In Me
  If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is XtremeSuiteControls.FlatEdit Then
        Ctrl.Text = fxSysCleanTxtInject(Ctrl.Text)
  End If
Next Ctrl

If Not IsNumeric(txtCodigo.Text) Then vMensaje = vMensaje & vbCrLf & " - El ID del Auto Registro no es válido!"


'Validar Cuentas Aqui

If txtDescripcion.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Descripció no es válida ..."
If txtPalabraClave.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Detalle no es válido ..."
If txtDetalle.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No ha indicado un detalle..."


If txtConceptoDesc.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No ha indicado un Concepto..."
If txtUnidadDesc.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No ha indicado una Unidad..."
If txtCuentaDesc.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No ha indicado una Cuenta Contable..."

If Not IsNumeric(txtMntInicio.Text) Then vMensaje = vMensaje & vbCrLf & " - Monto de Inicio no es válido!"
If Not IsNumeric(txtMntCorte.Text) Then vMensaje = vMensaje & vbCrLf & " - Monto de Corte no es válido!"

If chkInd_Referencia.Value = xtpChecked Then
 
    If txtBeneficiarioId.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No ha indicado la Identificación del Beneficiario"
    If txtBeneficiarioNombre.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No ha indicado el Nombre del Beneficiario"
    
End If



If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()



On Error GoTo vError

strSQL = "exec spTes_Auto_Registro_Add " & txtCodigo.Text & ", '" & txtDescripcion.Text & "', '" & txtPalabraClave.Text _
        & "', '" & txtDetalle.Text & "', '" & txtConcepto.Text & "', '" & fxgCntCuentaFormato(False, txtCuenta.Text, 0) & "', '" & txtUnidad.Text _
        & "', '" & txtCentro.Text & "', " & CCur(txtMntInicio.Text) & ", " & CCur(txtMntCorte.Text) _
        & ", " & chkAPL_CargasDiarias.Value & ", " & chkAPL_Conciliacion.Value & ", " & chkInd_Referencia.Value _
        & ", " & cboTipos.ItemData(cboTipos.ListIndex) & ", '" & txtBeneficiarioId.Text _
        & "', '" & txtBeneficiarioNombre.Text & "', " & chkActivo.Value & ", '" & glogon.Usuario & "', 'A" _
        & "', '" & Mid(cboTipoMov.Text, 1, 1) & "', '" & cboTipoDoc.ItemData(cboTipoDoc.ListIndex) _
        & "',  " & chkIgnoraRegistro.Value & ", " & chkFiltraCtasBancos.Value
            
Call OpenRecordSet(rs, strSQL)
            
txtCodigo.Text = rs!Auto_Id
vCodigo = rs!Auto_Id
            
If rs!Result = 1 Then
   Call Bitacora("Registra", "Auto Registro Id: " & vCodigo & "..: " & txtDescripcion.Text)
End If

If rs!Result = 2 Then
   Call Bitacora("Modifica", "Auto Registro Id: " & vCodigo & "..: " & txtDescripcion.Text)
End If


MsgBox "Información guardada satisfactoriamente...", vbInformation

 Call sbBarra_Accion("activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  
    strSQL = "exec spTes_Auto_Registro_Add " & txtCodigo.Text & ", '" & txtDescripcion.Text & "', '" & txtPalabraClave.Text _
            & "', '" & txtDetalle.Text & "', '" & txtConcepto.Text & "', '" & fxgCntCuentaFormato(False, txtCuenta.Text, 0) & "', '" & txtUnidad.Text _
            & "', '" & txtCentro.Text & "', " & CCur(txtMntInicio.Text) & ", " & CCur(txtMntCorte.Text) _
            & ", " & chkAPL_CargasDiarias.Value & ", " & chkAPL_Conciliacion.Value & ", " & chkInd_Referencia.Value _
            & ", " & cboTipos.ItemData(cboTipos.ListIndex) & ", '" & txtBeneficiarioId.Text _
            & "', '" & txtBeneficiarioNombre.Text & "', " & chkActivo.Value & ", '" & glogon.Usuario & "', 'E'"
                
    Call OpenRecordSet(rs, strSQL)
                
    If rs!Result = 3 Then
       Call Bitacora("Elimina", "Auto Registro Id: " & vCodigo & "..: " & txtDescripcion.Text)
    End If
  
  Call sbLimpiaPantalla
   Call sbBarra_Accion("nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCta01.SetFocus
'
'If KeyCode = vbKeyF4 Then
'  gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
'  gBusquedas.Filtro = " and COD_CONTABILIDAD = " & GLOBALES.gEnlace
'  gBusquedas.Columna = "tipo_asiento"
'  gBusquedas.Orden = "tipo_asiento"
'  frmBusquedas.Show vbModal
'  txtAsiento = gBusquedas.Resultado
'  txtAsientoDesc.Text = gBusquedas.Resultado2
'End If
End Sub

Private Sub sbCtas_Bancos_Load()

On Error GoTo vError

vPaso = True

txtFiltraCtas.Text = fxSysCleanTxtInject(txtFiltraCtas.Text)

vPaso = True
lsw.ListItems.Clear

strSQL = "exec spTes_Auto_Registro_Ctas " & txtCodigo.Text & ", '" & txtFiltraCtas.Text & "'"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Id_Banco)
      itmX.SubItems(1) = rs!cod_Divisa & ""
      itmX.SubItems(2) = rs!Cta & ""
      itmX.SubItems(3) = rs!Descripcion & ""
      itmX.SubItems(4) = rs!Cod_Cuenta_Mask & ""
      
      itmX.Checked = IIf((rs!asignado = 1), True, False)
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
    vPaso = False
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spTes_Auto_Registro_Ctas_Add " & txtCodigo.Text & ", " & Item.Text _
    & ",'" & IIf(Item.Checked, "A", "E") & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
  
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
   If txtCodigo.Text <> "" Then
        txtFiltraCtas.Text = ""
        Call sbCtas_Bancos_Load
   End If
End If

End Sub

Private Sub txtBeneficiarioId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtBeneficiarioNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  Select Case cboTipos.ItemData(cboTipos.ListIndex)
      Case 1 'Personas
        gBusquedas.Col1Name = "Cedula"
        gBusquedas.Col2Name = "Nombre"
        gBusquedas.Col3Name = "Id Alterno"
        gBusquedas.Columna = "cedula"
        gBusquedas.Orden = "cedula"
        gBusquedas.Consulta = "select cedula,nombre, cedular from socios"
      Case 2 'Bancos
        gBusquedas.Columna = "id_banco"
        gBusquedas.Orden = "id_banco"
        gBusquedas.Consulta = "select id_banco,descripcion from tes_bancos"
      Case 3 'PRoveedores
        gBusquedas.Columna = "cedjur"
        gBusquedas.Orden = "cedjur"
        gBusquedas.Consulta = "select cedjur,cod_proveedor,descripcion from cxp_proveedores"
     Case 4 'Acreedores
        gBusquedas.Columna = "cod_acreedor"
        gBusquedas.Orden = "cod_acreedor"
        gBusquedas.Consulta = "select cod_acreedor,descripcion from crd_apa_acreedores"
  
     Case 5 'Cuentas por Cobrar
        gBusquedas.Columna = "cedula"
        gBusquedas.Orden = "cedula"
        gBusquedas.Consulta = "select cedula,nombre from CXC_PERSONAS"
     
     Case 6 'Empleados
        gBusquedas.Columna = "Identificacion"
        gBusquedas.Orden = "Identificacion"
        gBusquedas.Consulta = "select Identificacion,Nombre_Completo from RH_PERSONAS"
     
     Case 7 'Directos
        gBusquedas.Columna = "Codigo"
        gBusquedas.Orden = "Codigo"
        gBusquedas.Consulta = "select Codigo,Beneficiario from vTes_Beneficiarios"
     
  End Select
  
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtBeneficiarioId.Text = gBusquedas.Resultado
  If txtBeneficiarioId <> "" Then txtBeneficiarioNombre.SetFocus
End If

gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""

End Sub

Private Sub txtBeneficiarioId_LostFocus()

On Error GoTo vError

Me.MousePointer = vbHourglass



Select Case cboTipos.ItemData(cboTipos.ListIndex)
   Case 1 'Personas
    If Trim(txtBeneficiarioId) <> "" Then
       strSQL = "Select Cedula,Nombre" _
              & " from Socios Where Cedula='" & txtBeneficiarioId & "'"
       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiarioNombre.Text = Trim(rs!Nombre & "")
       Else
          txtBeneficiarioNombre.Text = ""
       End If
    Else
       txtBeneficiarioNombre.Text = ""
    End If

Case 2 'Bancos
    If Trim(txtBeneficiarioId) <> "" Then
       strSQL = "select ID_BANCO,descripcion from TES_BANCOS where ID_BANCO  =" & txtBeneficiarioId & ""
       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiarioNombre.Text = Trim(rs!Descripcion)
       Else
          txtBeneficiarioNombre.Text = ""
       End If

    Else
       txtBeneficiarioNombre.Text = ""
    End If

Case 3 'Proveedores
    If Trim(txtBeneficiarioId) <> "" Then
       strSQL = "select CEDJUR, DESCRIPCION" _
              & " from CXP_PROVEEDORES where CEDJUR = '" & txtBeneficiarioId & "'"
       Call OpenRecordSet(rs, strSQL)
       
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiarioNombre.Text = Trim(rs!Descripcion & "")
       Else
          txtBeneficiarioNombre.Text = ""
       End If

    Else
       txtBeneficiarioNombre.Text = ""
    End If

Case 5 'Cuentas por Cobrar
    If Trim(txtBeneficiarioId) <> "" Then
       strSQL = "select Cod_Acreedor, DESCRIPCION  from CRD_APA_ACREEDORES where cod_acreedor = " & txtBeneficiarioId & ""
       Call OpenRecordSet(rs, strSQL)
       
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiarioNombre.Text = Trim(rs!Descripcion)
       Else
          txtBeneficiarioNombre.Text = ""
       End If

    Else
       txtBeneficiarioNombre.Text = ""
    End If


Case 6 'Empleados
    If Trim(txtBeneficiarioId) <> "" Then
       strSQL = "Select IDENTIFICACION, NOMBRE_COMPLETO" _
              & " from RH_PERSONAS Where IDENTIFICACION='" & txtBeneficiarioId & "'"
              
              
       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiarioNombre.Text = Trim(rs!NOMBRE_COMPLETO & "")
       Else
          txtBeneficiarioNombre.Text = ""
       End If

    Else
       txtBeneficiarioNombre.Text = ""
       
    End If


Case 7 'Directos
    If Trim(txtBeneficiarioId) <> "" Then
       strSQL = "Select CODIGO, BENEFICIARIO" _
              & " from vTes_Beneficiarios Where CODIGO ='" & txtBeneficiarioId & "'"
              
              
       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiarioNombre.Text = Trim(rs!Beneficiario & "")
       End If

    Else
       txtBeneficiarioNombre.Text = ""
       
    End If



End Select


Me.MousePointer = vbDefault

Exit Sub
vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCentro_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentroDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Centro"
  gBusquedas.Col2Name = "Descripción"
  gBusquedas.Col3Name = ""
  gBusquedas.Consulta = "select COD_CENTRO_COSTO, DESCRIPCION from vCNTX_CENTRO_COSTO_LOCAL"
  gBusquedas.Filtro = ""
  gBusquedas.Columna = "COD_CENTRO_COSTO"
  gBusquedas.Orden = "DESCRIPCION"
  frmBusquedas.Show vbModal
  
  If gBusquedas.Resultado <> "" Then
     txtCentro.Text = gBusquedas.Resultado
     txtCentroDesc.Text = gBusquedas.Resultado2
  End If
  
End If

End Sub

Private Sub txtCentro_LostFocus()
Call sbCodigoDescripcion("Cc", txtCentro.Text)

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "ID_Auto"
  gBusquedas.Orden = "ID_Auto"
  gBusquedas.Consulta = "select ID_Auto ,Descripcion from vTES_AUTO_REGISTRO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  
  If IsNumeric(gBusquedas.Resultado) Then
    txtCodigo.Text = gBusquedas.Resultado
    txtDescripcion.SetFocus
  End If
End If

End Sub

Private Sub txtCodigo_LostFocus()

If IsNumeric(txtCodigo.Text) Then
  If CLng(txtCodigo.Text) > 0 Then
        Call sbConsulta(txtCodigo.Text)
   End If
End If

End Sub


Private Sub sbCodigoDescripcion(pTipo As String, pCodigo As String)

On Error GoTo vError

Dim txt As XtremeSuiteControls.FlatEdit

Select Case pTipo
  Case "Cta"
  Case "Con"
    strSQL = "select DESCRIPCION as 'ItmX' from vTes_Conceptos Where cod_concepto = '" & pCodigo & "'"
    Set txt = txtConceptoDesc
    
  Case "Ud"
    strSQL = "select DESCRIPCION as 'ItmX' from vCNTX_UNIDADES_LOCAL Where cod_Unidad = '" & pCodigo & "'"
    Set txt = txtUnidadDesc
  Case "Cc"
    strSQL = "select DESCRIPCION as 'ItmX' from vCNTX_CENTRO_COSTO_LOCAL Where cod_Centro_Costo = '" & pCodigo & "'"
    Set txt = txtCentroDesc
End Select


Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    txt.Text = rs!itmX
End If
rs.Close


vError:
  
End Sub


Private Sub txtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConceptoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Concepto"
  gBusquedas.Col2Name = "Descripción"
  gBusquedas.Col3Name = "Cuenta"
  gBusquedas.Consulta = "select COD_CONCEPTO, DESCRIPCION, COD_CUENTA_MASK, DP_TRAMITE_APL from vTes_Conceptos"
  gBusquedas.Filtro = " AND AUTO_REGISTRO = 1 AND ESTADO = 'A'"
  gBusquedas.Columna = "COD_CONCEPTO"
  gBusquedas.Orden = "DESCRIPCION"
  frmBusquedas.Show vbModal
  
  If gBusquedas.Resultado <> "" Then
     txtConcepto.Text = gBusquedas.Resultado
     txtConceptoDesc.Text = gBusquedas.Resultado2
     
     If txtCuenta.Text = "" Then
       txtCuenta.Text = gBusquedas.Resultado3
     End If
  End If
  
End If

End Sub

Private Sub txtConcepto_LostFocus()

On Error GoTo vError

strSQL = "select COD_CONCEPTO, DESCRIPCION, COD_CUENTA_MASK, DP_TRAMITE_APL, CUENTA_DESC" _
       & " from vTes_Conceptos Where cod_Concepto = '" & txtConcepto.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtConceptoDesc.Text = rs!Descripcion
    txtCuenta.Text = rs!Cod_Cuenta_Mask
    txtCuentaDesc.Text = rs!Cuenta_Desc
    chkInd_Control_Depositos.Value = rs!DP_TRAMITE_APL

End If

Exit Sub

vError:

End Sub


Private Sub txtConceptoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidad.SetFocus
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuenta.Text = gCuenta
   txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta, 0)
End If

End Sub

Private Sub txtCuenta_LostFocus()
   txtCuentaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuenta, 0))
   txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta, 0)
End Sub




Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPalabraClave.SetFocus
End Sub


Private Sub txtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConcepto.SetFocus
End Sub


Private Sub txtFiltraCtas_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call sbCtas_Bancos_Load
End If

End Sub

Private Sub txtMntInicio_GotFocus()
On Error GoTo vError
  
  txtMntInicio.Text = CCur(txtMntInicio.Text)

vError:
End Sub

Private Sub txtMntInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMntCorte.SetFocus

End Sub

Private Sub txtMntInicio_LostFocus()
On Error GoTo vError
  
  txtMntInicio.Text = Format(CCur(txtMntInicio.Text), "Standard")

vError:

End Sub

Private Sub txtMntCorte_GotFocus()
On Error GoTo vError
  
  txtMntCorte.Text = CCur(txtMntCorte.Text)

vError:
End Sub

Private Sub txtMntCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkAPL_CargasDiarias.SetFocus

End Sub

Private Sub txtMntCorte_LostFocus()
On Error GoTo vError
  
  txtMntCorte.Text = Format(CCur(txtMntCorte.Text), "Standard")

vError:

End Sub


Private Sub txtPalabraClave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub txtUnidad_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidadDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Unidad"
  gBusquedas.Col2Name = "Descripción"
  gBusquedas.Col3Name = ""
  gBusquedas.Consulta = "select COD_UNIDAD, DESCRIPCION from vCNTX_UNIDADES_LOCAL"
  gBusquedas.Filtro = ""
  gBusquedas.Columna = "COD_UNIDAD"
  gBusquedas.Orden = "DESCRIPCION"
  frmBusquedas.Show vbModal
  
  If gBusquedas.Resultado <> "" Then
     txtUnidad.Text = gBusquedas.Resultado
     txtUnidadDesc.Text = gBusquedas.Resultado2
  End If
  
End If


End Sub

Private Sub txtUnidad_LostFocus()

Call sbCodigoDescripcion("Ud", txtUnidad.Text)

End Sub

Private Sub txtUnidadDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentro.SetFocus
End Sub
