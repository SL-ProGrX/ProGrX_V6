VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmIVR_Cat_Cuentas_Contables 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI Cuentas Contables "
   ClientHeight    =   9000
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.FlatEdit txtCentroCostoDesc 
      Height          =   312
      Left            =   4800
      TabIndex        =   54
      Top             =   960
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
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
   Begin XtremeSuiteControls.FlatEdit txtUnidadDesc 
      Height          =   312
      Left            =   4800
      TabIndex        =   51
      Top             =   600
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
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
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   10320
      Top             =   120
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   372
      Left            =   8880
      TabIndex        =   49
      Top             =   60
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Guardar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.GroupBox gbPrimasYDescuentos 
      Height          =   1932
      Left            =   240
      TabIndex        =   36
      Top             =   3600
      Width           =   10932
      _Version        =   1441793
      _ExtentX        =   19283
      _ExtentY        =   3408
      _StockProps     =   79
      Caption         =   "Primas y Descuentos"
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
      Begin XtremeSuiteControls.FlatEdit txtCtaPrima 
         Height          =   312
         Left            =   2760
         TabIndex        =   41
         Top             =   360
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtCtaPrimaDesc 
         Height          =   312
         Left            =   4560
         TabIndex        =   42
         Top             =   360
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
      Begin XtremeSuiteControls.FlatEdit txtCtaPrimaAmort 
         Height          =   312
         Left            =   2760
         TabIndex        =   43
         Top             =   720
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtCtaPrimaAmortDesc 
         Height          =   312
         Left            =   4560
         TabIndex        =   44
         Top             =   720
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
      Begin XtremeSuiteControls.FlatEdit txtCtaDescuentoAmort 
         Height          =   312
         Left            =   2760
         TabIndex        =   45
         Top             =   1440
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtCtaDescuentoAmortDesc 
         Height          =   312
         Left            =   4560
         TabIndex        =   46
         Top             =   1440
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
      Begin XtremeSuiteControls.FlatEdit txtCtaDescuentoDesc 
         Height          =   312
         Left            =   4560
         TabIndex        =   47
         Top             =   1080
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
      Begin XtremeSuiteControls.FlatEdit txtCtaDescuento 
         Height          =   312
         Left            =   2760
         TabIndex        =   48
         Top             =   1080
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Descuento"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   3
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Prima"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "(i) Amortización de descuento"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "(g) Amortización de Prima"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox gbValorizacion 
      Height          =   1932
      Left            =   240
      TabIndex        =   16
      Top             =   5640
      Width           =   10572
      _Version        =   1441793
      _ExtentX        =   18648
      _ExtentY        =   3408
      _StockProps     =   79
      Caption         =   "Valorización"
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
      Begin XtremeSuiteControls.FlatEdit txtCtaValActivo 
         Height          =   312
         Left            =   2760
         TabIndex        =   28
         Top             =   360
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtCtaValActivoDesc 
         Height          =   312
         Left            =   4560
         TabIndex        =   29
         Top             =   360
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
      Begin XtremeSuiteControls.FlatEdit txtCtaValPatrimonio 
         Height          =   312
         Left            =   2760
         TabIndex        =   30
         Top             =   720
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtCtaValPatrimonioDesc 
         Height          =   312
         Left            =   4560
         TabIndex        =   31
         Top             =   720
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
      Begin XtremeSuiteControls.FlatEdit txtCtaValResGasto 
         Height          =   312
         Left            =   2760
         TabIndex        =   32
         Top             =   1080
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtCtaValResGastoDesc 
         Height          =   312
         Left            =   4560
         TabIndex        =   33
         Top             =   1080
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
      Begin XtremeSuiteControls.FlatEdit txtCtaValResIngreso 
         Height          =   312
         Left            =   2760
         TabIndex        =   34
         Top             =   1440
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtCtaValResIngresoDesc 
         Height          =   312
         Left            =   4560
         TabIndex        =   35
         Top             =   1440
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   16
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Resultados Cuenta Ingresos"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Resultados Cuenta Gastos"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   8
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Valor al Patrimonio"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   7
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Valor del Activo"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCtaComisionGstDesc 
      Height          =   312
      Left            =   4800
      TabIndex        =   14
      Top             =   3120
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
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
   Begin XtremeSuiteControls.FlatEdit txtCtaInversion 
      Height          =   312
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtCtaInversionDesc 
      Height          =   312
      Left            =   4800
      TabIndex        =   6
      Top             =   1560
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
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
   Begin XtremeSuiteControls.FlatEdit txtCtaIntAcum 
      Height          =   312
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtCtaIntAcumDesc 
      Height          =   312
      Left            =   4800
      TabIndex        =   8
      Top             =   1920
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
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
   Begin XtremeSuiteControls.FlatEdit txtCtaIntereses 
      Height          =   312
      Left            =   3000
      TabIndex        =   9
      Top             =   2280
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtCtaInteresesDesc 
      Height          =   312
      Left            =   4800
      TabIndex        =   10
      Top             =   2280
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
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
   Begin XtremeSuiteControls.FlatEdit txtCtaIntFluctuacion 
      Height          =   312
      Left            =   3000
      TabIndex        =   11
      Top             =   2760
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtCtaIntFluctuacionDesc 
      Height          =   312
      Left            =   4800
      TabIndex        =   12
      Top             =   2760
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
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
   Begin XtremeSuiteControls.FlatEdit txtCtaComisionGst 
      Height          =   312
      Left            =   3000
      TabIndex        =   13
      Top             =   3120
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1332
      Left            =   240
      TabIndex        =   17
      Top             =   7680
      Width           =   10692
      _Version        =   1441793
      _ExtentX        =   18860
      _ExtentY        =   2350
      _StockProps     =   79
      Caption         =   "Recompras"
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
      Begin XtremeSuiteControls.FlatEdit txtCtaRePrincipal 
         Height          =   312
         Left            =   2760
         TabIndex        =   20
         Top             =   360
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtCtaRePrincipalDesc 
         Height          =   312
         Left            =   4560
         TabIndex        =   21
         Top             =   360
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
      Begin XtremeSuiteControls.FlatEdit txtCtaReIntereses 
         Height          =   312
         Left            =   2760
         TabIndex        =   22
         Top             =   720
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtCtaReInteresesDesc 
         Height          =   312
         Left            =   4560
         TabIndex        =   23
         Top             =   720
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   11
         Left            =   960
         TabIndex        =   19
         Top             =   720
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Intereses"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   10
         Left            =   960
         TabIndex        =   18
         Top             =   360
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Principal"
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtUnidad 
      Height          =   312
      Left            =   3000
      TabIndex        =   50
      Top             =   600
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtCentroCosto 
      Height          =   312
      Left            =   3000
      TabIndex        =   53
      Top             =   960
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   15
      Left            =   480
      TabIndex        =   55
      Top             =   960
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Centro de Costo"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   12
      Left            =   480
      TabIndex        =   52
      Top             =   600
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Unidad (Uen)"
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
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   492
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10932
      _Version        =   1441793
      _ExtentX        =   19283
      _ExtentY        =   868
      _StockProps     =   14
      Caption         =   "Cuentas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   14
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   2892
      _Version        =   1441793
      _ExtentX        =   5101
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "(g) Gasto por Comisiones"
      BackColor       =   -2147483633
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
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   13
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   2892
      _Version        =   1441793
      _ExtentX        =   5101
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "(i) Ingresos por Fluctuaciones"
      BackColor       =   -2147483633
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
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   6
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "(i) Registro de Intereses"
      BackColor       =   -2147483633
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
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   5
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   2652
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Int. Acumulados por Cobrar"
      BackColor       =   -2147483633
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
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1560
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Registro de Inversión"
      BackColor       =   -2147483633
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
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmIVR_Cat_Cuentas_Contables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub btnGuardar_Click()
Call sbGuardar
End Sub

Private Sub Form_Load()

scTitulo.Caption = gIVR_Cuentas.Descripcion

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbConsulta

End Sub


Private Sub sbConsulta()

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "select * from vIVR_CUENTAS" _
       & " Where Tipo_BASE = '" & gIVR_Cuentas.Tipo _
       & "' and CODIGO_01 = '" & gIVR_Cuentas.Codigo_1 _
       & "' and CODIGO_02 = '" & gIVR_Cuentas.Codigo_2 & "'"

Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    txtUnidad.Text = rs!Cod_Unidad & ""
    txtCentroCosto.Text = rs!Cod_Centro_Costo & ""
    
    txtUnidadDesc.Text = rs!Unidad_Desc
    txtCentroCostoDesc.Text = rs!Centro_Desc
    
    txtCtaInversion.Text = rs!CTA_INVERSION_MASK
    txtCtaInversionDesc.Text = rs!CTA_INVERSION_DESC
    
    txtCtaIntAcum.Text = rs!CTA_INTERESES_ACUM_COBRAR_MASK
    txtCtaIntAcumDesc.Text = rs!CTA_INTERESES_ACUM_COBRAR_DESC
    
    txtCtaIntereses.Text = rs!CTA_INGRESOS_INTERESES_MASK
    txtCtaInteresesDesc.Text = rs!CTA_INGRESOS_INTERESES_DESC
    
    txtCtaIntFluctuacion.Text = rs!CTA_INGRESOS_FLUCTUACIONES_MASK
    txtCtaIntFluctuacionDesc.Text = rs!CTA_INGRESOS_FLUCTUACIONES_DESC
    
    txtCtaComisionGst.Text = rs!CTA_GASTO_COMISIONES_MASK
    txtCtaComisionGstDesc.Text = rs!CTA_GASTO_COMISIONES_DESC
    
    txtCtaPrima.Text = rs!CTA_PRIMA_MASK
    txtCtaPrimaDesc.Text = rs!CTA_PRIMA_DESC
    txtCtaPrimaAmort.Text = rs!CTA_PRIMA_GASTO_AMORT_MASK
    txtCtaPrimaAmortDesc.Text = rs!CTA_PRIMA_GASTO_AMORT_DESC
    
    txtCtaDescuento.Text = rs!CTA_DESCUENTOS_MASK
    txtCtaDescuentoDesc.Text = rs!CTA_DESCUENTOS_DESC
    txtCtaDescuentoAmort.Text = rs!CTA_DESCUENTO_INGRESO_MASK
    txtCtaDescuentoAmortDesc.Text = rs!CTA_DESCUENTO_INGRESO_DESC
    
    
    txtCtaValActivo.Text = rs!CTA_VAL_ACTIVO_MASK
    txtCtaValActivoDesc.Text = rs!CTA_VAL_ACTIVO_DESC
    
    txtCtaValPatrimonio.Text = rs!CTA_VAL_PATRIMONIO_MASK
    txtCtaValPatrimonioDesc.Text = rs!CTA_VAL_PATRIMONIO_DESC
    
    txtCtaValResGasto.Text = rs!CTA_VAL_RESULTADOS_GASTOS_MASK
    txtCtaValResGastoDesc.Text = rs!CTA_VAL_RESULTADOS_GASTOS_DESC
    
    txtCtaValResIngreso.Text = rs!CTA_VAL_RESULTADOS_INGRESOS_MASK
    txtCtaValResIngresoDesc.Text = rs!CTA_VAL_RESULTADOS_INGRESOS_DESC
    
    txtCtaReIntereses.Text = rs!CTA_RECOMPRA_INTERESES_MASK
    txtCtaReInteresesDesc.Text = rs!CTA_RECOMPRA_INTERESES_DESC
    
    txtCtaRePrincipal.Text = rs!CTA_RECOMPRA_PRINCIPAL_MASK
    txtCtaRePrincipalDesc.Text = rs!CTA_RECOMPRA_PRINCIPAL_DESC
    
'      , isnull(Cixp.'','') as 'CTA_INTERESES_PAGAR_MASK', isnull(Cixp.DESCRIPCION,'')  as 'CTA_INTERESES_PAGAR_DESC'
'      , isnull(Cgxi.'','') as 'CTA_GASTO_INTERESES_MASK', isnull(Cgxi.DESCRIPCION,'')  as 'CTA_GASTO_INTERESES_DESC'


End If

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub


Private Sub sbGuardar()

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "exec spIVR_CUENTAS_REGISTRA '" & gIVR_Cuentas.Tipo & "','" & gIVR_Cuentas.Codigo_1 & "', '" & gIVR_Cuentas.Codigo_2 & "', '" & glogon.Usuario _
       & "', '" & txtUnidad.Text & "','" & txtCentroCosto.Text _
       & "', '" & txtCtaInversion.Text & "','" & txtCtaIntAcum.Text & "', '" & txtCtaIntereses.Text _
       & "', '" & txtCtaIntFluctuacion.Text & "' ,'" & txtCtaComisionGst.Text _
       & "', '" & txtCtaPrima.Text & "','" & txtCtaPrimaAmort.Text & "', '" & txtCtaDescuento.Text & "', '" & txtCtaDescuentoAmort.Text _
       & "', '" & txtCtaValActivo.Text & "','" & txtCtaValPatrimonio.Text & "', '" & txtCtaValResGasto.Text & "', '" & txtCtaValResIngreso.Text _
       & "', '" & txtCtaRePrincipal.Text & "','" & txtCtaReIntereses.Text _
       & "', '" & txtCtaIntereses.Text & "', '" & txtCtaIntereses.Text & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Información Actualizada Satisfactoriamente!", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub sbCuenta_Consulta(pCuenta As Object, pDesc As Object)
   frmCntX_ConsultaCuentas.Show vbModal
   pCuenta.Text = gCuenta
   pDesc.Text = fxgCntCuentaDesc(gCuenta)
   pCuenta.Text = fxgCntCuentaFormato(True, pCuenta.Text, 0)
End Sub

Private Sub sbCuenta_LostFocus(pCuenta As Object, pDesc As Object)
   
   pDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, pCuenta.Text, 0))
   pCuenta.Text = fxgCntCuentaFormato(True, pCuenta.Text, 0)

End Sub


Private Sub txtCentroCosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_centro_Costo"
  gBusquedas.Consulta = "select cod_centro_Costo as Centro,Descripcion from cntx_centro_costos"
  gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace & " and Activo = 1"
  gBusquedas.Orden = "cod_centro_Costo"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
     txtCentroCosto.Text = gBusquedas.Resultado
     txtCentroCostoDesc.Text = gBusquedas.Resultado2
  End If
End If

End Sub


Private Sub txtCentroCosto_LostFocus()
 txtCentroCostoDesc.Text = fxgCntCentroCostos(txtCentroCosto.Text)
End Sub

Private Sub txtCtaComisionGst_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaPrima.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaComisionGst, txtCtaComisionGstDesc)
End If

End Sub


Private Sub txtCtaComisionGst_LostFocus()
    Call sbCuenta_LostFocus(txtCtaComisionGst, txtCtaComisionGstDesc)
End Sub

Private Sub txtCtaDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaDescuentoAmort.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaDescuento, txtCtaDescuentoDesc)
End If

End Sub



Private Sub txtCtaDescuento_LostFocus()
    Call sbCuenta_LostFocus(txtCtaDescuento, txtCtaDescuentoDesc)
End Sub

Private Sub txtCtaDescuentoAmort_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaValActivo.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaDescuentoAmort, txtCtaDescuentoAmortDesc)
End If

End Sub

Private Sub txtCtaDescuentoAmort_LostFocus()
    Call sbCuenta_LostFocus(txtCtaDescuentoAmort, txtCtaDescuentoAmortDesc)
End Sub

Private Sub txtCtaIntAcum_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaIntereses.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaIntAcum, txtCtaIntAcumDesc)
End If

End Sub


Private Sub txtCtaIntAcum_LostFocus()
    Call sbCuenta_LostFocus(txtCtaIntAcum, txtCtaIntAcumDesc)
End Sub

Private Sub txtCtaIntereses_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaIntFluctuacion.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaIntereses, txtCtaInteresesDesc)
End If

End Sub

Private Sub txtCtaIntereses_LostFocus()
    Call sbCuenta_LostFocus(txtCtaIntereses, txtCtaInteresesDesc)
End Sub


Private Sub txtCtaIntFluctuacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaComisionGst.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaIntFluctuacion, txtCtaIntFluctuacionDesc)
End If

End Sub

Private Sub txtCtaIntFluctuacion_LostFocus()
    Call sbCuenta_LostFocus(txtCtaIntFluctuacion, txtCtaIntFluctuacionDesc)
End Sub

Private Sub txtCtaInversion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaIntAcum.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaInversion, txtCtaInversionDesc)
End If
End Sub

Private Sub txtCtaInversion_LostFocus()
    Call sbCuenta_LostFocus(txtCtaInversion, txtCtaInversionDesc)
End Sub


Private Sub txtCtaPrima_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaPrimaAmort.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaPrima, txtCtaPrimaDesc)
End If

End Sub


Private Sub txtCtaPrima_LostFocus()
    Call sbCuenta_LostFocus(txtCtaPrima, txtCtaPrimaDesc)
End Sub

Private Sub txtCtaPrimaAmort_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaDescuento.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaPrimaAmort, txtCtaPrimaAmortDesc)
End If

End Sub


Private Sub txtCtaPrimaAmort_LostFocus()
    Call sbCuenta_LostFocus(txtCtaPrimaAmort, txtCtaPrimaAmortDesc)
End Sub

Private Sub txtCtaReIntereses_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidad.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaReIntereses, txtCtaReInteresesDesc)
End If

End Sub

Private Sub txtCtaReIntereses_LostFocus()
    Call sbCuenta_LostFocus(txtCtaReIntereses, txtCtaReInteresesDesc)
End Sub

Private Sub txtCtaRePrincipal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaReIntereses.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaRePrincipal, txtCtaRePrincipalDesc)
End If

End Sub

Private Sub txtCtaRePrincipal_LostFocus()
    Call sbCuenta_LostFocus(txtCtaRePrincipal, txtCtaRePrincipalDesc)
End Sub

Private Sub txtCtaValActivo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaValPatrimonio.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaValActivo, txtCtaValActivoDesc)
End If

End Sub


Private Sub txtCtaValActivo_LostFocus()
    Call sbCuenta_LostFocus(txtCtaValActivo, txtCtaValActivoDesc)
End Sub

Private Sub txtCtaValPatrimonio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaValResGasto.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaValPatrimonio, txtCtaValPatrimonioDesc)
End If

End Sub

Private Sub txtCtaValPatrimonio_LostFocus()
    Call sbCuenta_LostFocus(txtCtaValPatrimonio, txtCtaValPatrimonioDesc)
End Sub

Private Sub txtCtaValResGasto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaValResIngreso.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaValResGasto, txtCtaValResGastoDesc)
End If

End Sub

Private Sub txtCtaValResGasto_LostFocus()
    Call sbCuenta_LostFocus(txtCtaValResGasto, txtCtaValResGastoDesc)
End Sub

Private Sub txtCtaValResIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaRePrincipal.SetFocus

If KeyCode = vbKeyF4 Then
    Call sbCuenta_Consulta(txtCtaValResIngreso, txtCtaValResIngresoDesc)
End If

End Sub

Private Sub txtCtaValResIngreso_LostFocus()
    Call sbCuenta_LostFocus(txtCtaInversion, txtCtaInversionDesc)
End Sub

Private Sub txtUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
  
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_unidad"
  gBusquedas.Consulta = "select cod_unidad as Unidad,Descripcion from CntX_Unidades"
  gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
  gBusquedas.Orden = "cod_unidad"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
     txtUnidad.Text = gBusquedas.Resultado
     txtUnidadDesc.Text = gBusquedas.Resultado2
  End If
End If
  
  
End Sub

Private Sub txtUnidad_LostFocus()
 txtUnidadDesc.Text = fxgCntUnidad(txtUnidad.Text)
End Sub
