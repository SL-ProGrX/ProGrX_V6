VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCxC_CuentasSGTRebajoCRD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rebajo de Operaciónes de Crédito"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswRefunde 
      Height          =   2295
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   10215
      _Version        =   1310723
      _ExtentX        =   18013
      _ExtentY        =   4043
      _StockProps     =   77
      BackColor       =   -2147483643
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.GroupBox fraRefunde 
      Height          =   6495
      Left            =   10080
      TabIndex        =   1
      Top             =   1800
      Width           =   10215
      _Version        =   1310723
      _ExtentX        =   18018
      _ExtentY        =   11456
      _StockProps     =   79
      Caption         =   "Datos de la Refundición"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.RadioButton optX 
         Height          =   255
         Index           =   0
         Left            =   7680
         TabIndex        =   9
         Top             =   3720
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cancelación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnCerrar 
         Height          =   492
         Left            =   8040
         TabIndex        =   7
         Top             =   5040
         Width           =   1212
         _Version        =   1310723
         _ExtentX        =   2138
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Cerrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCxC_CuentasSGTRebajoCRD.frx":0000
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnRefunde 
         Height          =   492
         Left            =   4920
         TabIndex        =   8
         Top             =   5040
         Width           =   1572
         _Version        =   1310723
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Refunde"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCxC_CuentasSGTRebajoCRD.frx":07CD
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.RadioButton optX 
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   10
         Top             =   4080
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuotas Pendientes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnActualizar 
         Height          =   492
         Left            =   6480
         TabIndex        =   11
         Top             =   5040
         Width           =   1452
         _Version        =   1310723
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Actualizar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCxC_CuentasSGTRebajoCRD.frx":0FA5
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   315
         Left            =   1320
         TabIndex        =   23
         Top             =   720
         Width           =   1575
         _Version        =   1310723
         _ExtentX        =   2773
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
         _Version        =   1310723
         _ExtentX        =   2773
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIntCor 
         Height          =   315
         Left            =   4800
         TabIndex        =   27
         Top             =   1560
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtIntMor 
         Height          =   315
         Left            =   4800
         TabIndex        =   28
         Top             =   1920
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtAmortizacion 
         Height          =   315
         Left            =   4800
         TabIndex        =   29
         Top             =   1200
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtCargos 
         Height          =   315
         Left            =   4800
         TabIndex        =   30
         Top             =   2280
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtPolizas 
         Height          =   315
         Left            =   4800
         TabIndex        =   31
         Top             =   2640
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   315
         Left            =   4800
         TabIndex        =   32
         Top             =   720
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAbono 
         Height          =   315
         Left            =   4800
         TabIndex        =   41
         Top             =   4440
         Width           =   2055
         _Version        =   1310723
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Cargos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   3120
         TabIndex        =   38
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Principal"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   37
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Int.Moratorio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   36
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Int.Corriente"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   35
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Pólizas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   3120
         TabIndex        =   34
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   33
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Línea"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   495
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Index           =   2
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   10212
         _Version        =   1310723
         _ExtentX        =   18013
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Datos de la Refundición o Abono a la operación:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.93
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   6
      End
      Begin VB.Label Label2 
         Caption         =   "Abono .....:"
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
         Index           =   10
         Left            =   3120
         TabIndex        =   6
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Poner al día"
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
         Index           =   11
         Left            =   3120
         TabIndex        =   5
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label lblMora 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   4800
         TabIndex        =   4
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Cancelación"
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
         Index           =   12
         Left            =   3120
         TabIndex        =   3
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label lblCancelacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   4800
         TabIndex        =   2
         Top             =   3720
         Width           =   2055
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   8880
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_CuentasSGTRebajoCRD.frx":1932
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_CuentasSGTRebajoCRD.frx":4DC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.CheckBox chkCtasActivas 
      Height          =   255
      Left            =   8160
      TabIndex        =   12
      Top             =   960
      Width           =   2295
      _Version        =   1310723
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Incluir Cuotas Activas?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3855
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   10215
      _Version        =   1310723
      _ExtentX        =   18013
      _ExtentY        =   6794
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
      Item(0).Caption =   "Propias"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lswPrestamos"
      Item(1).Caption =   "Terceros"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "lswTerceros"
      Item(1).Control(1)=   "txtConCedula"
      Item(1).Control(2)=   "txtConNombre"
      Item(1).Control(3)=   "Label1(26)"
      Begin XtremeSuiteControls.ListView lswPrestamos 
         Height          =   3372
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   10212
         _Version        =   1310723
         _ExtentX        =   18013
         _ExtentY        =   5948
         _StockProps     =   77
         BackColor       =   -2147483643
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswTerceros 
         Height          =   3012
         Left            =   -70000
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   10212
         _Version        =   1310723
         _ExtentX        =   18013
         _ExtentY        =   5313
         _StockProps     =   77
         BackColor       =   -2147483643
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtConCedula 
         Height          =   312
         Left            =   -67000
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtConNombre 
         Height          =   312
         Left            =   -65320
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1310723
         _ExtentX        =   9758
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Identificación"
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
         Index           =   26
         Left            =   -68200
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   1212
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Disponible:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   40
      Top             =   1400
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDisponible 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Height          =   315
      Left            =   8160
      TabIndex        =   39
      Top             =   1400
      Width           =   2055
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   10215
      _Version        =   1310723
      _ExtentX        =   18013
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Refundiciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   10215
      _Version        =   1310723
      _ExtentX        =   18013
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Operaciones activas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Refundición de Créditos"
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
      Height          =   372
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   5412
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmCxC_CuentasSGTRebajoCRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type OpARefundir
  Operacion As Long
  Saldo     As Currency
  Principal As Currency
  IntCor    As Currency
  IntMor    As Currency
  Cargos    As Currency
  Polizas   As Currency
End Type

Dim mRefunde As OpARefundir, mIngresosTotales As Currency
Dim mMonto As Currency, mRebajosTotales As Currency, mOperacion As Long, mCedula As String


Private Sub btnActualizar_Click()
Dim strSQL As String

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "exec spCxC_TraCrdRefActualiza " & mOperacion & "," & chkCtasActivas.Value
Call ConectionExecute(strSQL)

Call Form_Load
Call LimpiaDatos(False)


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnCerrar_Click()
    Call LimpiaDatos(False)
End Sub

Private Sub btnRefunde_Click()
    Call sbRefunde
End Sub

Private Sub chkCtasActivas_Click()
If ssTab.Tab = 0 Then
    Call sbCargaPrestamos
Else
    If txtConCedula.Text <> "" Then
        Call sbCargaLswTerceros(txtConCedula)
    End If
End If

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset


Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

mOperacion = GLOBALES.gTag

fraRefunde.Visible = False

With lswRefunde.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 2000
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Garantía", 1400, vbCenter
    .Add , , "Descripción", 3500
    .Add , , "Saldo", 1800, vbRightJustify
    .Add , , "Int.Cor.", 1800, vbRightJustify
    .Add , , "Int.Mor.", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "Pólizas", 1800, vbRightJustify
    .Add , , "Principal", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
    .Add , , "Tipo", 1800, vbCenter
End With


With lswPrestamos.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 2000
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Garantía", 1400, vbCenter
    .Add , , "Descripción", 3500
    .Add , , "Saldo", 1800, vbRightJustify
    .Add , , "Int.Cor.", 1800, vbRightJustify
    .Add , , "Int.Mor.", 1800, vbRightJustify
    .Add , , "Principal", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "Pólizas", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
    .Add , , "Tipo", 1800, vbCenter
End With

With lswTerceros.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 2000
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Garantía", 1400, vbCenter
    .Add , , "Descripción", 3500
    .Add , , "Saldo", 1800, vbRightJustify
    .Add , , "Int.Cor.", 1800, vbRightJustify
    .Add , , "Int.Mor.", 1800, vbRightJustify
    .Add , , "Principal", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "Pólizas", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
    .Add , , "Tipo", 1800, vbCenter
End With



tcMain.Item(0).Selected = True


strSQL = "Select isnull(dbo.fxCxC_CuentaRebajos(" & mOperacion & ",'TOT'),0) as 'Rebajos', Monto,cedula" _
       & ", isnull(dbo.fxCxC_CuentaIngresos(" & mOperacion & "),0) as 'Ingresos'" _
       & " from CxC_Cuentas Where Operacion = " & mOperacion
Call OpenRecordSet(rs, strSQL)
   mRebajosTotales = rs!Rebajos
   mIngresosTotales = rs!Ingresos
   mMonto = rs!Monto
   mCedula = Trim(rs!Cedula)
rs.Close

Me.Caption = "Operación : " & mOperacion
lblDisponible.Caption = Format(mMonto + mIngresosTotales - mRebajosTotales, "Standard")

Call sbCargaRebajos
Call sbCargaPrestamos

End Sub

Private Sub sbCargaRebajos()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

strSQL = "select R.*,X.codigo,C.descripcion,G.descripcion as GarantiaX" _
       & " from CxC_Cuentas_Rebajos_Crd R inner join reg_creditos X on R.id_solicitud = X.id_solicitud" _
       & " inner join Catalogo C on X.codigo = C.codigo" _
       & " inner join crd_garantia_tipos G on X.garantia = G.garantia" _
       & " where R.Operacion = " & mOperacion
Call OpenRecordSet(rs, strSQL, 0)
With lswRefunde
  .ListItems.Clear
  Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , rs!Id_Solicitud)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = rs!GarantiaX
     itmX.SubItems(3) = rs!Descripcion
     itmX.SubItems(4) = Format(rs!Monto, "Standard")
     itmX.SubItems(5) = Format(rs!Saldo, "Standard")
     itmX.SubItems(6) = Format(IIf(IsNull(rs!Int_Cor), 0, rs!Int_Cor), "Standard")
     itmX.SubItems(7) = Format(IIf(IsNull(rs!Int_Mor), 0, rs!Int_Mor), "Standard")
     itmX.SubItems(8) = Format(IIf(IsNull(rs!Cargos), 0, rs!Cargos), "Standard")
     itmX.SubItems(9) = Format(IIf(IsNull(rs!Poliza), 0, rs!Poliza), "Standard")
     itmX.SubItems(10) = Format(IIf(IsNull(rs!Principal), 0, rs!Principal), "Standard")
     itmX.SubItems(11) = rs!CTA_PENDIENTES
   rs.MoveNext
  Loop
End With
rs.Close

End Sub

Private Sub sbCargaPrestamos()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

strSQL = "exec spCrdSGTListaCreditosPersona '" & mCedula & "','N'," & chkCtasActivas.Value
Call OpenRecordSet(rs, strSQL, 0)

With lswPrestamos
  .ListItems.Clear
  Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , rs!Id_Solicitud)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = rs!GarantiaX
        itmX.SubItems(3) = rs!Descripcion
        itmX.SubItems(4) = Format(rs!Saldo, "Standard")
        itmX.SubItems(5) = Format(rs!IntC, "Standard")
        itmX.SubItems(6) = Format(rs!IntM, "Standard")
        itmX.SubItems(7) = Format(rs!Amortiza, "Standard")
        itmX.SubItems(8) = Format(rs!Cargos, "Standard")
        itmX.SubItems(9) = Format(rs!Polizas, "Standard")
   rs.MoveNext
  Loop
End With
rs.Close

End Sub


Private Function fxExisteRefundicion(vOperacion As Long) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from CxC_Cuentas_Rebajos_Crd" _
       & " where id_solicitud = " & vOperacion & " and Operacion = " & mOperacion
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
  fxExisteRefundicion = IIf((rs!Existe = 0), False, True)
rs.Close
End Function

Private Sub LimpiaDatos(Optional vVisible As Boolean = True)

mRefunde.Principal = 0
mRefunde.IntCor = 0
mRefunde.IntMor = 0
mRefunde.Saldo = 0
mRefunde.Cargos = 0
mRefunde.Operacion = 0

lblCodigo.Caption = ""
lblOperacion.Caption = ""
txtSaldo.Text = ""
txtIntCor.Text = ""
txtIntMor.Text = ""
txtPrincipal.Text = ""

lblCancelacion.Caption = 0
lblMora.Caption = 0

txtAbono.Text = "0"


If vVisible Then
   fraRefunde.Visible = vVisible
   fraRefunde.top = 960
Else
   fraRefunde.Visible = vVisible
End If

End Sub





Private Function fxValidaRefundicion() As Boolean
Dim vMensaje As String

fxValidaRefundicion = True
vMensaje = ""

If mRefunde.Operacion = 0 Then vMensaje = vMensaje & "- No se ha seleccionado ninguna operación"

If IsNumeric(txtAbono.Text) Then
 If CCur(txtAbono.Text) > CCur(lblDisponible.Caption) Then vMensaje = vMensaje & vbCrLf & "- El abono es mayor que el disponible"
 If CCur(txtAbono.Text) < 0 Then vMensaje = vMensaje & vbCrLf & "- El abono no es válido"
Else
  vMensaje = vMensaje & vbCrLf & "- El saldo no es válido"
End If

If Len(vMensaje) > 0 Then
 fxValidaRefundicion = False
 MsgBox vMensaje, vbCritical
End If

End Function


Private Sub lswPrestamos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError
   
Call LimpiaDatos(True)

lblOperacion.Caption = Item.Text
lblCodigo.Caption = Item.SubItems(1)
txtAbono.Text = "0"

txtSaldo.Text = Format(CCur(Item.SubItems(4)), "Standard")
txtIntCor.Text = Format(CCur(Item.SubItems(5)), "Standard")
txtIntMor.Text = Format(CCur(Item.SubItems(6)), "Standard")
txtPrincipal.Text = Format(CCur(Item.SubItems(7)), "Standard")
txtCargos.Text = Format(CCur(Item.SubItems(8)), "Standard")
txtPoliza.Text = Format(CCur(Item.SubItems(9)), "Standard")

mRefunde.Operacion = lblOperacion.Caption
mRefunde.Principal = txtPrincipal.Text
mRefunde.IntCor = txtIntCor.Text
mRefunde.IntMor = txtIntMor.Text
mRefunde.Saldo = txtSaldo.Text
mRefunde.Cargos = txtCargos.Text
mRefunde.Polizas = txtPoliza.Text

lblCancelacion.Caption = Format(CCur(txtSaldo.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) _
             + CCur(txtCargos.Text) + CCur(txtPoliza.Text), "Standard")
lblMora.Caption = Format(CCur(txtPrincipal.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text) + CCur(txtPoliza.Text), "Standard")
  
fraRefunde.Visible = True
fraRefunde.top = tcMain.top
fraRefunde.Left = 120

Call OptX_Click(0)
   

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswRefunde_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError


strSQL = "delete CxC_Cuentas_Rebajos_Crd where id_solicitud = " & Item.Text _
       & " and Operacion = " & mOperacion
Call ConectionExecute(strSQL)

lblDisponible.Caption = CCur(lblDisponible.Caption) + CCur(Item.SubItems(4))
lblDisponible.Caption = Format(lblDisponible, "Standard")


Call sbCargaRebajos

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswTerceros_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError


   Call LimpiaDatos(True)

   lblOperacion.Caption = Item.Text
   lblCodigo.Caption = Item.SubItems(1)
   txtSaldo.Text = Format(CCur(Item.SubItems(4)), "Standard")
   txtIntCor.Text = Format(CCur(Item.SubItems(5)), "Standard")
   txtIntMor.Text = Format(CCur(Item.SubItems(6)), "Standard")
   txtPrincipal.Text = Format(CCur(Item.SubItems(7)), "Standard")
   txtCargos.Text = Format(CCur(Item.SubItems(8)), "Standard")
   txtPoliza.Text = Format(CCur(Item.SubItems(9)), "Standard")
   
   
   lblCancelacion.Caption = Format(CCur(txtSaldo.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) _
                + CCur(txtCargos.Text) + CCur(txtPoliza.Text), "Standard")
   lblMora.Caption = Format(CCur(txtPrincipal.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text) + CCur(txtPoliza.Text), "Standard")
   
   mRefunde.Operacion = lblOperacion.Caption
   mRefunde.Principal = txtPrincipal.Text
   mRefunde.IntCor = txtIntCor.Text
   mRefunde.IntMor = txtIntMor.Text
   mRefunde.Saldo = txtSaldo.Text
   mRefunde.Cargos = txtCargos.Text
   mRefunde.Polizas = txtPoliza.Text
   
   fraRefunde.Visible = True
   fraRefunde.top = tcMain.top
   fraRefunde.Left = 120
      
   Call OptX_Click(0)
   

vError:
End Sub



Private Sub OptX_Click(Index As Integer)
On Error GoTo vError

    Select Case True
      Case optX.Item(0).Value 'Cancelacion
        If CCur(lblDisponible.Caption) >= CCur(lblCancelacion.Caption) Then
            txtAbono.Text = lblCancelacion.Caption
        Else
            txtAbono.Text = lblDisponible.Caption
        End If
      Case optX.Item(1).Value 'Mora
        If CCur(lblDisponible.Caption) >= CCur(lblMora.Caption) Then
            txtAbono.Text = lblMora.Caption
        Else
            txtAbono.Text = lblDisponible.Caption
        End If
    End Select
    
'    txtAbono.SetFocus
    
Exit Sub

vError:

End Sub



Private Sub sbRefunde()
Dim strSQL As String, curRefundir As Currency

On Error GoTo vError

If fxValidaRefundicion Then

curRefundir = CCur(txtAbono.Text)

If curRefundir > CCur(lblDisponible.Caption) Then
  MsgBox "El monto a refundir de la operación es mayor al disponible...", vbCritical
  Exit Sub
End If

If fxExisteRefundicion(lblOperacion.Caption) Then
  MsgBox "Esta Refundición Se encuentra Registrada VERIFIQUE...", vbInformation
  Exit Sub
Else
  strSQL = "insert CxC_Cuentas_Rebajos_Crd(Operacion,id_solicitud,Monto,Int_Cor,Int_Mor,Principal,cargos,Saldo,Poliza,CTA_PENDIENTES) " _
         & "values(" & mOperacion & "," & lblOperacion.Caption & "," & CCur(txtAbono.Text) & "," & CCur(txtIntCor.Text) _
         & "," & CCur(txtIntMor.Text) & "," & CCur(txtPrincipal.Text) & "," & CCur(txtCargos.Text) & "," & CCur(txtSaldo.Text) _
         & "," & CCur(txtPoliza.Text) & "," & chkCtasActivas.Value & ")"
  Call ConectionExecute(strSQL)
  
  lblDisponible.Caption = CCur(lblDisponible.Caption) - CCur(txtAbono.Text)
  lblDisponible.Caption = Format(lblDisponible, "Standard")
  
  Call sbCargaRebajos
  Call LimpiaDatos(False)
  
End If

End If 'Verificacion de OPERACION

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaLswTerceros(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem


strSQL = "exec spCrdSGTListaCreditosPersona '" & vCedula & "','N'," & chkCtasActivas.Value
Call OpenRecordSet(rs, strSQL, 0)
With lswTerceros
  .ListItems.Clear
  Do While Not rs.EOF
'    txtConNombre = rs!Nombre
    Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!Id_Solicitud)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = rs!GarantiaX
        itmX.SubItems(3) = rs!Descripcion
        itmX.SubItems(4) = Format(rs!Saldo, "Standard")
        itmX.SubItems(5) = Format(rs!IntC, "Standard")
        itmX.SubItems(6) = Format(rs!IntM, "Standard")
        itmX.SubItems(7) = Format(rs!Amortiza, "Standard")
        itmX.SubItems(8) = Format(rs!Cargos, "Standard")
        itmX.SubItems(9) = Format(rs!Polizas, "Standard")
        itmX.Tag = itmX.Index
   rs.MoveNext
  Loop
End With
rs.Close

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
 Case 1
   Call sbCargaLswTerceros(txtConCedula)
 Case Else
End Select

End Sub

Private Sub txtConCedula_Change()
lswTerceros.ListItems.Clear
End Sub

Private Sub txtConCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    txtConNombre = fxNombre(txtConCedula)
    Call sbCargaLswTerceros(txtConCedula)
End If

End Sub




