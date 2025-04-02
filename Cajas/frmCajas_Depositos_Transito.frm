VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmCajas_Depositos_Transito 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activación de Depositos de Cierre de Cajas"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10140
   ScaleWidth      =   16545
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   9975
      _Version        =   1441793
      _ExtentX        =   17595
      _ExtentY        =   12726
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDP_Id 
      Height          =   330
      Left            =   11400
      TabIndex        =   32
      Top             =   4560
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbRegistro 
      Height          =   3375
      Left            =   10320
      TabIndex        =   22
      Top             =   6600
      Width           =   6135
      _Version        =   1441793
      _ExtentX        =   10821
      _ExtentY        =   5953
      _StockProps     =   79
      Caption         =   "Activación"
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
      Begin XtremeSuiteControls.DateTimePicker dtpActiva_Fecha 
         Height          =   330
         Left            =   1560
         TabIndex        =   33
         Top             =   600
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtActiva_Numero 
         Height          =   330
         Left            =   1560
         TabIndex        =   34
         Top             =   1080
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnActivar 
         Height          =   615
         Left            =   3360
         TabIndex        =   35
         Top             =   2520
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Activar Depósito en Bancos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_Depositos_Transito.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Depósito"
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
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha "
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   582
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
      Height          =   330
      Left            =   2520
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   5400
      TabIndex        =   4
      Top             =   1440
      Width           =   3375
      _Version        =   1441793
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.FlatEdit txtMntInicio 
      Height          =   330
      Left            =   5400
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.FlatEdit txtMntCorte 
      Height          =   330
      Left            =   7080
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.FlatEdit txtNumero 
      Height          =   330
      Left            =   1080
      TabIndex        =   12
      Top             =   1920
      Width           =   2895
      _Version        =   1441793
      _ExtentX        =   5106
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   615
      Left            =   9000
      TabIndex        =   13
      Top             =   1680
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCajas_Depositos_Transito.frx":0727
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   615
      Left            =   10320
      TabIndex        =   14
      ToolTipText     =   "Exportar a Excel"
      Top             =   1680
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   1080
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCajas_Depositos_Transito.frx":0E27
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   9000
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   233
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCajaApertura 
      Height          =   330
      Left            =   11400
      TabIndex        =   26
      Top             =   3720
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCajaUsuario 
      Height          =   330
      Left            =   14400
      TabIndex        =   27
      Top             =   3720
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDP_Numero 
      Height          =   330
      Left            =   11400
      TabIndex        =   29
      Top             =   5520
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDP_Monto 
      Height          =   450
      Left            =   11400
      TabIndex        =   30
      Top             =   6000
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   794
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCajaId 
      Height          =   330
      Left            =   11400
      TabIndex        =   25
      Top             =   3240
      Width           =   5055
      _Version        =   1441793
      _ExtentX        =   8916
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDP_Cuenta 
      Height          =   330
      Left            =   11400
      TabIndex        =   28
      Top             =   5040
      Width           =   4935
      _Version        =   1441793
      _ExtentX        =   8705
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   495
      Index           =   1
      Left            =   10080
      TabIndex        =   36
      Top             =   2400
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
      _ExtentY        =   873
      _StockProps     =   14
      Caption         =   "Activar Depósito en Bancos"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   12
      Left            =   10320
      TabIndex        =   31
      Top             =   4560
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Linea Id"
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
      Height          =   255
      Index           =   9
      Left            =   10320
      TabIndex        =   21
      Top             =   6000
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Monto"
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
      Height          =   255
      Index           =   8
      Left            =   10320
      TabIndex        =   20
      Top             =   5520
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Deposito"
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
      Height          =   255
      Index           =   7
      Left            =   10320
      TabIndex        =   19
      Top             =   5040
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cuenta"
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
      Height          =   255
      Index           =   6
      Left            =   13320
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario"
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
      Height          =   255
      Index           =   5
      Left            =   10320
      TabIndex        =   17
      Top             =   3720
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Id AP/CR"
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
      Height          =   255
      Index           =   4
      Left            =   10320
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Caja"
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
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cuenta"
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
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   8
      Top             =   1920
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Rango"
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
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Número"
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
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   735
      _Version        =   1441793
      _ExtentX        =   1296
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fechas"
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   9975
      _Version        =   1441793
      _ExtentX        =   17595
      _ExtentY        =   873
      _StockProps     =   14
      Caption         =   "Lista de Depósitos pendientes de confirmación"
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Activación de Depositos"
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
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   16575
   End
End
Attribute VB_Name = "frmCajas_Depositos_Transito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean




Private Sub btnActivar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass


Me.MousePointer = vbDefault


Call sbConsulta

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBuscar_Click()

Call sbConsulta

End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub





Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub



Private Sub sbConsulta()


Dim pBanco As Long

On Error GoTo vError

Me.MousePointer = vbHourglass


txtDP_Cuenta.Text = ""
txtDP_Id.Text = ""
txtDP_Monto.Text = "0"
txtDP_Numero.Text = ""

txtCajaId.Text = ""
txtCajaId.Tag = ""
txtCajaApertura.Text = "0"
txtCajaUsuario.Text = ""


dtpActiva_Fecha.Value = dtpInicio.Value
txtActiva_Numero.Text = ""

lsw.ListItems.Clear



If cbo.Text = "TODOS" Then
    pBanco = 0
Else
    pBanco = cbo.ItemData(cbo.ListIndex)
End If

txtNumero.Text = fxSysCleanTxtInject(txtNumero.Text)


strSQL = "exec spCajas_Depositos_Transito '" & Format(dtpInicio.Value, "yyyy-mm-dd") & " 00:00:00', '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'" _
        & ", " & pBanco & ", '" & txtNumero.Text & "', " & CCur(txtMntInicio.Text) & ", " & CCur(txtMntCorte.Text)

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!DP_Numero)
      itmX.SubItems(1) = Format(rs!Monto, "Standard")
      itmX.SubItems(2) = Format(rs!Registro_Fecha, "yyyy-mm-dd")
      itmX.SubItems(3) = rs!Cuenta_Id
      itmX.SubItems(4) = rs!Cuenta_Banco
      itmX.SubItems(5) = rs!Cod_Caja
      itmX.SubItems(6) = rs!Caja_Desc
      itmX.SubItems(7) = rs!Cod_Apertura
      itmX.SubItems(8) = rs!Registro_Usuario
      itmX.SubItems(9) = rs!Linea
      itmX.SubItems(10) = rs!Id_Banco
  rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Load()

On Error GoTo vError

vModulo = 5


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)

txtMntCorte.Text = Format(9999999999.99, "Standard")
txtMntInicio.Text = Format(0, "Standard")


With lsw.ColumnHeaders
  .Clear
  .Add , , "No.Depósito", 2000
  .Add , , "Monto", 1800, vbRightJustify
  .Add , , "Fecha", 1800, vbCenter
  .Add , , "Cuenta", 2000
  .Add , , "Banco", 2500
  .Add , , "Caja", 1500, vbCenter
  .Add , , "Caja Desc", 2500
  .Add , , "Apertura", 1500, vbCenter
  .Add , , "Usuario", 2500, vbCenter
  .Add , , "D/Id", 700, vbCenter
  .Add , , "B/Id", 700, vbCenter
End With

vPaso = True

    strSQL = "exec spCajas_DepositosCuentasBancarias"
    Call sbCbo_Llena_New(cbo, strSQL, True, True)

vPaso = False


Call Formularios(Me)
Call RefrescaTags(Me)



Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)




'With lsw.ColumnHeaders
'  .Clear
'  .Add , , "No.Depósito", 2500
'1  .Add , , "Monto", 2500, vbRightJustify
'2  .Add , , "Fecha", 1800, vbCenter
'3  .Add , , "Cuenta", 2500
'4  .Add , , "Banco", 2500
'5  .Add , , "Caja", 1500, vbCenter
'6  .Add , , "Caja Desc", 2500
'7  .Add , , "Apertura", 1500, vbCenter
'8  .Add , , "Usuario", 2500, vbCenter
'9  .Add , , "D/Id", 500, vbCenter
'10  .Add , , "B/Id", 500, vbCenter
'End With



txtDP_Cuenta.Tag = Item.SubItems(10)
txtDP_Cuenta.Text = Item.SubItems(3)


txtCajaId.Tag = Item.SubItems(5)
txtCajaId.Text = Item.SubItems(6)
txtCajaApertura.Text = Item.SubItems(7)
txtCajaUsuario.Text = Item.SubItems(8)


txtDP_Id.Text = Item.SubItems(9)
txtDP_Monto.Text = Item.SubItems(1)
txtDP_Numero.Text = Item.Text

dtpActiva_Fecha.Value = Item.SubItems(2)
txtActiva_Numero.Text = Item.Text


End Sub

Private Sub txtMntCorte_GotFocus()
On Error GoTo vError

txtMntCorte.Text = CCur(txtMntCorte.Text)

Exit Sub

vError:
End Sub

Private Sub txtMntCorte_LostFocus()
On Error GoTo vError

txtMntCorte.Text = Format(CCur(txtMntCorte.Text), "Standard")

Exit Sub

vError:
End Sub

Private Sub txtMntInicio_GotFocus()
On Error GoTo vError

 txtMntInicio.Text = CCur(txtMntInicio.Text)

Exit Sub

vError:
End Sub

Private Sub txtMntInicio_LostFocus()
On Error GoTo vError

 txtMntInicio.Text = Format(CCur(txtMntInicio.Text), "Standard")

Exit Sub

vError:
End Sub
