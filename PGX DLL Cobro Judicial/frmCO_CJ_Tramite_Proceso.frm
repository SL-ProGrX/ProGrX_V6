VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCO_CJ_Tramite_Proceso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Proceso en el que se encuentra el Trámite"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox fraSentencia 
      Height          =   1695
      Left            =   5040
      TabIndex        =   24
      Top             =   3600
      Visible         =   0   'False
      Width           =   4095
      _Version        =   1310723
      _ExtentX        =   7223
      _ExtentY        =   2990
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.FlatEdit txtMontoSentencia 
         Height          =   315
         Left            =   1680
         TabIndex        =   26
         Top             =   600
         Width           =   2175
         _Version        =   1310723
         _ExtentX        =   3836
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   960
         Width           =   2175
         _Version        =   1310723
         _ExtentX        =   3836
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   29
         Top             =   960
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   27
         Top             =   600
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   4095
         _Version        =   1310723
         _ExtentX        =   7223
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Datos de la Sentencia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   555
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   979
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
      Text            =   "0000"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLineaCod 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLineaDesc 
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   5655
      _Version        =   1310723
      _ExtentX        =   9975
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      Top             =   1920
      Width           =   5655
      _Version        =   1310723
      _ExtentX        =   9975
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   555
      Left            =   3480
      TabIndex        =   5
      Top             =   960
      Width           =   5655
      _Version        =   1310723
      _ExtentX        =   9975
      _ExtentY        =   979
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTramite 
      Height          =   555
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   979
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0000"
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   405
      Left            =   1680
      TabIndex        =   7
      Top             =   3120
      Width           =   7470
      _Version        =   1310723
      _ExtentX        =   13176
      _ExtentY        =   714
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.CheckBox chkCierreSentencia 
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   3720
      Width           =   5895
      _Version        =   1310723
      _ExtentX        =   10398
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Aplicar el Cierre del Proceso con Sentencia ?"
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
   Begin XtremeSuiteControls.CheckBox chkHonorario 
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   5280
      Width           =   7215
      _Version        =   1310723
      _ExtentX        =   12726
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tramitar Pago de Honorarios (Registro para el pago y generación del gasto)"
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
   Begin XtremeSuiteControls.FlatEdit txtHonorarioMnt 
      Height          =   315
      Left            =   1680
      TabIndex        =   16
      Top             =   6000
      Width           =   2175
      _Version        =   1310723
      _ExtentX        =   3836
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbogado 
      Height          =   315
      Left            =   1680
      TabIndex        =   17
      Top             =   5640
      Width           =   6855
      _Version        =   1310723
      _ExtentX        =   12091
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   915
      Left            =   1680
      TabIndex        =   20
      Top             =   4200
      Width           =   7455
      _Version        =   1310723
      _ExtentX        =   13150
      _ExtentY        =   1614
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
   Begin XtremeSuiteControls.GroupBox gbAplicar 
      Height          =   975
      Left            =   0
      TabIndex        =   22
      Top             =   7320
      Width           =   9375
      _Version        =   1310723
      _ExtentX        =   16536
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   495
         Left            =   7560
         TabIndex        =   23
         Top             =   360
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2566
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
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCO_CJ_Tramite_Proceso.frx":0000
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   330
      Left            =   1680
      TabIndex        =   30
      Top             =   6360
      Width           =   6870
      _Version        =   1310723
      _ExtentX        =   12118
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.ComboBox cboCuenta 
      Height          =   330
      Left            =   1680
      TabIndex        =   31
      Top             =   6720
      Width           =   3990
      _Version        =   1310723
      _ExtentX        =   7038
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.ComboBox cboEmite 
      Height          =   330
      Left            =   6600
      TabIndex        =   32
      Top             =   6720
      Width           =   1950
      _Version        =   1310723
      _ExtentX        =   3440
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   35
      Top             =   6360
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Banco"
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
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   34
      Top             =   6720
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cuenta"
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
      Height          =   255
      Index           =   10
      Left            =   5880
      TabIndex        =   33
      Top             =   6720
      Width           =   855
      _Version        =   1310723
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Emite"
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
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   21
      Top             =   4200
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Notas"
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
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   19
      Top             =   5640
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Abogado"
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
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   18
      Top             =   6000
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Monto"
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
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   960
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "No. Operación"
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
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Línea Crédito"
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
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
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
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Tramite Id:"
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
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   9255
      _Version        =   1310723
      _ExtentX        =   16325
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Actualización del Proceso"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
      Height          =   495
      Index           =   5
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Proceso Actual ..:"
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
Attribute VB_Name = "frmCO_CJ_Tramite_Proceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mUltimoProceso As String

Dim vPaso As Boolean
Dim vAbogado As String, vAbogadoId As Long, vAbogadoCedula As String
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub cboBanco_Click()
If vPaso Or cboBanco.ListCount = 0 Or cboBanco.Text = "" Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & vAbogadoCedula & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub

Private Sub cboProceso_Click()
If vPaso Then Exit Sub

Call sbDatosProceso(cboProceso.ItemData(cboProceso.ListIndex))

End Sub

Private Sub btnAplicar_Click()

Dim pEmite As String
Dim pGasto As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If Trim(txtNotas) = "" Or Len(txtNotas) < 10 Then
  MsgBox "Debe incluir una nota valida...", vbInformation
  Me.MousePointer = vbDefault
  Exit Sub
End If

If chkCierreSentencia.Value = vbChecked Then
    If Val(txtMontoSentencia) > 0 Then
        strSQL = "Update cbr_cj_tramite set SENTENCIA_FECHA = '" & Format(dtpFecha.Value, "yyyymmdd") _
               & "',SENTENCIA_MONTO = " & CCur(txtMontoSentencia.Text) _
               & " where cod_tramite = " & txtTramite
    Else
      MsgBox "Monto de sentencia invalido...", vbInformation
      Me.MousePointer = vbDefault
      Exit Sub
    End If
End If


If cboEmite.Text = "Transferencia" Then
    pEmite = "TE"
Else
    pEmite = "CK"
End If

If cboCuenta.ListCount = 0 Then
   pEmite = "CK"
End If



strSQL = "INSERT CBR_CJ_TRAMITE_PROCESO(NUM_LINEA,COD_TRAMITE,COD_PROCESO,NOTAS,APLICA_CIERRE_SENTENCIA," _
       & " REGISTRO_FECHA ,REGISTRO_USUARIO) Values(" & fxLineaTramite(txtTramite) & "," & txtTramite.Text & "," _
       & " '" & cboProceso.ItemData(cboProceso.ListIndex) & "','" & txtNotas.Text & "'," & chkCierreSentencia.Value _
       & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"
Call ConectionExecute(strSQL)

'Registro del Gasto
If CCur(txtHonorarioMnt.Text) > 0 Then
    pGasto = fxCBR_CJ_Parametros("01")
    
    strSQL = "exec spCBR_CJ_CargaGastos " & txtTramite.Text & " ,'" & pGasto & "'," & CCur(txtHonorarioMnt.Text) & "," _
                                    & " '','" & txtNotas & "','" & vAbogado & "'," _
                                    & " '" & glogon.Usuario & "'," & chkHonorario.Value & ", '" & vAbogadoCedula & "'," & vAbogadoId _
                                    & ", " & cboBanco.ItemData(cboBanco.ListIndex) & ", '" & pEmite & "', '" & cboCuenta.ItemData(cboCuenta.ListIndex) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If rs!NumDoc <> "" Then
        Call sbImprimeRecibo(rs!NumDoc, rs!TipoDoc, False)
    End If
    
    rs.Close

End If


MsgBox "Proceso guardado satisfactoriamente...", vbInformation

fraSentencia.Visible = False

Me.MousePointer = vbDefault
UnLoad Me

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Me.MousePointer = vbDefault
End Sub

Private Sub chkCierreSentencia_Click()
If chkCierreSentencia.Value = vbChecked Then
   fraSentencia.Visible = True
   dtpFecha.Value = fxFechaServidor
   txtMontoSentencia.Text = 0
   txtMontoSentencia.SetFocus
Else
  fraSentencia.Visible = False
End If
End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNotas.SetFocus
End Sub

Private Sub Form_Load()

Dim strSQL As String


vPaso = True

txtTramite.Text = GLOBALES.gTag

txtHonorarioMnt.Text = 0

cboEmite.Clear
cboEmite.AddItem "Transferencia"
cboEmite.AddItem "Cheque"
cboEmite.Text = "Transferencia"

strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)


Call sbDatosIniciales

mUltimoProceso = Trim(fxOrdenProceso(txtTramite))

strSQL = "select cod_proceso as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  cbr_cj_proceso where activo = 1 and orden between '" & mUltimoProceso _
         & "' and '" & fxOrdenProcesoSeguiente(mUltimoProceso) & "'"
         
Call sbCbo_Llena_New(cboProceso, strSQL, False, True)

vPaso = False

Call cboBanco_Click

End Sub

Private Sub sbDatosIniciales()

strSQL = "exec spCbr_CJ_Tramite_Consulta_Lite " & txtTramite.Text
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
      
    txtOperacion.Text = rs!Id_Solicitud
    
    txtLineaCod.Text = rs!Codigo
    txtLineaDesc.Text = rs!LineaDesc
    
    txtCedula.Text = RTrim(rs!Cedula)
    txtNombre.Text = rs!Nombre
    
    txtProceso.Text = rs!Proceso_Ult_Desc
    
    If rs!BuffeteId > 0 Then
        vAbogado = rs!BuffeteDesc
        vAbogadoCedula = rs!BuffeteCedula
        vAbogadoId = rs!BuffeteId
    Else
        vAbogado = rs!AbogadoDesc
        vAbogadoCedula = rs!AbogadoCedula
        vAbogadoId = rs!AbogadoId
    End If
    
End If
rs.Close


txtAbogado.Text = vAbogado

End Sub

Private Function fxOrdenProceso(vTramite As Long) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select orden from CBR_CJ_PROCESO where COD_PROCESO " _
        & "in(select MAX(cod_proceso) from CBR_CJ_TRAMITE_PROCESO Where cod_tramite = " & vTramite & " )"

Call OpenRecordSet(rs, strSQL)
fxOrdenProceso = rs!Orden
rs.Close
End Function


Private Function fxOrdenProcesoSeguiente(vOrden As String) As String
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select min(orden) as 'Orden' from CBR_CJ_PROCESO  where Orden > '" & Trim(vOrden) & "'"
Call OpenRecordSet(rs, strSQL)
fxOrdenProcesoSeguiente = rs!Orden & ""
rs.Close
End Function




Private Function fxLineaTramite(vTramite As Long) As Integer
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(max(num_linea),0)+1 as 'Linea'" _
        & " from cbr_cj_tramite_proceso where cod_tramite =" & vTramite & ""
Call OpenRecordSet(rs, strSQL)
fxLineaTramite = rs!Linea
rs.Close
End Function

Private Sub sbDatosProceso(vProceso As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "SELECT isnull(HONORARIOS_APLICA,0) as 'HONORARIOS_APLICA',isnull(HONORARIOS_MONTO,0) as 'HONORARIOS_MONTO' " _
        & " FROM CBR_CJ_PROCESO WHERE COD_PROCESO = '" & vProceso & "'"
Call OpenRecordSet(rs, strSQL)

If rs!HONORARIOS_APLICA = 1 Then
    chkHonorario.Value = vbChecked
    chkHonorario.Enabled = True
    txtHonorarioMnt.Enabled = True
    txtHonorarioMnt.Text = Format(rs!HONORARIOS_MONTO, "Standard")
Else
    chkHonorario.Value = vbUnchecked
    chkHonorario.Enabled = False
    txtHonorarioMnt.Enabled = False
    txtHonorarioMnt.Text = 0
End If
rs.Close
End Sub

Private Sub txtHonorarioMnt_GotFocus()
txtHonorarioMnt.Text = CCur(txtHonorarioMnt)
End Sub

Private Sub txtHonorarioMnt_LostFocus()
txtHonorarioMnt.Text = Format(txtHonorarioMnt, "Standard")
End Sub



Private Function fxUltimoProceso()
Dim strSQL As String, rs As New ADODB.Recordset
strSQL = "select descripcion from CBR_CJ_PROCESO where COD_PROCESO in( " _
       & " select  MAX(cod_proceso) from CBR_CJ_TRAMITE_PROCESO " _
       & " where COD_TRAMITE = " & txtTramite & " )"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
    fxUltimoProceso = rs!Descripcion
Else
    fxUltimoProceso = ""
End If
rs.Close

End Function


Private Sub txtMontoSentencia_Change()
If Not IsNumeric(txtMontoSentencia) Then
   MsgBox "Debe incluir un monto valido...", vbInformation
   txtMontoSentencia = 0
End If
End Sub

Private Sub txtMontoSentencia_GotFocus()
txtMontoSentencia.Text = CCur(txtMontoSentencia.Text)
End Sub

Private Sub txtMontoSentencia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  txtMontoSentencia.Text = Format(txtMontoSentencia.Text, "Standard")
  dtpFecha.SetFocus
End If
End Sub
