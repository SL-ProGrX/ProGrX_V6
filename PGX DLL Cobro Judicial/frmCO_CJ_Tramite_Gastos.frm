VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCO_CJ_Tramite_Gastos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Gastos "
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox gbAplicar 
      Height          =   855
      Left            =   120
      TabIndex        =   25
      Top             =   7200
      Width           =   9375
      _Version        =   1310723
      _ExtentX        =   16536
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   495
         Left            =   7560
         TabIndex        =   26
         Top             =   240
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
         Picture         =   "frmCO_CJ_Tramite_Gastos.frx":0000
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.RadioButton optDesembolso 
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   21
      Top             =   5520
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkDesembolso 
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   5040
      Width           =   2175
      _Version        =   1310723
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Aplica Desembolso ?"
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
      TabIndex        =   4
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
      TabIndex        =   5
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
      TabIndex        =   3
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
      TabIndex        =   9
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
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Top             =   3600
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
   Begin XtremeSuiteControls.FlatEdit txtDocumento 
      Height          =   315
      Left            =   6960
      TabIndex        =   13
      Top             =   3600
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   915
      Left            =   1680
      TabIndex        =   17
      Top             =   3960
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
   Begin XtremeSuiteControls.ComboBox cboGasto 
      Height          =   405
      Left            =   1680
      TabIndex        =   19
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
   Begin XtremeSuiteControls.RadioButton optDesembolso 
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   22
      Top             =   5880
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Otro..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtAbogado 
      Height          =   315
      Left            =   2280
      TabIndex        =   23
      Top             =   5520
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
   Begin XtremeSuiteControls.FlatEdit txtBeneficiario 
      Height          =   315
      Left            =   2280
      TabIndex        =   24
      Top             =   5880
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
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   330
      Left            =   2280
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
      Left            =   2280
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
      Left            =   7200
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
   Begin XtremeSuiteControls.FlatEdit txtDesembolsoId 
      Height          =   315
      Left            =   7080
      TabIndex        =   33
      Top             =   5160
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   34
      Top             =   5160
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Id Beneficiario"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   10
      Left            =   6480
      TabIndex        =   29
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
      Index           =   9
      Left            =   720
      TabIndex        =   28
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
      Index           =   8
      Left            =   720
      TabIndex        =   27
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
      Index           =   7
      Left            =   360
      TabIndex        =   18
      Top             =   3960
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
      Left            =   5160
      TabIndex        =   16
      Top             =   3600
      Width           =   1575
      _Version        =   1310723
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Documento"
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
      Index           =   5
      Left            =   360
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Gasto"
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
      TabIndex        =   14
      Top             =   3600
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
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   9255
      _Version        =   1310723
      _ExtentX        =   16325
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Registro de Gasto"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   8
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
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
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
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   6
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
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   360
      X2              =   10080
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frmCO_CJ_Tramite_Gastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim vAbogado As String, vAbogadoId As Long, vAbogadoCedula As String
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub cboBanco_Click()
If vPaso Or cboBanco.ListCount = 0 Or cboBanco.Text = "" Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & txtDesembolsoId.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub

Private Sub cboGasto_Click()

If vPaso Then Exit Sub



strSQL = "select isnull(APLICA_DESEMBOLSO,0)as 'APLICA_DESEMBOLSO', isnull(monto,0) as 'Monto' " _
        & " from CBR_CJ_TIPOS_GASTOS where TIPO_GASTO = '" & cboGasto.ItemData(cboGasto.ListIndex) & "' "
Call OpenRecordSet(rs, strSQL)

If rs!Aplica_Desembolso = 1 Then
    chkDesembolso.Value = vbChecked
    chkDesembolso.Enabled = True
    
    If rs!Monto > 0 Then
       txtMonto = Format(rs!Monto, "Standard")
    End If
    optDesembolso(0).Enabled = True
    optDesembolso(1).Enabled = True
Else
    chkDesembolso.Value = vbUnchecked
    chkDesembolso.Enabled = False
    optDesembolso(0).Enabled = False
    optDesembolso(1).Enabled = False
    
End If

End Sub

Private Sub btnAplicar_Click()

Dim pBeneficiario As String, pBeneficiarioId As Long, pEmite As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If optDesembolso(0).Value = True Then
   pBeneficiario = txtAbogado.Text
   pBeneficiarioId = vAbogadoId
Else
   pBeneficiario = txtBeneficiario.Text
   pBeneficiarioId = 0
End If

If cboEmite.Text = "Transferencia" Then
    pEmite = "TE"
Else
    pEmite = "CK"
End If

If cboCuenta.ListCount = 0 Then
   pEmite = "CK"
End If


If Trim(pBeneficiario) = "" Then
  MsgBox "Debe incluir un beneficiario...", vbInformation
  Me.MousePointer = vbDefault
  Exit Sub
End If

If Trim(txtNotas) = "" Or Len(txtNotas) < 10 Then
  MsgBox "Debe incluir uan nota valida...", vbInformation
  Me.MousePointer = vbDefault
  Exit Sub
Else

strSQL = "exec spCBR_CJ_CargaGastos " & txtTramite.Text & " ,'" & cboGasto.ItemData(cboGasto.ListIndex) & "'," & CCur(txtMonto) & "," _
                                & " '" & txtDocumento.Text & "','" & txtNotas & "','" & pBeneficiario & "'," _
                                & " '" & glogon.Usuario & "'," & chkDesembolso.Value & ", '" & txtDesembolsoId.Text & "'," & pBeneficiarioId _
                                & ", " & cboBanco.ItemData(cboBanco.ListIndex) & ", '" & pEmite & "', '" & cboCuenta.ItemData(cboCuenta.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

If rs!NumDoc <> "" Then
    Call sbImprimeRecibo(rs!NumDoc, rs!TipoDoc, False)
End If

rs.Close

MsgBox "Gasto registrado correctamente...", vbInformation


Me.MousePointer = vbDefault

UnLoad Me
    
End If
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

vPaso = True

txtTramite.Text = GLOBALES.gTag

txtMonto.Text = 0

cboEmite.Clear
cboEmite.AddItem "Transferencia"
cboEmite.AddItem "Cheque"
cboEmite.Text = "Transferencia"

strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

Call sbDatosIniciales

strSQL = "select TIPO_GASTO as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  CBR_CJ_TIPOS_GASTOS where activo = 1"
Call sbCbo_Llena_New(cboGasto, strSQL, False, True)


vPaso = False

Call cboGasto_Click
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
txtDesembolsoId.Text = vAbogadoCedula

End Sub


Private Sub optDesembolso_Click(Index As Integer)

If vPaso Then Exit Sub

If chkDesembolso.Value = vbChecked Then
    
    If Index = 0 Then
        txtAbogado.Text = vAbogado
        txtDesembolsoId.Text = vAbogadoCedula
        txtDesembolsoId.Locked = True
    Else
        txtDesembolsoId.Text = ""
        txtBeneficiario.Text = ""
        txtBeneficiario.Locked = False
        txtDesembolsoId.Locked = False
    End If

End If

Call cboBanco_Click

End Sub



Private Sub txtDesembolsoId_LostFocus()
    Call cboBanco_Click
End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNotas.SetFocus
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtDocumento.SetFocus

End Sub

Private Sub txtMonto_LostFocus()
txtMonto.Text = Format(txtMonto, "Standard")
End Sub


