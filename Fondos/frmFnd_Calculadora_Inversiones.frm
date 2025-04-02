VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmFnd_Calculadora_Inversiones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calculadora de Inversiones"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2535
      Left            =   120
      TabIndex        =   27
      Top             =   5760
      Width           =   10695
      _Version        =   1572864
      _ExtentX        =   18865
      _ExtentY        =   4471
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
      Appearance      =   21
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   2
      Left            =   10320
      TabIndex        =   34
      Top             =   5400
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmFnd_Calculadora_Inversiones.frx":0000
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   240
      Top             =   1440
   End
   Begin XtremeSuiteControls.GroupBox gbCalcular 
      Height          =   975
      Left            =   240
      TabIndex        =   28
      Top             =   4320
      Width           =   10455
      _Version        =   1572864
      _ExtentX        =   18441
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkCapitaliza 
         Height          =   375
         Left            =   4320
         TabIndex        =   37
         Top             =   360
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Capitaliza Rendimientos?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnCalcular 
         Height          =   615
         Left            =   6720
         TabIndex        =   29
         Top             =   360
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Calcular"
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
         Appearance      =   21
         Picture         =   "frmFnd_Calculadora_Inversiones.frx":016A
      End
      Begin XtremeSuiteControls.FlatEdit txtT_Intereses 
         Height          =   315
         Left            =   2160
         TabIndex        =   32
         Top             =   240
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtT_Inversion 
         Height          =   315
         Left            =   2160
         TabIndex        =   33
         Top             =   600
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Aportes o Inversión:"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Intereses Estimados:"
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
      End
   End
   Begin XtremeSuiteControls.GroupBox gbAhorro 
      Height          =   1095
      Left            =   2400
      TabIndex        =   8
      Top             =   2640
      Width           =   6735
      _Version        =   1572864
      _ExtentX        =   11880
      _ExtentY        =   1931
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2984
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboPlazo 
         Height          =   315
         Left            =   2280
         TabIndex        =   18
         Top             =   600
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   330
         Left            =   1560
         TabIndex        =   19
         Top             =   600
         Width           =   735
         _Version        =   1572864
         _ExtentX        =   1291
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo"
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
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblMensualidad 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensualidad"
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
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin XtremeSuiteControls.RadioButton rbTipo 
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
      _Version        =   1572864
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Plan de Ahorro"
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
      Appearance      =   21
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   1440
      Width           =   5055
      _Version        =   1572864
      _ExtentX        =   8911
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
   Begin XtremeSuiteControls.RadioButton rbTipo 
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
      _Version        =   1572864
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Certificados a Plazo"
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
      Appearance      =   21
   End
   Begin XtremeSuiteControls.ComboBox cboPlan 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   2280
      Width           =   6735
      _Version        =   1572864
      _ExtentX        =   11880
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
   Begin XtremeSuiteControls.GroupBox gbCDP 
      Height          =   1215
      Left            =   2400
      TabIndex        =   9
      Top             =   2640
      Width           =   6735
      _Version        =   1572864
      _ExtentX        =   11880
      _ExtentY        =   2143
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtInversion 
         Height          =   315
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2984
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboPlazoInversion 
         Height          =   330
         Left            =   1560
         TabIndex        =   21
         Top             =   600
         Width           =   1695
         _Version        =   1572864
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
      Begin XtremeSuiteControls.CheckBox chkCuponPaga 
         Height          =   255
         Left            =   5160
         TabIndex        =   23
         Top             =   240
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Paga Cupón?"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboCuponFrecuencia 
         Height          =   330
         Left            =   4560
         TabIndex        =   24
         Top             =   600
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
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
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cupón:"
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
         Left            =   3360
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo"
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
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblInversion 
         BackStyle       =   0  'Transparent
         Caption         =   "Inversión"
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
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   3960
      TabIndex        =   12
      Top             =   3960
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtTasa 
      Height          =   315
      Left            =   6960
      TabIndex        =   13
      Top             =   3960
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   550
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
      BackColor       =   16777215
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   9360
      TabIndex        =   35
      Top             =   5400
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmFnd_Calculadora_Inversiones.frx":04B6
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   36
      Top             =   5400
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmFnd_Calculadora_Inversiones.frx":0BBD
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   5400
      Width           =   10695
      _Version        =   1572864
      _ExtentX        =   18865
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Flujo Esperado"
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tasa Ref."
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
      Index           =   9
      Left            =   5880
      TabIndex        =   15
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
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
      Index           =   7
      Left            =   2640
      TabIndex        =   14
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Inversión"
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
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculadora de Inversiones"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   3
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   6372
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmFnd_Calculadora_Inversiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean, vFecha As Date
Dim pCalculoId As Long

Private Sub sbBoleta()

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = False
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Módulo de Planes de Ahorros"

   .Connect = glogon.ConectRPT
   
   strSQL = ""

            
   .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Calculo_Inversion.rpt")

    
    .Formulas(0) = "Fecha='Fecha: " & Format(fxFechaServidor, "yyyy-mm-dd") & "'"
    .Formulas(1) = "Usuario='Usuario: " & Trim(glogon.Usuario) & "'"
    .Formulas(2) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
    .Formulas(3) = "SubTitulo='Cálculo para Inversiones'"
    
    
    strSQL = "{vFnd_Calculadora_Inversiones.IdCalculo} = " & pCalculoId
    .SelectionFormula = strSQL
          
   .SubreportToChange = "sbFlujo"
   .StoredProcParam(0) = pCalculoId
   
   .Action = 1
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbEmail()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spFnd_Calculadora_Inversiones_Email " & pCalculoId & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Correo enviado a la persona!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnAccion_Click(Index As Integer)

If pCalculoId = 0 Then Exit Sub

Select Case Index
    Case 0 'Boleta
        Call sbBoleta
    Case 1 'Envia Correo
    
        Call sbEmail
        
    Case 2 'Exporta
        Call Excel_Exportar_Lsw(lsw)
End Select

End Sub

Private Sub btnCalcular_Click()
Dim pFrecuenciaPago As String

On Error GoTo vError

If txtNombre.Text = "" Then
    MsgBox "Indique a una persona para registrar el cálculo!", vbExclamation
    Exit Sub
End If

Select Case True
    Case rbTipo(0).Value 'Ahorros
        pFrecuenciaPago = "Mensual"
    Case rbTipo(1).Value 'CDPs
       If cboPlazoInversion.ListCount = 0 Then
            pFrecuenciaPago = "Mensual"
       
       Else
            pFrecuenciaPago = cboCuponFrecuencia.Text
       End If
       
End Select

strSQL = "exec spFnd_Calculadora_Inversiones_Registro " & IIf((pCalculoId = 0), "Null", pCalculoId) _
       & ", " & CCur(txtInversion.Text) & ",'" & Format(vFecha, "yyyy-mm-dd") & "', " & txtPlazo.Text _
       & ", " & txtTasa.Text & ", '" & pFrecuenciaPago & "', " & CCur(txtMonto.Text) & ", 360, " & chkCapitaliza.Value _
       & ", '" & txtCedula.Text & "', '" & cboPlan.ItemData(cboPlan.ListIndex) & "', 'ProGrX', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
    pCalculoId = rs!IdCalculo
rs.Close

Dim curIntereses As Currency, curInversion As Currency

curIntereses = 0
curInversion = 0

With lsw.ListItems
    .Clear
 
    strSQL = "exec spFnd_Calculadora_Inversiones_Flujo " & pCalculoId
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!Secuencia)
          itmX.SubItems(1) = Format(rs!FechaVencimiento, "yyyy-MM-dd")
          itmX.SubItems(2) = rs!DiasReconocimiento
          itmX.SubItems(3) = Format(rs!Tasa * 100, "Standard")
          itmX.SubItems(4) = Format(rs!BaseCalculo, "Standard")
          itmX.SubItems(5) = Format(rs!InteresesGanados, "Standard")
          itmX.SubItems(6) = Format(rs!AportacionExtra, "Standard")
          itmX.SubItems(7) = Format(rs!InteresGanadoAcumulado, "Standard")
          
          curIntereses = rs!InteresGanadoAcumulado
          curInversion = rs!BaseCalculo + rs!AportacionExtra
      rs.MoveNext
    Loop
    rs.Close

End With


txtT_Intereses.Text = Format(curIntereses, "Standard")
txtT_Inversion.Text = Format(curInversion, "Standard")

Exit Sub

vError:
  

End Sub


Private Sub sbConsultaPlan()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select TIPO_DEDUC, PORC_DEDUC, TIPO_CDP, PAGO_CUPONES, CAPITALIZA_RENDIMIENTOS" _
       & " from fnd_Planes" _
       & " where cod_operadora = 1 " _
       & " and cod_plan='" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
    chkCapitaliza.Value = rs!CAPITALIZA_RENDIMIENTOS
    chkCuponPaga.Value = rs!PAGO_CUPONES
End If

If rbTipo(1).Value Then

 vPaso = True
    strSQL = "exec spFnd_Inversion_Plazos '" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
    Call sbCbo_Llena_New(cboPlazoInversion, strSQL, False, True)
 vPaso = False
    
    Call cboPlazoInversion_Click
End If


Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cboCuponFrecuencia_Click()
Dim vAddMonths As Integer

If vPaso Then Exit Sub
If cboCuponFrecuencia.ListCount = 0 Then Exit Sub


Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "exec spFnd_Cupon_Frecuencia_Meses " & cboCuponFrecuencia.ItemData(cboCuponFrecuencia.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.BOF Then
    vAddMonths = rs!Frecuencia_Meses
Else
    vAddMonths = 0
End If
rs.Close

If chkCuponPaga.Value = xtpUnchecked Then
 'Al Vencimiento
 vAddMonths = 1000
End If

Call txtPlazo_KeyUp(vbKeyF4, 0)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboPlan_Click()
If vPaso Then Exit Sub

Call sbConsultaPlan

End Sub

Private Sub txtInversion_GotFocus()
On Error GoTo vError
  txtInversion = CCur(txtInversion)
vError:
End Sub

Private Sub txtInversion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  cboPlazoInversion.SetFocus
End If
End Sub

Private Sub txtInversion_LostFocus()
On Error GoTo vError

If IsNumeric(txtInversion) Then
      txtInversion = Format(CCur(txtInversion), "Standard")
End If
vError:
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
  txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
       txtPlazo.SetFocus
End If
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

If IsNumeric(txtMonto) Then
      txtMonto = Format(CCur(txtMonto), "Standard")
End If

vError:

End Sub


Private Sub Form_Load()
Dim strSQL As String, i As Integer

pCalculoId = 0

vModulo = 18 'Fondo de Inversion
Me.imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

vFecha = fxFechaServidor

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1500
    .Add , , "Vencimiento", 1800, vbCenter
    .Add , , "Dias", 1000, vbCenter
    .Add , , "Tasa", 1000, vbRightJustify
    .Add , , "Inversión", 2200, vbRightJustify
    .Add , , "Intereses", 2100, vbRightJustify
    .Add , , "Aportación", 2100, vbRightJustify
    .Add , , "Int.Acumulados", 2100, vbRightJustify
End With

cboPlazo.Clear
cboPlazo.AddItem "Días"
cboPlazo.AddItem "Meses"
cboPlazo.Text = "Días"

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'Idx' from FND_Operadoras"
'Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

vPaso = True
    strSQL = "SELECT ID_FRECUENCIACUPON as 'Idx' , dbo.fxSys_Cadena_Capitaliza ( CUPON ) as 'ItmX'" _
           & " FROM FND_CDP_FRECUENCIACUPONES Where Estado = 1 Order by FRECUENCIA_DIAS asc"
    Call sbCbo_Llena_New(cboCuponFrecuencia, strSQL, False, True)
    
    strSQL = "select ID_PLAZO as 'IdX', dbo.fxSys_Cadena_Capitaliza ( PLAZO ) as 'ItmX'" _
           & " From FND_CDP_PLAZOS  Where Estado = 1  Order by PLAZO_DIAS  asc"
    Call sbCbo_Llena_New(cboPlazoInversion, strSQL, False, True)
vPaso = False
'
'With cboCuponFrecuencia
'  .Clear
'  .AddItem "No Aplica"
'  .AddItem "Mensuales"
'  .AddItem "Trimestrales"
'  .AddItem "Semestrales"
'  .AddItem "Anual"
'  .AddItem "Cuatrimestral"
'  .AddItem "Al Vencimiento"
'End With
'
vPaso = False

 
Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbLimpia()

lsw.ListItems.Clear
txtT_Intereses.Text = "0.00"
txtT_Inversion.Text = "0.00"
txtTasa.Text = "0.00"

txtPlazo.Text = "30"
txtMonto.Text = "0.00"
txtInversion.Text = "0.00"


End Sub

Private Function fxTasaRef(xPlazo As Long, xTipo As String, xPlan As String _
                         , xOperadora As Integer) As Currency
Dim strSQL As String, rs As New ADODB.Recordset
Dim xTasa As Currency

On Error GoTo vError

If chkCuponPaga.Value = xtpUnchecked Or Not rbTipo(1).Value Then
    strSQL = "select dbo.fxFNDCalcularTasaRefContrato(" & xOperadora & ", '" & xPlan & "', " & txtPlazo.Text & ", '" & xTipo & "', Null, Null, 0) as 'Tasa'"
Else
    If cboCuponFrecuencia.ListIndex = -1 Then
      If IsNumeric(txtTasa.Text) Then
          fxTasaRef = CCur(txtTasa.Text)
      Else
          fxTasaRef = 0
      End If
      Exit Function
    End If
    
    strSQL = "exec dbo.spFnd_Inversion_Tasas_Condiciones " & xOperadora & ", '" & xPlan & "', " & cboPlazoInversion.ItemData(cboPlazoInversion.ListIndex) & ", " & cboCuponFrecuencia.ItemData(cboCuponFrecuencia.ListIndex)
End If

Call OpenRecordSet(rs, strSQL)
    xTasa = rs!Tasa
rs.Close

'If vCarga Then
'    fxTasaRef = CCur(txtTasa.Text)
'Else
'    fxTasaRef = xTasa
'End If
fxTasaRef = xTasa

Exit Function

vError:
    fxTasaRef = CCur(txtTasa.Text)

End Function



Private Sub txtPlazo_GotFocus()
On Error GoTo vError

'If CCur(txtTasa) > 0 Then
'   txtIntereses = CCur(txtInversion) * IIf((Mid(cboPlazo.Text, 1, 1) = "D"), CLng(txtPlazo), CLng(txtPlazo) * 30) * CCur(txtTasa) / 36500
'   txtIntereses = Format(txtIntereses, "Standard")
'End If

vError:
End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub txtPlazo_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If Mid(cboPlazo.Text, 1, 1) = "D" Then
    dtpCorte.Value = DateAdd("d", CDbl(txtPlazo), CDate(vFecha))
Else
    dtpCorte.Value = DateAdd("m", CDbl(txtPlazo), CDate(vFecha))
End If

txtTasa.Text = Format(fxTasaRef(txtPlazo.Text, Mid(cboPlazo.Text, 1, 1), cboPlan.ItemData(cboPlan.ListIndex), 1), "##0.00")

'If CCur(txtTasa) > 0 Then
'   txtT_Intereses.Text = CCur(txtInversion.Text) * IIf((Mid(cboPlazo.Text, 1, 1) = "D"), CLng(txtPlazo), CLng(txtPlazo) * 30) * CCur(txtTasa.Text) / 36500
'   txtT_Intereses.Text = Format(txtIntereses.Text, "Standard")
'End If

vError:

End Sub



Private Sub cboPlazoInversion_Click()
If vPaso Then Exit Sub


If cboPlazoInversion.ListCount = 0 Then Exit Sub
On Error GoTo vError

Me.MousePointer = vbHourglass


 vPaso = True
    strSQL = "exec spFnd_Cupon_Frecuencia " & cboPlazoInversion.ItemData(cboPlazoInversion.ListIndex) _
           & ",  '" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
    Call sbCbo_Llena_New(cboCuponFrecuencia, strSQL, False, True)
 vPaso = False
 
 
  strSQL = "exec spFnd_Inversion_Plazos_Dias " & cboPlazoInversion.ItemData(cboPlazoInversion.ListIndex)
  Call OpenRecordSet(rs, strSQL)
    If Mid(cboPlazo.Text, 1, 1) = "D" Then
      txtPlazo.Text = rs!PLAZO_DIAS
    Else
      txtPlazo.Text = rs!Plazo_Meses
    End If
  rs.Close
      
 Call txtPlazo_KeyUp(vbKeyF4, 0)

Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub rbTipo_Click(Index As Integer)
 
On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpia

Select Case Index
    
    Case 0
        strSQL = "exec spFnd_Calculadora_Planes 'APL'"
        gbAhorro.Visible = True
        gbCDP.Visible = False
    
    Case 1
        strSQL = "exec spFnd_Calculadora_Planes 'CDP'"
        gbAhorro.Visible = False
        gbCDP.Visible = True
End Select

vPaso = True
    Call sbCbo_Llena_New(cboPlan, strSQL, False, True)
vPaso = False

If Index = 1 Then 'CDP
    vPaso = True
        
    strSQL = "exec spFnd_Inversion_Plazos '" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
    Call sbCbo_Llena_New(cboPlazoInversion, strSQL, False, True)
    
    vPaso = False
    
    Call cboPlazoInversion_Click
End If

Me.MousePointer = vbDefault

 Call txtPlazo_KeyUp(vbKeyF4, 0)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call rbTipo_Click(0)

If GLOBALES.gCedulaActual <> "" Then
    txtCedula.Text = GLOBALES.gCedulaActual
    txtCedula_LostFocus
End If

End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cedula"
   gBusquedas.Orden = "cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "select cedula,nombre from socios"
   frmBusquedas.Show vbModal
   txtNombre.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      pCalculoId = 0
      txtCedula = Trim(gBusquedas.Resultado)
      txtNombre = Trim(gBusquedas.Resultado2)
      Call sbLimpia
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

End Sub

Private Sub txtCedula_LostFocus()
 pCalculoId = 0
 txtNombre = fxNombre(txtCedula)
End Sub

