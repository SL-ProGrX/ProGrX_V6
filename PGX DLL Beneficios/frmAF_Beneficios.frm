VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_Beneficios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Beneficios"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   9015
      _Version        =   1441793
      _ExtentX        =   15901
      _ExtentY        =   11245
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
      ItemCount       =   3
      Item(0).Caption =   "General"
      Item(0).ControlCount=   23
      Item(0).Control(0)=   "Label2(0)"
      Item(0).Control(1)=   "Label2(1)"
      Item(0).Control(2)=   "Label2(2)"
      Item(0).Control(3)=   "Label2(3)"
      Item(0).Control(4)=   "Label2(4)"
      Item(0).Control(5)=   "chkMonetario"
      Item(0).Control(6)=   "chkProducto"
      Item(0).Control(7)=   "chkAplBeneficiarios"
      Item(0).Control(8)=   "chkAplCambioMonto"
      Item(0).Control(9)=   "chkParcial"
      Item(0).Control(10)=   "cboEstado"
      Item(0).Control(11)=   "cboTipo"
      Item(0).Control(12)=   "txtNotas"
      Item(0).Control(13)=   "Label2(5)"
      Item(0).Control(14)=   "txtOtorgamiento"
      Item(0).Control(15)=   "txtDiferencia"
      Item(0).Control(16)=   "GroupBox1"
      Item(0).Control(17)=   "cboTipoBeneficio"
      Item(0).Control(18)=   "txtVigenciaMeses"
      Item(0).Control(19)=   "Label2(7)"
      Item(0).Control(20)=   "Label2(8)"
      Item(0).Control(21)=   "Label2(9)"
      Item(0).Control(22)=   "cboGrupo"
      Item(1).Caption =   "Membresías"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Item(2).Caption =   "Asignación de Roles"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5775
         Left            =   -69880
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   8775
         _Version        =   1441793
         _ExtentX        =   15478
         _ExtentY        =   10186
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   5280
         Width           =   8655
         _Version        =   1441793
         _ExtentX        =   15266
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Cuenta Contable:"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCtaCod 
            Height          =   312
            Left            =   720
            TabIndex        =   23
            Top             =   480
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3619
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
         Begin XtremeSuiteControls.FlatEdit txtCtaDes 
            Height          =   315
            Left            =   2760
            TabIndex        =   24
            Top             =   480
            Width           =   5775
            _Version        =   1441793
            _ExtentX        =   10186
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
      End
      Begin XtremeSuiteControls.CheckBox chkMonetario 
         Height          =   255
         Left            =   6120
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monetario"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkProducto 
         Height          =   255
         Left            =   7560
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Productos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkAplBeneficiarios 
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   3720
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10604
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica también a Beneficiarios de la Persona"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkAplCambioMonto 
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   4080
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10604
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica Cambio de Monto del Beneficio x Caso"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkParcial 
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   4440
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10604
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica Monto parcial"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5655
         Left            =   -69280
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   7335
         _Version        =   524288
         _ExtentX        =   12938
         _ExtentY        =   9975
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
         MaxCols         =   498
         ScrollBars      =   2
         SpreadDesigner  =   "frmAF_Beneficios.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   1800
         TabIndex        =   14
         Top             =   1320
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
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
         Height          =   330
         Left            =   4320
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   915
         Left            =   1800
         TabIndex        =   18
         Top             =   1800
         Width           =   6975
         _Version        =   1441793
         _ExtentX        =   12303
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
      Begin XtremeSuiteControls.FlatEdit txtOtorgamiento 
         Height          =   315
         Left            =   1800
         TabIndex        =   20
         Top             =   2880
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.FlatEdit txtDiferencia 
         Height          =   315
         Left            =   1800
         TabIndex        =   21
         Top             =   4800
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.ComboBox cboTipoBeneficio 
         Height          =   345
         Left            =   1800
         TabIndex        =   27
         Top             =   480
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin XtremeSuiteControls.FlatEdit txtVigenciaMeses 
         Height          =   315
         Left            =   1800
         TabIndex        =   28
         Top             =   3240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.ComboBox cboGrupo 
         Height          =   345
         Left            =   1800
         TabIndex        =   31
         Top             =   840
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   32
         Top             =   840
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Grupo"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   30
         Top             =   480
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Categoría"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   7
         Left            =   3360
         TabIndex        =   29
         Top             =   3240
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "(Vigencia del Beneficio en meses )"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   19
         Top             =   4800
         Width           =   3255
         _Version        =   1441793
         _ExtentX        =   5736
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Diferencia de Monto x Cambio"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   9
         Top             =   2880
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "(Máximo de veces)"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   8
         Top             =   2880
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Otorgamiento"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Estado"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   4
         Top             =   1320
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Entrega"
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
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reportes"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1560
      TabIndex        =   16
      Top             =   600
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3000
      TabIndex        =   17
      Top             =   600
      Width           =   5172
      _Version        =   1441793
      _ExtentX        =   9123
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8280
      TabIndex        =   26
      Top             =   600
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   6
      Left            =   0
      TabIndex        =   25
      Top             =   600
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Beneficio:"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAF_Beneficios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, strObjeto As String
Dim vScroll As Boolean, vPaso As Boolean


Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOtorgamiento.SetFocus
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub cboTipoBeneficio_Click()
If vPaso Then Exit Sub

Call sbGrupos_Load

End Sub

Private Sub chkAplCambioMonto_Click()
If chkAplCambioMonto.Value = xtpChecked Then
    txtDiferencia.Locked = False
Else
    txtDiferencia.Locked = True
End If
End Sub


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 Cod_Beneficio from afi_beneficios"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where Cod_Beneficio > '" & txtCodigo.Text & "' order by Cod_Beneficio asc"
    Else
       strSQL = strSQL & " where Cod_Beneficio < '" & txtCodigo.Text & "' order by Cod_Beneficio desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Cod_Beneficio
      Call sbConsulta(rs!Cod_Beneficio)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 7

End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

 vModulo = 7
 
 tcMain.Item(0).Selected = True
 
 vGrid.AppearanceStyle = fxGridStyle
     
 With lsw.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2100
    .Add , , "Descripción", 3800
 End With
    
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
 vEdita = True
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 
 vPaso = True
    strSQL = "select cod_categoria as 'IdX', descripcion as 'ItmX'" _
           & " From afi_bene_categorias  where Activo = 1 order by descripcion"
    Call sbCbo_Llena_New(cboTipoBeneficio, strSQL, False, True)
 vPaso = False
 
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbGrupos_Load()
Dim strSQL As String

On Error GoTo vError

vPaso = True

strSQL = " select COD_GRUPO as 'IdX', DESCRIPCION as 'ItmX'" _
       & "   From AFI_BENE_GRUPOS" _
       & "  Where Cod_Categoria = '" & cboTipoBeneficio.ItemData(cboTipoBeneficio.ListIndex) & "'" _
       & "  and Estado = 1 order by DESCRIPCION"
Call sbCbo_Llena_New(cboGrupo, strSQL, False, True)

vPaso = False



Exit Sub

vError:
End Sub

Private Sub sbLimpiaPantalla()
vCodigo = ""
txtCodigo.Text = ""

tcMain.Item(0).Selected = True

cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "Inactivo"
cboEstado.Text = "Activo"

cboTipo.Clear
cboTipo.AddItem "Monetario"
cboTipo.AddItem "Producto"
cboTipo.AddItem "Ambos"
cboTipo = "Monetario"

txtDescripcion.Text = ""
txtNotas.Text = ""
txtCtaCod.Text = ""
txtCtaDes.Text = ""


txtOtorgamiento.Text = 0
txtDiferencia.Text = 0

chkMonetario.Value = xtpUnchecked
chkProducto.Value = xtpUnchecked

txtVigenciaMeses.Text = "12"

chkAplBeneficiarios.Value = xtpUnchecked
chkAplCambioMonto.Value = xtpUnchecked
chkParcial.Value = xtpUnchecked

Call sbGrupos_Load

End Sub




Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert AFI_BENE_GRUPOSB(cod_grupo,cod_beneficio) values('" & Item.Text _
            & "','" & vCodigo & "')"
Else
   strSQL = "Delete AFI_BENE_GRUPOSB where cod_grupo = '" & Item.Text _
          & "' and cod_beneficio = '" & vCodigo & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset

If vCodigo = "" And Item.Index <> 0 Then
    tcMain.Item(0).Selected = True
    Exit Sub
End If

Select Case Item.Index
  Case 1 'Membresias
    strSQL = "select id_bene,inicio,corte,monto from afi_beneficio_montos" _
           & " where cod_beneficio = '" & vCodigo & "'"
    Call sbCargaGrid(vGrid, 4, strSQL)
   Case 2 'Roles
     Call sbCarga_Roles
End Select
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
      
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
      
    Case "BORRAR"
      Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
     
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
         Select Case strObjeto
            Case "txtCodigo", ""
                txtCodigo.SetFocus
                Call txtCodigo_KeyDown(vbKeyF4, 0)
            Case Else
                txtDescripcion.SetFocus
                Call txtDescripcion_KeyDown(vbKeyF4, 0)
           
         End Select
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vAFI_Beneficios_Catalogo where Cod_Beneficio = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!Cod_Beneficio
  txtCodigo.Text = rs!Cod_Beneficio
  
  txtDescripcion.Text = rs!Descripcion & ""
  txtNotas.Text = rs!notas & ""
    
  If rs!Estado = "A" Then
    cboEstado.Text = "Activo"
  Else
    cboEstado.Text = "Inactivo"
  End If
  
  Select Case rs!Tipo
    Case "M"
        cboTipo.Text = "Monetario"
    Case "P"
        cboTipo.Text = "Producto"
    Case "A"
        cboTipo.Text = "Ambos"
  End Select
  chkMonetario.Value = IIf(IsNull(rs!tipo_monetario), 0, rs!tipo_monetario)
  chkProducto.Value = IIf(IsNull(rs!tipo_producto), 0, rs!tipo_producto)
  
  
  chkAplBeneficiarios.Value = rs!aplica_beneficiarios
  chkAplCambioMonto.Value = rs!modifica_monto
  chkParcial.Value = IIf(IsNull(rs!aplica_parcial), 0, rs!aplica_parcial)
  
  txtOtorgamiento.Text = CStr(rs!maximo_otorga)
  txtDiferencia.Text = Format(rs!modifica_diferencia, "Standard")

  txtCtaCod.Text = fxgCntCuentaFormato(True, rs!cod_cuenta)
  txtCtaDes.Text = fxgCntCuentaDesc(rs!cod_cuenta)

  txtVigenciaMeses.Text = rs!VIGENCIA_MESES
  
  
  Call sbCboAsignaDato(cboTipoBeneficio, rs!Categoria_Desc, True, rs!Cod_Categoria)
  Call sbCboAsignaDato(cboGrupo, rs!Grupo_Desc, True, rs!Cod_Grupo)

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Me.MousePointer = vbDefault

Call RefrescaTags(Me)
tcMain.Item(0).Selected = True

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Proveedor no es válido ..."
If chkMonetario.Value = 0 And chkProducto.Value = 0 Then vMensaje = vMensaje & vbCrLf & " - Defina Monetario o Producto ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update Afi_beneficios set descripcion = '" & Trim(txtDescripcion.Text) & "'" _
         & ",  notas = '" & txtNotas & "',estado = '" & Mid(cboEstado.Text, 1, 1) _
         & "', aplica_beneficiarios = " & chkAplBeneficiarios.Value & ",modifica_monto = " & chkAplCambioMonto.Value _
         & ",  cod_cuenta = '" & fxgCntCuentaFormato(False, txtCtaCod) & "',tipo = '" & Mid(cboTipo, 1, 1) & "'" _
         & ",  modifica_diferencia = " & CCur(txtDiferencia) & ",maximo_otorga = " & Val(txtOtorgamiento.Text) _
         & ",  aplica_parcial = " & chkParcial.Value _
         & ",  tipo_monetario = " & chkMonetario.Value & ",tipo_producto = " & chkProducto.Value _
         & ",  Cod_Categoria = '" & cboTipoBeneficio.ItemData(cboTipoBeneficio.ListIndex) _
         & "', Cod_Grupo = '" & cboGrupo.ItemData(cboGrupo.ListIndex) & "', VIGENCIA_MESES = " & txtVigenciaMeses.Text _
         & ", Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" & glogon.Usuario & "'" _
         & " where cod_beneficio = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Beneficio : " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert into afi_beneficios (cod_beneficio,descripcion,notas,estado,registra_fecha" _
          & ",registra_user, maximo_otorga, modifica_monto" _
          & ",modifica_diferencia, cod_cuenta, aplica_beneficiarios, aplica_parcial, tipo_monetario, tipo_producto, tipo" _
          & ", Cod_Categoria, Cod_Grupo, VIGENCIA_MESES)" _
          & " values('" & vCodigo & "','" & Trim(txtDescripcion) & "','" & txtNotas & "','" _
          & Mid(cboEstado.Text, 1, 1) & "', dbo.MyGetdate() ,'" & glogon.Usuario & "'," & Val(txtOtorgamiento) & "," & "" _
          & chkAplCambioMonto.Value & " , " & CCur(txtDiferencia) & ", '" & fxgCntCuentaFormato(False, txtCtaCod) & "', " & "" _
          & chkAplBeneficiarios.Value & "," & chkParcial.Value & "," & chkMonetario.Value & "," & chkProducto.Value _
          & ", '" & Mid(cboTipo.Text, 1, 1) & "', '" & cboTipoBeneficio.ItemData(cboTipoBeneficio.ListIndex) _
          & "', '" & cboGrupo.ItemData(cboGrupo.ListIndex) & "', " & txtVigenciaMeses.Text & ")"
   
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Beneficio: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbToolBar(tlb, "activo")

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
  strSQL = "delete afi_beneficios where cod_beneficio = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Bodega : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_Click()
strObjeto = txtCodigo.Name
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  
  If txtCodigo <> "" And vEdita Then
    Call sbConsulta(txtCodigo)
  End If
  
  tcMain.Item(0).Selected = True
  txtDescripcion.SetFocus
  
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Cod_Beneficio"
  gBusquedas.Orden = "Cod_Beneficio"
  gBusquedas.Consulta = "select Cod_Beneficio,descripcion from afi_beneficios"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtCtaCod_GotFocus()
txtCtaCod.Text = fxgCntCuentaFormato(False, txtCtaCod.Text)
End Sub

Private Sub txtCtaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaDes.SetFocus

If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta
  
  txtCtaCod.Text = fxgCntCuentaFormato(True, gBusquedas.Resultado)
  txtCtaDes.Text = fxgCntCuentaDesc(gBusquedas.Resultado)
End If

End Sub

Private Sub txtCtaCod_LostFocus()
txtCtaCod.Text = fxgCntCuentaFormato(True, txtCtaCod.Text)
txtCtaDes.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaCod.Text))
End Sub


Private Sub txtDescripcion_Click()
strObjeto = txtDescripcion.Name
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_beneficio,descripcion from afi_beneficios"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtDiferencia_GotFocus()
On Error GoTo vError
  txtDiferencia = CCur(txtDiferencia)
vError:
End Sub

Private Sub txtDiferencia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaCod.SetFocus

End Sub


Private Sub txtDiferencia_LostFocus()
On Error GoTo vError
  txtDiferencia = Format(CCur(txtDiferencia), "Standard")
vError:
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub


Private Sub txtOtorgamiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDiferencia.SetFocus

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

If KeyCode = vbKeyReturn And vGrid.ActiveCol = 4 Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   If vGrid.Text = "" Then
      strSQL = "select isnull(max(id_bene),0) + 1 as Secuencia" _
             & " from afi_beneficio_montos where cod_beneficio = '" _
             & vCodigo & "'"
      Call OpenRecordSet(rs, strSQL)
        i = rs!secuencia
      rs.Close
      
      vGrid.Col = 2
      strSQL = "insert afi_beneficio_montos(id_bene,cod_beneficio,inicio,corte,monto) values(" _
             & i & ",'" & vCodigo & "'," & CCur(vGrid.Text) & ","
      vGrid.Col = 3
      strSQL = strSQL & CCur(vGrid.Text) & ","
      vGrid.Col = 4
      strSQL = strSQL & CCur(vGrid.Text) & ")"
      Call ConectionExecute(strSQL)
      
      vGrid.Col = 1
      vGrid.Text = CStr(i)
      
      vGrid.MaxRows = vGrid.MaxRows + 1
      
   Else
      vGrid.Col = 2
      strSQL = "update afi_beneficio_montos set inicio = " & CCur(vGrid.Text) & ",corte = "
      vGrid.Col = 3
      strSQL = strSQL & CCur(vGrid.Text) & ",monto = "
      vGrid.Col = 4
      strSQL = strSQL & CCur(vGrid.Text) & " where cod_beneficio = '" & vCodigo _
             & "' and id_bene = "
      vGrid.Col = 1
      strSQL = strSQL & vGrid.Text
      Call ConectionExecute(strSQL)
   End If
End If

If KeyCode = vbKeyDelete Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   If vGrid.Text <> "" Then
        strSQL = "delete afi_beneficio_montos where cod_beneficio = '" & vCodigo _
               & "' and id_bene = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
   End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation


End Sub

Private Sub sbCarga_Roles()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, i As Integer

lsw.ListItems.Clear
vPaso = True

strSQL = "select B.cod_grupo as 'Grupo',B.descripcion, A.cod_grupo" _
        & " from AFI_BENEFICIO_GRUPOS  B left join AFI_BENE_GRUPOSB A on B.cod_grupo = A.cod_grupo" _
        & " and  A.cod_beneficio = '" & vCodigo & "' " _
        & " order by A.cod_grupo desc,B.descripcion asc"
          
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Grupo)
     itmX.SubItems(1) = rs!Descripcion
 If Not IsNull(rs!Cod_Grupo) Then
     itmX.Checked = True
     itmX.ForeColor = vbBlue
 End If
 
 rs.MoveNext
Loop
rs.Close

vPaso = False

End Sub
