VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmPreaEstudio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estudio de Crédito"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   Icon            =   "frmPreaEstudio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6015
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   10455
      _Version        =   1572864
      _ExtentX        =   18441
      _ExtentY        =   10610
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
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   6
      Item(0).Caption =   "Datos"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "GroupBox1(0)"
      Item(0).Control(1)=   "GroupBox1(1)"
      Item(1).Caption =   "Cálculo"
      Item(1).ControlCount=   69
      Item(1).Control(0)=   "txtSalarioReal"
      Item(1).Control(1)=   "txtExtrasFijas"
      Item(1).Control(2)=   "txtDevengadoMes"
      Item(1).Control(3)=   "txtTotal_Cargas_CCSS"
      Item(1).Control(4)=   "txtPorcSobreSalario"
      Item(1).Control(5)=   "txtDeducciones"
      Item(1).Control(6)=   "txtCrdTransitoCancelados"
      Item(1).Control(7)=   "txtCrdTransitoXCobrar"
      Item(1).Control(8)=   "txtFianzas"
      Item(1).Control(9)=   "txtLiquidezSinFianza"
      Item(1).Control(10)=   "txtLiquidezPorcSinFianza"
      Item(1).Control(11)=   "txtLiquidezConFianza"
      Item(1).Control(12)=   "txtLiquidezPorcConFianza"
      Item(1).Control(13)=   "txtPSD"
      Item(1).Control(14)=   "txtMontoGirar"
      Item(1).Control(15)=   "txtCuotaDiferencia"
      Item(1).Control(16)=   "lblMontoGirar(34)"
      Item(1).Control(17)=   "Label1(32)"
      Item(1).Control(18)=   "Label1(33)"
      Item(1).Control(19)=   "Label1(31)"
      Item(1).Control(20)=   "Label1(30)"
      Item(1).Control(21)=   "Label1(28)"
      Item(1).Control(22)=   "Label1(27)"
      Item(1).Control(23)=   "Label1(42)"
      Item(1).Control(24)=   "lblMontoGirar(0)"
      Item(1).Control(25)=   "imgCuotaDif"
      Item(1).Control(26)=   "Label1(22)"
      Item(1).Control(27)=   "Label1(21)"
      Item(1).Control(28)=   "Label1(20)"
      Item(1).Control(29)=   "Label1(19)"
      Item(1).Control(30)=   "lblPorcentajeSalario"
      Item(1).Control(31)=   "Label1(13)"
      Item(1).Control(32)=   "Label1(12)"
      Item(1).Control(33)=   "Label1(11)"
      Item(1).Control(34)=   "txtRebajoExtras"
      Item(1).Control(35)=   "txtTotalLiquido"
      Item(1).Control(36)=   "Label1(26)"
      Item(1).Control(37)=   "Label1(10)"
      Item(1).Control(38)=   "txtSalarioDevengado"
      Item(1).Control(39)=   "lblSalarioDevengado(9)"
      Item(1).Control(40)=   "dtpCorte"
      Item(1).Control(41)=   "Label1(18)"
      Item(1).Control(42)=   "txtRefundiciones"
      Item(1).Control(43)=   "Label1(25)"
      Item(1).Control(44)=   "txtDesembolsos"
      Item(1).Control(45)=   "Label1(29)"
      Item(1).Control(46)=   "cboSalario"
      Item(1).Control(47)=   "txtSalarioLiquido"
      Item(1).Control(48)=   "Label1(5)"
      Item(1).Control(49)=   "Label1(24)"
      Item(1).Control(50)=   "Line3(2)"
      Item(1).Control(51)=   "Line3(0)"
      Item(1).Control(52)=   "Line3(3)"
      Item(1).Control(53)=   "Line2"
      Item(1).Control(54)=   "Line3(4)"
      Item(1).Control(55)=   "Line4"
      Item(1).Control(56)=   "Line3(1)"
      Item(1).Control(57)=   "fraDCargas"
      Item(1).Control(58)=   "FrameSalarios"
      Item(1).Control(59)=   "btnDetalle(0)"
      Item(1).Control(60)=   "btnDetalle(1)"
      Item(1).Control(61)=   "btnDetalle(2)"
      Item(1).Control(62)=   "btnDetalle(3)"
      Item(1).Control(63)=   "btnDetalle(4)"
      Item(1).Control(64)=   "btnDetalle(5)"
      Item(1).Control(65)=   "btnDetalle(6)"
      Item(1).Control(66)=   "btnDetalle(7)"
      Item(1).Control(67)=   "txtTotalLiquidoGrupo"
      Item(1).Control(68)=   "Label1(0)"
      Item(2).Caption =   "Observaciones"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "tcAux"
      Item(3).Caption =   "Calificación"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "vGrid"
      Item(3).Control(1)=   "txtCumplimientoNotas"
      Item(4).Caption =   "Etiquetas"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "vGridTags"
      Item(5).Caption =   "Causas"
      Item(5).ControlCount=   3
      Item(5).Control(0)=   "optCausas(0)"
      Item(5).Control(1)=   "optCausas(1)"
      Item(5).Control(2)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5055
         Left            =   -70000
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1572864
         _ExtentX        =   18230
         _ExtentY        =   8916
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Frame fraDCargas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cargas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2055
         Left            =   -64000
         TabIndex        =   118
         Top             =   3360
         Visible         =   0   'False
         Width           =   4815
         Begin XtremeSuiteControls.CheckBox chkCargaAsociacion 
            Height          =   252
            Left            =   3360
            TabIndex        =   127
            Top             =   360
            Width           =   1212
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkCargaFrap 
            Height          =   252
            Left            =   3360
            TabIndex        =   128
            Top             =   720
            Width           =   1212
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin VB.Label lblCargaImpSalario 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
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
            Height          =   315
            Left            =   1800
            TabIndex        =   126
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label lblCargaCCSS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
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
            Height          =   315
            Left            =   1800
            TabIndex        =   125
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblCargaFrap 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
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
            Height          =   315
            Left            =   1800
            TabIndex        =   124
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblCargaAsociacion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
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
            Height          =   315
            Left            =   1800
            TabIndex        =   123
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(-) C.C.S.S."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   14
            Left            =   360
            TabIndex        =   122
            Top             =   1200
            Width           =   1212
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(-) Asociación"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   15
            Left            =   360
            TabIndex        =   121
            Top             =   360
            Width           =   1332
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(-) FAP/FRAP"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   17
            Left            =   360
            TabIndex        =   120
            Top             =   720
            Width           =   1332
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(-) Imp.Salario"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   23
            Left            =   360
            TabIndex        =   119
            Top             =   1560
            Width           =   1332
         End
         Begin VB.Image imgFraCerrar 
            Height          =   240
            Left            =   4530
            Picture         =   "frmPreaEstudio.frx":000C
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.Frame FrameSalarios 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3015
         Left            =   -64000
         TabIndex        =   129
         Top             =   600
         Visible         =   0   'False
         Width           =   4815
         Begin XtremeSuiteControls.ListView ltvSalarios 
            Height          =   2535
            Left            =   5160
            TabIndex        =   130
            Top             =   360
            Width           =   4575
            _Version        =   1572864
            _ExtentX        =   8064
            _ExtentY        =   4466
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
            ShowBorder      =   0   'False
         End
         Begin VB.Image CerrarFrameSalarios 
            Height          =   240
            Left            =   4560
            Picture         =   "frmPreaEstudio.frx":0712
            Top             =   0
            Width           =   240
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   3972
         Index           =   1
         Left            =   120
         TabIndex        =   72
         Top             =   1920
         Width           =   9972
         _Version        =   1572864
         _ExtentX        =   17590
         _ExtentY        =   7006
         _StockProps     =   79
         Caption         =   "Datos para el Crédito"
         ForeColor       =   8421504
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
         BorderStyle     =   1
         Begin VB.ComboBox cboComite 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPreaEstudio.frx":0E18
            Left            =   1440
            List            =   "frmPreaEstudio.frx":0E1A
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   1608
            Width           =   2652
         End
         Begin VB.ComboBox cboFondo 
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
            Height          =   330
            ItemData        =   "frmPreaEstudio.frx":0E1C
            Left            =   5280
            List            =   "frmPreaEstudio.frx":0E1E
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   1248
            Width           =   4572
         End
         Begin VB.ComboBox cboGarantia 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPreaEstudio.frx":0E20
            Left            =   1440
            List            =   "frmPreaEstudio.frx":0E22
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   1248
            Width           =   2652
         End
         Begin VB.ComboBox cboDestino 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPreaEstudio.frx":0E24
            Left            =   1440
            List            =   "frmPreaEstudio.frx":0E26
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   888
            Width           =   8412
         End
         Begin XtremeSuiteControls.CheckBox chkPolizaVida 
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   2040
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Póliza de Vida"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkPolizaIncendio 
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   2400
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Póliza de Incendio"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkPrimerCuota 
            Height          =   612
            Left            =   4560
            TabIndex        =   90
            Top             =   3240
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Primera Cuota"
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
         Begin XtremeSuiteControls.FlatEdit txtLinea 
            Height          =   312
            Left            =   1440
            TabIndex        =   91
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   480
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDesLineaCredito 
            Height          =   312
            Left            =   3240
            TabIndex        =   92
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   480
            Width           =   6612
            _Version        =   1572864
            _ExtentX        =   11663
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnComiteCambio 
            Height          =   312
            Left            =   4200
            TabIndex        =   93
            ToolTipText     =   "Cambio de Evaluador"
            Top             =   1608
            Width           =   315
            _Version        =   1572864
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Appearance      =   16
            Picture         =   "frmPreaEstudio.frx":0E28
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizaVida 
            Height          =   312
            Left            =   2520
            TabIndex        =   94
            Top             =   2040
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
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
            BackColor       =   16777215
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizaIncendio 
            Height          =   312
            Left            =   2520
            TabIndex        =   95
            Top             =   2400
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAsignado 
            Height          =   312
            Left            =   2520
            TabIndex        =   96
            Top             =   3240
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtClasificacion 
            Height          =   312
            Left            =   2520
            TabIndex        =   97
            Top             =   3600
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuota 
            Height          =   312
            Left            =   7800
            TabIndex        =   98
            Top             =   3240
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
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
         Begin XtremeSuiteControls.FlatEdit txtCompromiso 
            Height          =   312
            Left            =   7800
            TabIndex        =   99
            Top             =   3600
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
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
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   312
            Left            =   7800
            TabIndex        =   100
            Top             =   2040
            Width           =   2052
            _Version        =   1572864
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTasa 
            Height          =   312
            Left            =   8760
            TabIndex        =   101
            Top             =   2760
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlazo 
            Height          =   312
            Left            =   8760
            TabIndex        =   102
            Top             =   2400
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlMax 
            Height          =   312
            Left            =   7800
            TabIndex        =   103
            Top             =   2400
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
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
         Begin XtremeSuiteControls.CheckBox chkPolizaDesempleo 
            Height          =   255
            Left            =   120
            TabIndex        =   150
            Top             =   2760
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Póliza de Desempleo"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizaDesempleo 
            Height          =   312
            Left            =   2520
            TabIndex        =   151
            Top             =   2760
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   324
            Index           =   1
            Left            =   120
            TabIndex        =   115
            Top             =   480
            Width           =   1212
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Respaldo"
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
            Height          =   312
            Index           =   34
            Left            =   4176
            TabIndex        =   114
            Top             =   1248
            Width           =   1092
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Evaluado"
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
            Height          =   312
            Index           =   7
            Left            =   120
            TabIndex        =   113
            Top             =   1560
            Width           =   1212
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Destino"
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
            Height          =   312
            Index           =   6
            Left            =   120
            TabIndex        =   112
            Top             =   840
            Width           =   1212
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Garantía"
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
            Height          =   312
            Index           =   4
            Left            =   120
            TabIndex        =   111
            Top             =   1200
            Width           =   1212
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Compromiso"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   6360
            TabIndex        =   110
            Top             =   3600
            Width           =   1092
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuota"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   8
            Left            =   6360
            TabIndex        =   109
            Top             =   3240
            Width           =   612
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Asignado a la Operación #"
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
            Index           =   39
            Left            =   120
            TabIndex        =   108
            Top             =   3240
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificación Crediticia"
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
            Index           =   40
            Left            =   120
            TabIndex        =   107
            Top             =   3600
            Width           =   2415
         End
         Begin VB.Label LblTasa 
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   6360
            TabIndex        =   106
            Top             =   2760
            Width           =   972
         End
         Begin VB.Label LblPlazo 
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
            Height          =   252
            Left            =   6360
            TabIndex        =   105
            Top             =   2400
            Width           =   972
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   16
            Left            =   6360
            TabIndex        =   104
            Top             =   2040
            Width           =   1092
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1332
         Index           =   0
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   9852
         _Version        =   1572864
         _ExtentX        =   17378
         _ExtentY        =   2350
         _StockProps     =   79
         Caption         =   "Datos Personales"
         ForeColor       =   8421504
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
         BorderStyle     =   1
         Begin VB.ComboBox cboSexo 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPreaEstudio.frx":151B
            Left            =   3240
            List            =   "frmPreaEstudio.frx":1525
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   768
            Width           =   1455
         End
         Begin VB.ComboBox cboCantidadFiadores 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPreaEstudio.frx":153E
            Left            =   8880
            List            =   "frmPreaEstudio.frx":1545
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   768
            Width           =   972
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFecNac 
            Height          =   312
            Left            =   6240
            TabIndex        =   77
            Top             =   768
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtCedula 
            Height          =   312
            Left            =   1440
            TabIndex        =   78
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   360
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   312
            Left            =   3240
            TabIndex        =   79
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   360
            Width           =   6612
            _Version        =   1572864
            _ExtentX        =   11663
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nacimiento"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   38
            Left            =   5040
            TabIndex        =   83
            Top             =   792
            Width           =   972
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Genero"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   37
            Left            =   2280
            TabIndex        =   82
            Top             =   792
            Width           =   852
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Identificación"
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
            Height          =   324
            Index           =   3
            Left            =   120
            TabIndex        =   81
            Top             =   408
            Width           =   1212
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Fiadores"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   43
            Left            =   7920
            TabIndex        =   80
            Top             =   795
            Width           =   735
         End
      End
      Begin VB.ComboBox cboSalario 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPreaEstudio.frx":154C
         Left            =   -67480
         List            =   "frmPreaEstudio.frx":154E
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   5535
         Left            =   -70000
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   10455
         _Version        =   1572864
         _ExtentX        =   18441
         _ExtentY        =   9763
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
         Item(0).Caption =   "Notas"
         Item(0).ControlCount=   5
         Item(0).Control(0)=   "optObservacion(0)"
         Item(0).Control(1)=   "optObservacion(1)"
         Item(0).Control(2)=   "optObservacion(2)"
         Item(0).Control(3)=   "cmdGuardaObservaciones"
         Item(0).Control(4)=   "txtObservaciones"
         Item(1).Caption =   "Resolución"
         Item(1).ControlCount=   8
         Item(1).Control(0)=   "txtActa"
         Item(1).Control(1)=   "Label1(9)"
         Item(1).Control(2)=   "txtActaFecha"
         Item(1).Control(3)=   "Label1(35)"
         Item(1).Control(4)=   "lswAutorizadores"
         Item(1).Control(5)=   "rbActas(0)"
         Item(1).Control(6)=   "rbActas(1)"
         Item(1).Control(7)=   "rbActas(2)"
         Begin XtremeSuiteControls.ListView lswAutorizadores 
            Height          =   4215
            Left            =   -69880
            TabIndex        =   5
            Top             =   1320
            Visible         =   0   'False
            Width           =   10215
            _Version        =   1572864
            _ExtentX        =   18018
            _ExtentY        =   7435
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
            Checkboxes      =   -1  'True
            View            =   3
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optObservacion 
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   147
            Top             =   600
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Analistas de Crédito"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbActas 
            Height          =   252
            Index           =   0
            Left            =   -68080
            TabIndex        =   6
            Top             =   960
            Visible         =   0   'False
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Resoluciones"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton cmdGuardaObservaciones 
            Height          =   615
            Left            =   1320
            TabIndex        =   7
            Top             =   4800
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "&Guardar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   21
            Picture         =   "frmPreaEstudio.frx":1550
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtObservaciones 
            Height          =   4815
            Left            =   3000
            TabIndex        =   8
            Top             =   600
            Width           =   7095
            _Version        =   1572864
            _ExtentX        =   12515
            _ExtentY        =   8493
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtActa 
            Height          =   312
            Left            =   -68080
            TabIndex        =   9
            Top             =   480
            Visible         =   0   'False
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtActaFecha 
            Height          =   312
            Left            =   -64360
            TabIndex        =   10
            Top             =   480
            Visible         =   0   'False
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.RadioButton rbActas 
            Height          =   252
            Index           =   1
            Left            =   -65680
            TabIndex        =   11
            Top             =   960
            Visible         =   0   'False
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Autorizadores"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton rbActas 
            Height          =   252
            Index           =   2
            Left            =   -63520
            TabIndex        =   12
            Top             =   960
            Visible         =   0   'False
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Asistencia"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optObservacion 
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   148
            Top             =   1080
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Resolución del Comité"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.RadioButton optObservacion 
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   149
            Top             =   1560
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Junta Directiva"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin VB.Label Label1 
            Caption         =   "No. Acta"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   9
            Left            =   -69280
            TabIndex        =   14
            Top             =   480
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   35
            Left            =   -65560
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   1572
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   2652
         Left            =   -68320
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   6612
         _Version        =   524288
         _ExtentX        =   11663
         _ExtentY        =   4678
         _StockProps     =   64
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScrollBars      =   0
         SpreadDesigner  =   "frmPreaEstudio.frx":1C81
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCumplimientoNotas 
         Height          =   1872
         Left            =   -68320
         TabIndex        =   16
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3552
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1572864
         _ExtentX        =   11663
         _ExtentY        =   3302
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
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGridTags 
         Height          =   5415
         Left            =   -70000
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   10335
         _Version        =   524288
         _ExtentX        =   18230
         _ExtentY        =   9551
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
         MaxCols         =   5
         SpreadDesigner  =   "frmPreaEstudio.frx":224D
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.RadioButton optCausas 
         Height          =   492
         Index           =   0
         Left            =   -67480
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Causas para Denegación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
      Begin XtremeSuiteControls.RadioButton optCausas 
         Height          =   492
         Index           =   1
         Left            =   -64840
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Pendientes para Estudio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtSalarioReal 
         Height          =   312
         Left            =   -67480
         TabIndex        =   21
         Top             =   2040
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtExtrasFijas 
         Height          =   312
         Left            =   -67480
         TabIndex        =   22
         Top             =   2400
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDevengadoMes 
         Height          =   312
         Left            =   -67480
         TabIndex        =   23
         Top             =   2760
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotal_Cargas_CCSS 
         Height          =   312
         Left            =   -67480
         TabIndex        =   24
         Top             =   3480
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcSobreSalario 
         Height          =   312
         Left            =   -67480
         TabIndex        =   25
         Top             =   3840
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDeducciones 
         Height          =   312
         Left            =   -67480
         TabIndex        =   26
         Top             =   4200
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCrdTransitoCancelados 
         Height          =   312
         Left            =   -67480
         TabIndex        =   27
         Top             =   4560
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCrdTransitoXCobrar 
         Height          =   312
         Left            =   -67480
         TabIndex        =   28
         Top             =   4920
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFianzas 
         Height          =   312
         Left            =   -62680
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLiquidezSinFianza 
         Height          =   312
         Left            =   -61960
         TabIndex        =   30
         Top             =   3000
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLiquidezPorcSinFianza 
         Height          =   312
         Left            =   -61960
         TabIndex        =   31
         Top             =   3360
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLiquidezConFianza 
         Height          =   312
         Left            =   -61960
         TabIndex        =   32
         Top             =   3840
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLiquidezPorcConFianza 
         Height          =   312
         Left            =   -61960
         TabIndex        =   33
         Top             =   4200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPSD 
         Height          =   312
         Left            =   -61960
         TabIndex        =   34
         Top             =   4560
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMontoGirar 
         Height          =   312
         Left            =   -62680
         TabIndex        =   35
         Top             =   4920
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuotaDiferencia 
         Height          =   312
         Left            =   -62680
         TabIndex        =   36
         Top             =   5280
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRebajoExtras 
         Height          =   312
         Left            =   -67480
         TabIndex        =   54
         Top             =   1680
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalLiquido 
         Height          =   312
         Left            =   -62680
         TabIndex        =   55
         Top             =   1680
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSalarioDevengado 
         Height          =   312
         Left            =   -67480
         TabIndex        =   58
         Top             =   1320
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -67480
         TabIndex        =   60
         Top             =   960
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
      Begin XtremeSuiteControls.FlatEdit txtRefundiciones 
         Height          =   312
         Left            =   -62680
         TabIndex        =   62
         Top             =   960
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDesembolsos 
         Height          =   312
         Left            =   -62680
         TabIndex        =   64
         Top             =   1320
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSalarioLiquido 
         Height          =   312
         Left            =   -62680
         TabIndex        =   68
         Top             =   600
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   312
         Index           =   0
         Left            =   -65320
         TabIndex        =   138
         ToolTipText     =   "Cambio de Evaluador"
         Top             =   1680
         Visible         =   0   'False
         Width           =   312
         _Version        =   1572864
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmPreaEstudio.frx":2912
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   312
         Index           =   1
         Left            =   -65320
         TabIndex        =   139
         ToolTipText     =   "Cambio de Evaluador"
         Top             =   3480
         Visible         =   0   'False
         Width           =   312
         _Version        =   1572864
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmPreaEstudio.frx":3032
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   312
         Index           =   2
         Left            =   -65320
         TabIndex        =   140
         ToolTipText     =   "Cambio de Evaluador"
         Top             =   4200
         Visible         =   0   'False
         Width           =   312
         _Version        =   1572864
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmPreaEstudio.frx":3752
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   312
         Index           =   3
         Left            =   -60520
         TabIndex        =   141
         ToolTipText     =   "Cambio de Evaluador"
         Top             =   960
         Visible         =   0   'False
         Width           =   312
         _Version        =   1572864
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmPreaEstudio.frx":3E72
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   312
         Index           =   4
         Left            =   -60520
         TabIndex        =   142
         ToolTipText     =   "Cambio de Evaluador"
         Top             =   1320
         Visible         =   0   'False
         Width           =   312
         _Version        =   1572864
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmPreaEstudio.frx":4592
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   312
         Index           =   5
         Left            =   -60520
         TabIndex        =   143
         ToolTipText     =   "Cambio de Evaluador"
         Top             =   2400
         Visible         =   0   'False
         Width           =   312
         _Version        =   1572864
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmPreaEstudio.frx":4CB2
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   312
         Index           =   6
         Left            =   -65320
         TabIndex        =   144
         ToolTipText     =   "Cambio de Evaluador"
         Top             =   4560
         Visible         =   0   'False
         Width           =   312
         _Version        =   1572864
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmPreaEstudio.frx":53D2
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnDetalle 
         Height          =   312
         Index           =   7
         Left            =   -65320
         TabIndex        =   145
         ToolTipText     =   "Cambio de Evaluador"
         Top             =   4920
         Visible         =   0   'False
         Width           =   312
         _Version        =   1572864
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmPreaEstudio.frx":5AF2
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalLiquidoGrupo 
         Height          =   312
         Left            =   -62680
         TabIndex        =   152
         Top             =   2040
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "T. Liquido Grupo"
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
         Height          =   252
         Index           =   0
         Left            =   -64360
         TabIndex        =   153
         ToolTipText     =   "Es el Total Liquido del Deudor + Co Deudores"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   6240
         X2              =   6000
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   5640
         X2              =   5640
         Y1              =   3120
         Y2              =   3720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   6000
         X2              =   5640
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   6000
         X2              =   6000
         Y1              =   3120
         Y2              =   4320
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   6240
         X2              =   6000
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   6240
         X2              =   6000
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   6240
         X2              =   6000
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label1 
         Caption         =   "Salario Liquido"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   24
         Left            =   -64360
         TabIndex        =   70
         Top             =   600
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Salario"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   -69520
         TabIndex        =   69
         Top             =   600
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         Caption         =   "(+) Desembolsos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   29
         Left            =   -64360
         TabIndex        =   65
         Top             =   1320
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "(+) Refundiciones"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   25
         Left            =   -64360
         TabIndex        =   63
         Top             =   960
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte de Colilla"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   18
         Left            =   -69520
         TabIndex        =   61
         Top             =   960
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label lblSalarioDevengado 
         BackStyle       =   0  'Transparent
         Caption         =   "Salario Devengado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   9
         Left            =   -69520
         TabIndex        =   59
         Top             =   1320
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Rebajo de Extras"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   10
         Left            =   -69520
         TabIndex        =   57
         Top             =   1680
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         Caption         =   "T. Liquido Persona"
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
         Height          =   252
         Index           =   26
         Left            =   -64360
         TabIndex        =   56
         Top             =   1680
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Salario Real"
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
         Height          =   252
         Index           =   11
         Left            =   -69520
         TabIndex        =   53
         Top             =   2040
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Extras Fijas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   12
         Left            =   -69520
         TabIndex        =   52
         Top             =   2400
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Devengado del Mes"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   13
         Left            =   -69520
         TabIndex        =   51
         Top             =   2760
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label lblPorcentajeSalario 
         BackStyle       =   0  'Transparent
         Caption         =   "(%?) Sobre Salario"
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
         Height          =   252
         Left            =   -69520
         TabIndex        =   50
         Top             =   3840
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Cargas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   19
         Left            =   -69520
         TabIndex        =   49
         Top             =   3480
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Deducciones"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   20
         Left            =   -69520
         TabIndex        =   48
         Top             =   4200
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(+) Créditos Cancelados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   21
         Left            =   -69520
         TabIndex        =   47
         Top             =   4560
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Créditos x Cobrar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   22
         Left            =   -69520
         TabIndex        =   46
         Top             =   4920
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Image imgCuotaDif 
         Height          =   240
         Left            =   -60520
         Stretch         =   -1  'True
         ToolTipText     =   "Causas"
         Top             =   5280
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblMontoGirar 
         Caption         =   "Diferencia Cuota.:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   -64360
         TabIndex        =   45
         Top             =   5280
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "P.S.D."
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
         Height          =   252
         Index           =   42
         Left            =   -64360
         TabIndex        =   44
         Top             =   4560
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "Fianzas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   27
         Left            =   -64360
         TabIndex        =   43
         Top             =   2400
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "Liquidez"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   28
         Left            =   -64360
         TabIndex        =   42
         Top             =   2760
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "Sin Fianzas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   30
         Left            =   -63640
         TabIndex        =   41
         Top             =   3000
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Con Fianzas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   31
         Left            =   -63640
         TabIndex        =   40
         Top             =   3840
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "[%] Sin Fianzas"
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
         Index           =   33
         Left            =   -63640
         TabIndex        =   39
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "[%] Con Fianzas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   32
         Left            =   -63640
         TabIndex        =   38
         Top             =   4200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblMontoGirar 
         Caption         =   "Monto a Girar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   34
         Left            =   -64360
         TabIndex        =   37
         Top             =   4920
         Visible         =   0   'False
         Width           =   1572
      End
   End
   Begin VB.Timer Timer 
      Left            =   9240
      Top             =   600
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4939
            MinWidth        =   4939
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   5010
            MinWidth        =   5010
            Object.ToolTipText     =   "Fecha"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   5644
            MinWidth        =   5644
            Object.ToolTipText     =   "Oficina"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListtblGestion 
      Left            =   8520
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":6212
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":6BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":7189
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":798E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":82F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":8A8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":944E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":9A2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":A234
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":AA39
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":B228
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPreaEstudio.frx":BB9C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tbl_Desplazamiento 
      Height          =   330
      Left            =   6840
      TabIndex        =   1
      Top             =   840
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageListtblGestion"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Siguiente"
            Object.ToolTipText     =   "Seguiente"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   264
      Left            =   2040
      TabIndex        =   66
      Top             =   480
      Width           =   3708
      _ExtentX        =   6535
      _ExtentY        =   476
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
            Key             =   "BORRAR"
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
   Begin XtremeSuiteControls.PushButton btnSolicitado 
      Height          =   312
      Left            =   360
      TabIndex        =   73
      ToolTipText     =   "Poner en Estado de Solicitado"
      Top             =   36
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Volver a Solicitado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmPreaEstudio.frx":C332
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnCopiar 
      Height          =   312
      Left            =   2160
      TabIndex        =   74
      ToolTipText     =   "Copiar Expediente"
      Top             =   36
      Width           =   972
      _Version        =   1572864
      _ExtentX        =   1714
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Copia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmPreaEstudio.frx":CA32
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnGestion 
      Height          =   312
      Index           =   0
      Left            =   3600
      TabIndex        =   132
      ToolTipText     =   "Copiar Expediente"
      Top             =   36
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Causas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmPreaEstudio.frx":D122
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnGestion 
      Height          =   312
      Index           =   1
      Left            =   4800
      TabIndex        =   133
      ToolTipText     =   "Copiar Expediente"
      Top             =   36
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Etiquetas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmPreaEstudio.frx":D829
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnGestion 
      Height          =   312
      Index           =   2
      Left            =   6600
      TabIndex        =   134
      ToolTipText     =   "Copiar Expediente"
      Top             =   36
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Gestión"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmPreaEstudio.frx":DF42
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnGestion 
      Height          =   312
      Index           =   3
      Left            =   7920
      TabIndex        =   135
      ToolTipText     =   "Copiar Expediente"
      Top             =   40
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Solicitud de Crédito"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmPreaEstudio.frx":E669
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtExpediente 
      Height          =   400
      Left            =   2040
      TabIndex        =   136
      Top             =   840
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   706
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.ComboBox cboSubExpediente 
      Height          =   384
      Left            =   4080
      TabIndex        =   146
      Top             =   840
      Width           =   2652
      _Version        =   1572864
      _ExtentX        =   4683
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Left            =   360
      TabIndex        =   137
      Top             =   840
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Expediente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaptionTitle 
      Height          =   382
      Left            =   0
      TabIndex        =   131
      Top             =   0
      Width           =   12732
      _Version        =   1572864
      _ExtentX        =   22458
      _ExtentY        =   674
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
   End
   Begin XtremeShortcutBar.ShortcutCaption lblEstado 
      Height          =   375
      Left            =   5400
      TabIndex        =   117
      ToolTipText     =   "Estado del Estudio"
      Top             =   1440
      Width           =   5055
      _Version        =   1572864
      _ExtentX        =   8916
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Estado del Estudio"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption lblEdad 
      Height          =   375
      Left            =   0
      TabIndex        =   116
      ToolTipText     =   "Edad para Jubilación"
      Top             =   7920
      Width           =   10455
      _Version        =   1572864
      _ExtentX        =   18441
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Edad para Jubilación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption lblEstadoSocio 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Estado de la Persona"
      Top             =   1440
      Width           =   5415
      _Version        =   1572864
      _ExtentX        =   9551
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Estado de la Persona:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Alignment       =   1
   End
End
Attribute VB_Name = "frmPreaEstudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vPaso As Boolean

Dim m_curValor_Anterior As Currency ' Variable para comparar en los txt si vario el dato

Dim mFrecuenciaPago As String

Private clsMensajes As New ProGrX_EstudioCrd.clsEstudioMensajes
Private clsEntidad As New ProGrX_EstudioCrd.clsEntidad
Private clsNull As New ProGrX_EstudioCrd.clsNull

Public Item_Seleccionado As XtremeSuiteControls.ListViewItem
Public litem As XtremeSuiteControls.ListViewItem

Private m_ventanaEnModo As eVentanaEnModo
Private vCodExpediente As String
Private m_CambioDatos As Boolean
Private m_CambioCalculo As Boolean
Private m_CambioObservaciones As Boolean
Private m_valorComboExp As String
Private m_Cargando As Boolean
Private m_Paso As Boolean
Private m_CargoSalario As Boolean

Private m_MuestraMensaje As Boolean 'Control de mensajes  para el usuario
Private m_Expediente As String 'Numero de expediente actual consultado.
Private m_expedienteAnterior As String 'Almacena de forma temporal el expediente anterior consultado.
Private m_PreviousTab As Integer 'Mantiene el tab anterior.
Private m_FiadoresRegistrador As Integer ' Almacena la cantidad de fiadores por expedientes
Private m_estadoPreanalisis As String 'Mantiene el estado del preanalisis
Private m_DesplegoMensaje As Boolean 'Controla la despliegue de mensajes de información al usuario
Private m_CargoCombo As Boolean 'Para que no cargue el combo nuevamente

Private m_SoloVerSalarios As Boolean ' Variable que indica si en la lista de salarios solo es para consulta.

Dim m_SalarioDevengadoGrupo As Currency

Dim vPasoCarga As Boolean 'Para que Bancos Click cbo lo ignore
Dim m_NumPagos As Integer

Private m_FECHA_CREACION As String 'Para el calulo de colilla

Dim pTabBefore As Integer

Private m_Formulas As eFormulas

'Rutinas y funciones del forms

Private Sub DespliegaEdad(ByVal Anos As Integer, ByVal Meses As Integer, Optional ByVal Dias As Integer)
Dim tMeses As String, tAnos As String, tDias As String
' Para dar formato a las fechas
On Error GoTo vError
    
    tAnos = ""
    If (Anos = 1) Then
        tAnos = Anos & " Año "
    Else
        If (Anos > 1) Then
            tAnos = Anos & " Años "
        End If
    End If

    tMeses = ""
    If (Meses = 1) Then
        tMeses = Meses & " Mes "
    Else
        If (Meses > 1) Then
            tMeses = Meses & " Meses "
        End If
    End If

    tDias = Empty
    If Dias = 1 Then
        tDias = Dias & " Día"
    Else
        If (Dias > 1) Then
            tDias = Dias & " Días"
        End If
    End If

    lblEdad.Caption = tAnos & tMeses & tDias
    

    Exit Sub
vError:
    MsgBox "Ocurrió un error al ejecutar el despliegue de la edad. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub


Public Function fxCalculaEdadAnos(ByVal pFechaNacimiento As String, ByVal ValorReturn As String) As Integer
Dim MesesOri As Integer
Dim FechaMasMeses As Date
Dim FechaActual As Date
Dim FechaNacimiento As Date
Dim fAnios As Integer
Dim fMeses As Integer
Dim Diasint As Integer

fxCalculaEdadAnos = 0

On Error GoTo vError
   
   FechaNacimiento = Format(pFechaNacimiento, "dd/mm/yyyy")
   If m_FECHA_CREACION = "-1" Or m_FECHA_CREACION = "" Then
       FechaActual = Format(clsEntidad.fxTraerFechaServidor, "dd/mm/yyyy")
   Else
       FechaActual = Format(m_FECHA_CREACION, "dd/mm/yyyy")
   End If
   
   MesesOri = DateDiff("m", FechaNacimiento, FechaActual)
   FechaMasMeses = DateAdd("m", MesesOri, FechaNacimiento)
   If Format(FechaMasMeses, "dd/mm/yyyy") > FechaActual Then
       MesesOri = MesesOri - 1
       FechaMasMeses = DateAdd("m", MesesOri, FechaNacimiento)
   End If
   
   fAnios = Int(MesesOri / 12)
   fMeses = MesesOri - (fAnios * 12)
   Diasint = DateDiff("d", FechaMasMeses, FechaActual)
  
    Call DespliegaEdad(fAnios, fMeses, Diasint)
    
    If ValorReturn = "D" Then
    fxCalculaEdadAnos = Diasint
    ElseIf ValorReturn = "M" Then
     fxCalculaEdadAnos = fMeses
    ElseIf ValorReturn = "A" Then
     fxCalculaEdadAnos = fAnios
    End If

    Exit Function
vError:
    MsgBox "Ocurrió un error al calcular la edad. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Function


Public Sub sbEstructuraActualiza(vNivel As Integer, MontoGirar As Boolean)
    '' Procedimiento para actualizar formulas y campos de la forma
    
    vNivel = vNivel + 1
    
    ''Nivel 1 compuesto por: Salario Devengado y Rebajo de Extras
    ''Cualquier cambio en N1 debe recalcular los niveles a partir de N2
    
    ''Nivel 2 compuesto por: Salario Real y Extras fijas
    ''Cualquier cambio en N2 debe recalcular los niveles a partir de N3

    
    'Si es Asociado Aplicar la deducción de la Asociacion
    If gPreAnalisis.Socio = "S" Then
       chkCargaAsociacion.Value = vbChecked
    End If
    
    If vNivel <= 2 Then
    
        Call sbAplicarFormulas(eFormulas.eSalarioReal)
        
    End If
    
    ''Nivel 3 compuesto por: devengado del mes, cargas sociales, porcentaje salario,
    '' deducciones, créditos cancelados, créditos por cobrar
    ''Cualquier cambio en N3 debe recalcular los niveles a partir de N4
    
    If vNivel <= 3 Then
    
        Call sbAplicarFormulas(eFormulas.eDevengadoDelMes)
        Call SbReCalculaCargasCCSS
        Call sbAplicarFormulas(eFormulas.ePorcentajeSobreSalario)
        
    End If
    
    ''Nivel 4 compuesto por: salario liquido, refundiciones, desembolsos
    ''Cualquier cambio en N4 debe recalcular los niveles a partir de N5
    
    If vNivel <= 4 Then
        
        Call sbAplicarFormulas(eFormulas.eSalarioLiquido)
        
    End If
    
    ''Nivel 5 compuesto por: Total liquido, fianzas, nueva couta
    ''Cualquier cambio en N5 debe recalcular los niveles a partir de N6
    
    If vNivel <= 5 Then
        
        Call sbAplicarFormulas(eFormulas.eTotalLiquido)
        
    End If
    
    ''Nivel 6 compuesto por: liquidez sin fianza, liquidez con fianza
    ''este nivel no recalcula ningún siguiente nivel
    
    
    If vNivel <= 6 Then
    
        Call sbAplicarFormulas(eFormulas.eLiquidezConFianza)
        Call sbAplicarFormulas(eFormulas.eLiquidezSinFianzas)
        Call sbAplicarFormulas(eFormulas.eLiquidezPorcConFianza)
        Call sbAplicarFormulas(eFormulas.eLiquidezPorcSinFianzas)
 
        
    End If
    
    '' Nivel exclusivo para recalcular el monto a girar
    
    If MontoGirar = True Then '' Calcula Monto a Girar
        
        Call sbAplicarFormulas(ePolizaSD)
        Call sbAplicarFormulas(eFormulas.eMontoGirar)
        
    End If


End Sub



Public Sub sbCalcularPlazoMaximo()
Dim v_AnosEdad As Double
Dim v_MesesEdad As Double
Dim v_EdadMax As Integer
Dim v_Meses As Integer
Dim v_Dias As Integer

On Error GoTo vError

    v_Meses = 0
    v_Dias = 0
    
    'If InStr(1, TxtExpediente.Text, "-", vbTextCompare) > 0 Then Exit Sub
    
    v_AnosEdad = fxCalculaEdadAnos(dtpFecNac.Value, "A")
    
    v_MesesEdad = fxCalculaEdadAnos(dtpFecNac.Value, "M")
    
    v_AnosEdad = v_AnosEdad + (v_MesesEdad / 12)
    
    If fxSexoItemData(cboSexo.ListIndex) = "M" Then
    
        v_EdadMax = GlobalEdadMaximaPermitidaHombre
        
    ElseIf fxSexoItemData(cboSexo.ListIndex) = "F" Then
    
        v_EdadMax = GlobalEdadMaximaPermitidaMujeres
        
    End If
    
    txtPlMax.Text = CStr(CInt((v_EdadMax - v_AnosEdad) * 12))
    
    If (v_AnosEdad + (Val(txtPlazo.Text) / 12)) >= v_EdadMax Then
    
        lblEdad.Caption = lblEdad.Caption & " >> La edad supera el límite autorizado << "
    Else
        
        lblEdad.Caption = lblEdad.Caption & " >> La edad es satisfactoria << "
    End If
        
    Exit Sub
vError:
    MsgBox "Ocurrió un error al cacular plazo máximo. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Public Sub sbPosicionFrameSalarios()

On Error GoTo vError

    FrameSalarios.Top = 480

    FrameSalarios.Left = 1920

    Exit Sub
vError:
    MsgBox "Ocurrió un error al inicializar posición de la ventana de Salarios. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Public Sub sbPosicionFrameCargas()
On Error GoTo vError

    fraDCargas.Top = 2880
    fraDCargas.Left = 3900

    Exit Sub

vError:
    MsgBox "Ocurrió un error al inicializar posición de la ventana de cargas. " & "-" & Err.Description, vbExclamation, gMsgTitulo

    
End Sub

Public Sub sbAplicarFormulas(vFormula As eFormulas)
Dim vPorc As Double
Dim v_Expediente As Boolean
Dim v_Compromiso As Integer
Dim vCodigo As String
Dim vMontoFianzas As Double, vTotalLiquido As Currency, vDevengado As Currency

On Error GoTo vError

If vPaso Then Exit Sub



vMontoFianzas = 0
vTotalLiquido = 0
vDevengado = 0

vGrid.Col = 1
vGrid.Row = 2
vCodigo = vGrid.Text

If Val(cboCantidadFiadores.Text) = 0 Then
    v_Compromiso = 1
End If

If InStr(1, txtExpediente.Text, "-", vbTextCompare) = 0 Then
    v_Expediente = True
    v_Compromiso = 1
Else
    v_Compromiso = Val(cboCantidadFiadores.Text)
End If

If cboSubExpediente.Text = "Nuevo Expediente" Or cboSubExpediente.Text = "Nuevo SubExpediente" Then Exit Sub


Me.MousePointer = vbHourglass



'Revisión e inicialización de campos vacíos
If Val(txtRefundiciones.ToolTipText) = 0 Then
    txtRefundiciones.ToolTipText = Format(0, "Standard")
End If
If Val(txtDesembolsos.ToolTipText) = 0 Then
    txtDesembolsos.ToolTipText = Format(0, "Standard")
End If

If txtSalarioLiquido.Text = "" Then
    txtSalarioLiquido.Text = Format(0, "Standard")
End If
If txtRefundiciones.ToolTipText = "" Then
    txtRefundiciones.ToolTipText = Format(0, "Standard")
End If
If txtDesembolsos.ToolTipText = "" Then
    txtDesembolsos.ToolTipText = Format(0, "Standard")
End If

If txtSalarioDevengado.Text = "" Then
    txtSalarioDevengado.Text = Format(0, "Standard")
End If

If txtRebajoExtras.Text = "" Then
    txtRebajoExtras.Text = Format(0, "Standard")
End If

If Val(txtExtrasFijas.Text) = 0 Then
    txtExtrasFijas.Text = Format(0, "Standard")
End If

If txtSalarioReal.Text = "" Then
    txtSalarioReal.Text = Format(0, "Standard")
End If
If txtExtrasFijas = "" Then
     txtExtrasFijas.Text = Format(0, "Standard")
End If

If txtDevengadoMes.Text = "" Then
    txtDevengadoMes.Text = Format(0, "Standard")
End If
If txtCrdTransitoCancelados.Text = "" Then
    txtCrdTransitoCancelados.Text = Format(0, "Standard")
End If
If txtTotal_Cargas_CCSS.Text = "" Then
    txtTotal_Cargas_CCSS.Text = Format(0, "Standard")
End If
If txtDeducciones.Text = "" Then
    txtDeducciones.Text = Format(0, "Standard")
End If
If txtCrdTransitoXCobrar.Text = "" Then
    txtCrdTransitoXCobrar.Text = Format(0, "Standard")
End If
If Val(txtCompromiso.Text) = 0 Then
    txtCompromiso.Text = Format(0, "Standard")
End If
If txtTotalLiquido.Text = "" Then
    txtTotalLiquido.Text = Format(0, "Standard")
End If
If txtCompromiso.Text = "" Then
    txtCompromiso.Text = Format(0, "Standard")
End If
If txtLiquidezSinFianza.Text = "" Then
    txtLiquidezSinFianza.Text = Format(0, "Standard")
End If

If Val(txtFianzas.Text) = 0 Then
    txtFianzas.Text = Format(0, "Standard")
End If

If txtLiquidezConFianza.Text = "" Then
    txtLiquidezConFianza.Text = Format(0, "Standard")
End If

If txtMonto.Text = "" Then
    txtMonto.Text = Format(0, "Standard")
End If
If txtCuota.Text = "" Then
    txtCuota.Text = Format(0, "Standard")
End If
If txtPSD.Text = "" Then
    txtPSD.Text = Format(0, "Standard")
End If

'Total Liquido de Grupo
    txtTotalLiquido.Text = Format(CDbl(txtSalarioLiquido.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText), "Standard")
    
    vTotalLiquido = CCur(txtTotalLiquido.Text)
    vDevengado = CCur(txtDevengadoMes.Text)
    
    If IsNumeric(txtTotalLiquidoGrupo.Text) Then
        If (txtTotalLiquidoGrupo.Text) > vTotalLiquido Then
            vTotalLiquido = CCur(txtTotalLiquidoGrupo.Text)
        End If
    
        If m_SalarioDevengadoGrupo > vDevengado Then
            vDevengado = m_SalarioDevengadoGrupo
        End If
    
    End If

'Aplicación de Formulas
Select Case vFormula

    Case eFormulas.eSalarioReal
        txtSalarioReal.Text = Format(CCur(txtSalarioDevengado.Text) - CCur(txtRebajoExtras.Text), "Standard")
        
        
    Case eFormulas.eDevengadoDelMes
        
        If cboSalario.Text <> "" Then
            If Left(Right(cboSalario.Text, 2), 1) = "e" Then
                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * m_NumPagos) + CDbl(txtExtrasFijas.Text), "Standard")
            Else
                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1) + CDbl(txtExtrasFijas.Text), "Standard")
            End If
        Else
            txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1) + CDbl(txtExtrasFijas.Text), "Standard")
        End If

          
    Case eFormulas.ePorcentajeSobreSalario
        If IsNumeric(GlobalPorcLiquidezLibre) Then
            txtPorcSobreSalario.Text = Format((CDbl(txtDevengadoMes.Text) * GlobalPorcLiquidezLibre / 100), "Standard")
        Else
            MsgBox "El parámetro Porcentaje de Liquidez Libre no es un valor numérico.", vbExclamation, gMsgTitulo
        End If
    
    Case eFormulas.eSalarioLiquido
        If Right(cboSalario.Text, 3) <> "(g)" Then
            txtSalarioLiquido.Text = Format((CDbl(txtDevengadoMes.Text) + CDbl(txtCrdTransitoCancelados.Text)) - (CDbl(txtTotal_Cargas_CCSS.Text) + CDbl(txtDeducciones.Text) + CDbl(txtCrdTransitoXCobrar.Text)), "Standard")
        End If


       
    Case eFormulas.eTotalLiquido
                
        txtTotalLiquido.Text = Format(CDbl(txtSalarioLiquido.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText), "Standard")
        
        vTotalLiquido = CCur(txtTotalLiquido.Text)
        
        
        
        If IsNumeric(txtTotalLiquidoGrupo.Text) Then
            If v_Compromiso = 1 Then
                txtTotalLiquidoGrupo.Text = Format(vTotalLiquido, "Standard")
            End If
            
            If (txtTotalLiquidoGrupo.Text) > vTotalLiquido Then
                vTotalLiquido = CCur(txtTotalLiquidoGrupo.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText)
            End If
        End If
        
    
    Case eFormulas.eLiquidezSinFianzas
        
        txtLiquidezSinFianza.Text = Format(vTotalLiquido - (CDbl(txtCompromiso.Text) / v_Compromiso), "Standard")
        
    Case eFormulas.eLiquidezPorcSinFianzas
        
        If ((Val(txtDevengadoMes.Text) = 0) And (Val(txtLiquidezSinFianza.Text) > 0 Or Val(txtLiquidezSinFianza.Text) < 0)) Then
            txtLiquidezPorcSinFianza.Text = 0
        Else
           If CCur(txtDevengadoMes.Text) > 0 Then
                txtLiquidezPorcSinFianza.Text = Format((CDbl(txtLiquidezSinFianza.Text) / vDevengado) * 100, "Standard")
           Else
                txtLiquidezPorcSinFianza.Text = 0
           End If
        End If
        
    Case eFormulas.eLiquidezConFianza
        
        vMontoFianzas = CDbl(txtFianzas.Text)
       txtLiquidezConFianza.Text = Format(vTotalLiquido - ((CDbl(txtCompromiso.Text) / v_Compromiso) + CDbl(vMontoFianzas)), "Standard")
       
    
    Case eFormulas.eLiquidezPorcConFianza
        'Cálculo del [%] Con Fianzas
        If ((Val(txtLiquidezConFianza.Text) = 0) And vDevengado > 0) Then
            txtLiquidezPorcConFianza.Text = 0
        Else
            If vDevengado > 0 Then
                txtLiquidezPorcConFianza.Text = Format((CDbl(txtLiquidezConFianza.Text) / vDevengado) * 100, "Standard")
            End If
        End If
        Call CargarGrid ' Actualiza clasificacion en la ventana.
        
    Case eFormulas.eMontoGirar
       'Calcula monto a girar solo para deudores
        txtMontoGirar.Text = Format(0, "Standard")
        
        If v_Expediente Then
            If Val(txtPSD.Text) = 0 Then
           txtPSD.Text = 0
            End If
                    'TODO Los Intereses dias se establecen como 0 (cero) posteriormente se debe de cambiar
            If chkPrimerCuota.Value = vbChecked Then
                txtMontoGirar.Text = Format(CDbl(txtMonto.Text) - (0 + CDbl(txtRefundiciones.Text) + CDbl(txtDesembolsos.Text) + CDbl(txtCuota.Text) + CDbl(txtPSD.Text)), "Standard")
            Else
                txtMontoGirar.Text = Format(CDbl(txtMonto.Text) - (0 + CDbl(txtRefundiciones.Text) + CDbl(txtDesembolsos.Text) + CDbl(txtPSD.Text)), "Standard")
            End If
        End If
        
    Case eFormulas.ePolizaSD
        'Calcula solo para los expedientes la Póliza saldo deudor
        txtPSD.Text = 0
        
        If txtMonto.Text = "" Then
            txtMonto.Text = Format(0, "Standard")
        End If
        
        If v_Expediente Then
            txtPSD.Text = Format((CDbl(txtMonto.Text) * GlobalPorcPSD) / 100, "Standard")
        End If
        
'******************Todas las formulas****************************************************

    Case eFormulas.eAplicarTodas
        'Aplicada todas la formulas en orden
        txtSalarioReal.Text = Format(CCur(txtSalarioDevengado.Text) - CCur(txtRebajoExtras.Text), "Standard")
        
        If cboSalario.Text <> "" Then
            If Left(Right(cboSalario.Text, 2), 1) = "e" Then
                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * m_NumPagos) + CDbl(txtExtrasFijas.Text), "Standard")
            Else
                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1) + CDbl(txtExtrasFijas.Text), "Standard")
            End If
        Else
            txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1) + CDbl(txtExtrasFijas.Text), "Standard")
        End If
        
        'Fin Calcula devengado del mes
        If Right(cboSalario.Text, 3) <> "(g)" Then
            txtSalarioLiquido.Text = Format((CDbl(txtDevengadoMes.Text) + CDbl(txtCrdTransitoCancelados.Text)) - (CDbl(txtTotal_Cargas_CCSS.Text) + CDbl(txtDeducciones.Text) + CDbl(txtCrdTransitoXCobrar.Text)), "Standard")
        End If
        
        'Calculo del Total Liquido y Salario Devengado (Base)
        txtTotalLiquido.Text = Format(CDbl(txtSalarioLiquido.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText), "Standard")
        
        vTotalLiquido = CCur(txtTotalLiquido.Text)
        vDevengado = CCur(txtDevengadoMes.Text)
        
        If IsNumeric(txtTotalLiquidoGrupo.Text) Then
            If (txtTotalLiquidoGrupo.Text) > vTotalLiquido Then
                vTotalLiquido = CCur(txtTotalLiquidoGrupo.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText)
            End If
        
            If m_SalarioDevengadoGrupo > vDevengado Then
                vDevengado = m_SalarioDevengadoGrupo
            End If
        
        End If
        
        
        txtLiquidezSinFianza.Text = Format(vTotalLiquido - (CDbl(txtCompromiso.Text) / v_Compromiso), "Standard")
        
        If Not ((vDevengado = 0) And (Val(txtLiquidezSinFianza.Text) > 0 Or Val(txtLiquidezSinFianza.Text) < 0)) Then
            txtLiquidezPorcSinFianza.Text = Format((CDbl(txtLiquidezSinFianza.Text) / vDevengado) * 100, "Standard")
        Else
            txtLiquidezPorcSinFianza.Text = 0
        End If
        
        'Calcula Liquiedez Con Fianzas
        If Val(txtCompromiso.Text) = 0 Then
            txtCompromiso.Text = Format(0, "Standard")
        End If
        If Val(txtFianzas.Text) = 0 Then
            txtFianzas.Text = Format(0, "Standard")
        End If
        vMontoFianzas = CDbl(txtFianzas.Text)
        
       txtLiquidezConFianza.Text = Format(vTotalLiquido - ((CDbl(txtCompromiso.Text) / v_Compromiso) + CDbl(vMontoFianzas)), "Standard")

        
        'Calcula Liquiedez [%] Con Fianzas
        If Not ((Val(txtLiquidezConFianza.Text) = 0) And vDevengado > 0) Then
            If CDbl(txtDevengadoMes.Text) > 0 Then
                txtLiquidezPorcConFianza.Text = Format((CDbl(txtLiquidezConFianza.Text) / vDevengado) * 100, "Standard")
            Else
                txtLiquidezPorcConFianza.Text = 0
            End If
        Else
           txtLiquidezPorcConFianza.Text = 0
        End If
        
        Call CargarGrid ' Actualiza clasificacion en la ventana.
        
        'Calcula solo para los expedientes la Póliza saldo deudor
        txtPSD.Text = 0
        If v_Expediente Then
            txtPSD.Text = Format((CDbl(txtMonto.Text) * GlobalPorcPSD) / 100, "Standard")
        End If
        
        'TODO Los Intereses dias se establecen como 0 (cero) posteriormente se debe de cambiar
         'Calcula monto a girar solo para deudores
        txtMontoGirar.Text = Format(0, "Standard")
        If v_Expediente Then
            If Val(txtPSD.Text) = 0 Then
               txtPSD.Text = 0
            End If
            
            'TODO Los Intereses dias se establecen como 0 (cero) posteriormente se debe de cambiar
            If chkPrimerCuota.Value = vbChecked Then
                txtMontoGirar.Text = Format(CDbl(txtMonto.Text) - (0 + CDbl(txtRefundiciones.Text) + CDbl(txtDesembolsos.Text) + CDbl(txtCuota.Text) + CDbl(txtPSD.Text)), "Standard")
            Else
                txtMontoGirar.Text = Format(CDbl(txtMonto.Text) - (0 + CDbl(txtRefundiciones.Text) + CDbl(txtDesembolsos.Text) + CDbl(txtPSD.Text)), "Standard")
            End If
        End If
        
    End Select


'Calcula Cuota diferencia para la Persona
txtTotalLiquido.Text = Format(CDbl(txtSalarioLiquido.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText), "Standard")

txtCuotaDiferencia.Text = Format(CCur(txtCuota.Text) - (CCur(txtRefundiciones.ToolTipText) + CCur(txtDesembolsos.ToolTipText)), "Standard")
If CCur(txtCuotaDiferencia.Text) > 0 Then
   txtCuotaDiferencia.ForeColor = vbRed
   Set imgCuotaDif.Picture = ImageListtblGestion.ListImages.Item(5).Picture
Else
   txtCuotaDiferencia.ForeColor = vbBlack
   Set imgCuotaDif.Picture = ImageListtblGestion.ListImages.Item(4).Picture

End If
txtCuotaDiferencia.Text = Format(Abs(CCur(txtCuotaDiferencia.Text)), "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler("Ocurrió un error al aplicar las formulas. " & "-" & Err.Description), vbExclamation, gMsgTitulo

End Sub

Public Sub SbCalculaCargasCCSSPorDefecto()
Dim vTotalSuma As Double
Dim vCargaAsociacion As Double
Dim vCargaCCSS As Double
Dim vCargaImpSalario As Double

On Error GoTo vError


vCargaAsociacion = 0
vCargaCCSS = 0
vCargaImpSalario = 0

If Val(txtDevengadoMes.Text) = 0 Then
    txtDevengadoMes.Text = 0
End If
If gPreAnalisis.Socio = "S" Then
    If IsNumeric(GlobalPorcAsocSolidarista) Then
        vCargaAsociacion = Format((GlobalPorcAsocSolidarista * txtDevengadoMes.Text) / 100, "Standard")
    End If
    
End If

If IsNumeric(GlobalPorcCCSS) Then
        vCargaCCSS = Format((GlobalPorcCCSS * txtDevengadoMes.Text) / 100, "Standard")
End If

lblCargaImpSalario.Caption = 0
If Val(txtDevengadoMes.Text) > 0 Then
'    glogon.strSQL = "select dbo.fxCRDPreaCalculaRenta (" & fxFormatearValor(CDbl(txtDevengadoMes.Text), Numerico) & ")"
'    If (execSql(glogon.strSQL, True)) Then
'        If glogon.Recordset(0) & "" = "" Then Exit Sub
'        vCargaImpSalario = Format(glogon.Recordset(0), "Standard")
'    End If
    
    vCargaImpSalario = fxRentaCalculo(CCur(txtDevengadoMes.Text))
    
End If

    txtTotal_Cargas_CCSS.Text = Format(vCargaAsociacion + vCargaCCSS + vCargaImpSalario, "Standard")
    
If Right(cboSalario.Text, 3) = "(g)" Then
    Call sbActCtlConstExternos(False, "(g)")
Else
    Call sbActCtlConstExternos(True, "(d)")
End If

Exit Sub

vError:
    MsgBox "Ocurrió un error al calcular las cargas sociales. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Public Sub SbReCalculaCargasCCSS()
Dim vTotalSuma As Double
Dim vCargaAsociacion As Double
Dim vCargaCCSS As Double
Dim vCargaImpSalario As Double
Dim vCargaFrap As Double


On Error GoTo vError

If vPaso Then Exit Sub


vCargaAsociacion = 0
vCargaCCSS = 0
vCargaImpSalario = 0
vCargaFrap = 0

If Val(txtDevengadoMes.Text) = 0 Then
    txtDevengadoMes.Text = 0
End If

If IsNumeric(GlobalPorcAsocSolidarista) Then
     vCargaAsociacion = Format((GlobalPorcAsocSolidarista * txtDevengadoMes.Text) / 100, "Standard")
End If
    
If IsNumeric(GlobalPorcCCSS) Then
    vCargaCCSS = Format((GlobalPorcCCSS * txtDevengadoMes.Text) / 100, "Standard")
End If

If IsNumeric(GlobalPorcFRAPFAP) Then
    vCargaFrap = Format((GlobalPorcFRAPFAP * txtDevengadoMes.Text) / 100, "Standard")
End If

lblCargaImpSalario.Caption = 0
If Val(txtDevengadoMes.Text) > 0 Then
    vCargaImpSalario = fxRentaCalculo(CCur(txtDevengadoMes.Text))
'    glogon.strSQL = "select dbo.fxCRDPreaCalculaRenta (" & fxFormatearValor(CDbl(txtDevengadoMes.Text), Numerico) & ")"
'    If (execSql(glogon.strSQL, True)) Then
'        If glogon.Recordset(0) & "" = "" Then Exit Sub
'        vCargaImpSalario = Format(glogon.Recordset(0), "Standard")
'    End If
End If

If chkCargaAsociacion.Value = Checked Then
    vTotalSuma = vCargaAsociacion
End If

If chkCargaFrap.Value = Checked Then
    vTotalSuma = vTotalSuma + vCargaFrap
End If
    
vTotalSuma = vTotalSuma + vCargaImpSalario

vTotalSuma = vTotalSuma + vCargaCCSS

txtTotal_Cargas_CCSS.Text = Format(vTotalSuma, "Standard")
    
If Right(cboSalario.Text, 3) = "(g)" Then
    Call sbActCtlConstExternos(False, "(g)")
Else
    Call sbActCtlConstExternos(True, "(d)")
    
End If
    
Exit Sub

vError:
    MsgBox "Ocurrió un error al calcular las cargas sociales. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub

Public Sub SbCalculaCargasCCSS()
On Error GoTo vError

If IsNumeric(GlobalPorcAsocSolidarista) Then
    lblCargaAsociacion.Caption = Format((GlobalPorcAsocSolidarista * txtDevengadoMes.Text) / 100, "Standard")
End If
    
If clsMensajes.Estado = "N" Then
    chkCargaAsociacion.Enabled = True
End If

If chkCargaFrap.Tag = "S" Then
    chkCargaFrap.Value = Checked
    chkCargaFrap.Tag = "S"
Else
    chkCargaFrap.Value = 0
    chkCargaFrap.Tag = "N"
End If

If IsNumeric(GlobalPorcFRAPFAP) Then
    lblCargaFrap.Caption = Format((GlobalPorcFRAPFAP * txtDevengadoMes.Text) / 100, "Standard")
End If

If IsNumeric(GlobalPorcCCSS) Then
    lblCargaCCSS.Caption = Format((GlobalPorcCCSS * txtDevengadoMes.Text) / 100, "Standard")
End If

lblCargaImpSalario.Caption = Format(0, "Standard")

If Val(txtDevengadoMes.Text) > 0 Then
'    glogon.strSQL = "select dbo.fxCRDPreaCalculaRenta (" & fxFormatearValor(CDbl(txtDevengadoMes.Text), Numerico) & ")"
'    If (execSql(glogon.strSQL, True)) Then
'        If glogon.Recordset(0) & "" = "" Then Exit Sub
'        lblCargaImpSalario.Caption = Format(glogon.Recordset(0), "Standard")
'    End If

    lblCargaImpSalario.Caption = Format(fxRentaCalculo(CCur(txtDevengadoMes.Text)), "Standard")

End If

Exit Sub

vError:
    MsgBox "Ocurrió un error al calcular las cargas sociales. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Public Sub SbBloquearTxtSalario(ByVal pCodigo As String)
On Error GoTo vError

clsEntidad.tablaName = "spCRDPreaTIPO_SALARIO"
txtSalarioDevengado.Locked = True
txtExtrasFijas.Locked = True

btnDetalle.Item(1).Visible = False

If clsEntidad.fxTraerUno(fxFormatearValor(pCodigo, caracter)) Then
    If glogon.Recordset!MODIFICA_DEVENGADO = 1 Then
       txtSalarioDevengado.Locked = False
    End If
    
    If glogon.Recordset!MODIFICA_REBAJO_EXTRAS = 1 Then
       btnDetalle.Item(1).Visible = True
    End If
    
    If glogon.Recordset!MODIFICA_EXTRAS_FIJAS = 1 Then
       txtExtrasFijas.Locked = False
    End If

End If

Exit Sub

vError:
  MsgBox "Ocurrió un error desbloquear campos. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Public Sub sbBloquearTab()

On Error GoTo vError

If tcMain.SelectedItem = 0 Then
    If Len(txtNombre.Text) = 0 Then
        
        tcMain.Item(0).Enabled = True
        tcMain.Item(1).Enabled = False
        tcMain.Item(2).Enabled = False
        tcMain.Item(3).Enabled = False
        tcMain.Item(4).Enabled = False
    Else
        tcMain.Item(1).Enabled = True
        tcMain.Item(2).Enabled = True
        tcMain.Item(3).Enabled = True
        tcMain.Item(4).Enabled = True
    End If
End If

Exit Sub
vError:
   MsgBox "Ocurrió un error desbloquear campos. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Public Sub sbBloquearControles(ByRef forms As Form, ByVal Expediente As eTipoExpediente)
'Esta rutina se encarga de inicializar los valores de los controles que se encuentra pegados en la patalla
Dim vControl As Control, vValor As Boolean

On Error GoTo vError

If (Expediente = SubExpediente) Then
    vValor = False
Else
    vValor = True
End If

For Each vControl In forms
    
        Select Case vControl.Name
            Case "txtLinea"
                vControl.Enabled = vValor
            Case "txtDesLineaCredito"
                vControl.Enabled = vValor
            
            Case "txtPolizaVida"
                vControl.Enabled = vValor
            Case "txtPolizaIncendio"
                vControl.Enabled = vValor
            Case "txtPolizaDesempleo"
                vControl.Enabled = vValor
            
            Case "txtMonto"
                vControl.Enabled = vValor
            Case "txtPlazo"
                vControl.Enabled = vValor
            Case "txtTasa"
                vControl.Enabled = vValor
            
            Case "cboCantidadFiadores"
                vControl.Enabled = vValor
            Case "cboGarantia"
                vControl.Enabled = vValor
            Case "cboDestino"
                vControl.Enabled = vValor
            Case "cboFondo"
                vControl.Enabled = vValor
            Case "cboComite"
                vControl.Enabled = vValor
            
            Case "chkPolizaVida"
                vControl.Enabled = vValor
            
            Case "chkPolizaIncendio"
                vControl.Enabled = vValor
            
            Case "chkPolizaDesempleo"
                vControl.Enabled = vValor
            
            Case "chkPrimerCuota"
                vControl.Enabled = vValor
    End Select
Next vControl

chkCargaAsociacion.Enabled = True

Exit Sub

vError:
   MsgBox "Ocurrió un error desbloquear campos. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Private Sub sbLigarDatos(ByVal rs As ADODB.Recordset)
Dim Codigo As String
Dim Item As String

On Error GoTo vError

With rs
'Información del Tab Datos


m_CargoSalario = True
txtFianzas.Text = Format(0, "Standard")

clsMensajes.TASA_PTS_BONO = 0

txtCedula.Text = !Cedula & ""

DoEvents
Call txtCedula_LostFocus
txtNombre.Text = !Nombre & ""
gPreAnalisis.Expediente = txtExpediente.Text

clsMensajes.Estado = Trim(!Estado & "")

gPreAnalisis.Estado = Trim(!Estado & "")

If clsMensajes.Estado = "P" Then
    lblEstado.Caption = "Pendiente"
    
ElseIf clsMensajes.Estado = "R" Then
    lblEstado.Caption = "Recibido"
ElseIf clsMensajes.Estado = "A" Then
    lblEstado.Caption = "Aprobado"
ElseIf clsMensajes.Estado = "D" Then
    lblEstado.Caption = "Denegado"
End If

m_estadoPreanalisis = clsMensajes.Estado

Call sbBloquearTab

Call sbSeleccionaSexo(cboSexo, Trim(!sexo & ""))

stBar.Panels(1).Text = "US: " & !Usuario
stBar.Panels(2).Text = "CR: " & Format(!FECHA_CREACION, "dd-mm-yyyy")
stBar.Panels(3).Text = !OFICINA

m_FECHA_CREACION = Format(!FECHA_CREACION, "dd-mm-yyyy")

If Trim(!fecha_nacimiento & "") <> "" Then
    dtpFecNac.Value = !fecha_nacimiento
End If
If Trim(!NSUB_EXP & "") <> "" Then
    cboCantidadFiadores.Text = !NSUB_EXP
Else
    cboCantidadFiadores.Text = 0
End If

If Trim(!Cod_Linea & "") <> "" Then
    txtLinea.Text = !Cod_Linea
    Call txtLinea_LostFocus
End If
If Len(txtDesLineaCredito.Text) > 0 Then
    Call sbSTCargaCboGarantia(cboGarantia, txtLinea.Text)
End If

If Trim(!GARANTIA & "") <> "" Then
    Call sbCboAsignaDato(cboGarantia, fxGarantia(!GARANTIA), True, 0)
End If


If Trim(!ID_COMITE & "") <> "" Then
    Call sbCboAsignaDato(cboComite, fxComite(!ID_COMITE), True, !ID_COMITE)
Else
    cboComite.Text = " "
End If

If Trim(!cod_destino & "") <> "" Then
    Call sbCboAsignaDato(cboDestino, Trim(!DescDestino), True, 0)
End If

chkPrimerCuota.Value = !apl_primer_cuota

chkPolizaVida.Value = !APL_POLIZA_VIDA
chkPolizaIncendio.Value = !apl_poliza_incendio
chkPolizaDesempleo.Value = !APL_POLIZA_DESEMPLEO

txtPolizaVida.Text = IIf(Val(!MONTO_POLIZA_VIDA & "") = clsNull.NullNumerico, 0, Format(!MONTO_POLIZA_VIDA, "Standard"))
txtPolizaIncendio.Text = IIf(Val(!MONTO_POLIZA_INCENDIO) = clsNull.NullNumerico, 0, Format(!MONTO_POLIZA_INCENDIO, "Standard"))
txtPolizaDesempleo.Text = IIf(Val(!MONTO_POLIZA_DESEMPLEO) = clsNull.NullNumerico, 0, Format(!MONTO_POLIZA_DESEMPLEO, "Standard"))

txtMonto.Text = IIf(Val(!Monto & "") = clsNull.NullNumerico, 0, Format(!Monto, "Standard"))
txtPlazo.Text = IIf(Val(!Plazo & "") = clsNull.NullNumerico, 0, !Plazo)
txtTasa.Text = IIf(Val(!TASA & "") = clsNull.NullNumerico, 0, Format(!TASA, "Standard"))

txtCumplimientoNotas.Text = !CUMPLIMIENTO_NOTAS & ""

If !TASA_PTS_BONO > 0 Then
    txtTasa.ToolTipText = "Bono por Membresia de " & Format(!TASA_PTS_BONO, "Standard")
Else
    txtTasa.ToolTipText = Empty
End If

txtCuota.Text = IIf(Val(!Cuota & "") = clsNull.NullNumerico, 0, Format(!Cuota, "Standard"))
txtCompromiso.Text = IIf(Val(!COMPROMISO & "") = clsNull.NullNumerico, 0, Format(!COMPROMISO, "Standard"))
txtAsignado.Text = IIf(Val(Trim(!ID_SOLICITUD & "")) = clsNull.NullNumerico, 0, !ID_SOLICITUD)
Call sbCalcularPlazoMaximo

'Información del Tab Calculo

'Call sbSeleccionarItemCombo(cboSalario, Trim(!tipo_salario & ""))

If Trim(!tipo_salario & "") <> "" Then
    
    Call sbAsignaTipoSalario(Trim(!DescTipoSalario))
    
End If

Codigo = fxDeCodificaPrimaryKey(cboSalario.Text, 1, "-")

Call SbBloquearTxtSalario(Trim(Codigo))

If !FECHA_CORTE_COLIILA & "" <> "" Then
    dtpCorte.Value = !FECHA_CORTE_COLIILA
End If

Item = Left(Right(cboSalario.Text, 2), 1)
    
'Call sbHabilitaFechaColilla(Trim(Item))


txtSalarioDevengado.Text = IIf(Val(!SALARIO_DEVENGADO_COLILLA & "") = clsNull.NullNumerico, 0, Format(!SALARIO_DEVENGADO_COLILLA, "Standard"))

txtSalarioDevengado.ToolTipText = IIf(Val(!SALARIO_DEVENGADO_GRUPO & "") = clsNull.NullNumerico, 0, Format(!SALARIO_DEVENGADO_GRUPO, "Standard"))
m_SalarioDevengadoGrupo = IIf(Val(!SALARIO_DEVENGADO_GRUPO & "") = clsNull.NullNumerico, 0, Format(!SALARIO_DEVENGADO_GRUPO, "Standard"))

txtRebajoExtras.Text = IIf(Val(!REBAJO_EXTRAS & "") = clsNull.NullNumerico, 0, Format(!REBAJO_EXTRAS, "Standard"))

txtSalarioReal.Text = IIf(Val(!SALARIO_REAL & "") = clsNull.NullNumerico, 0, Format(!SALARIO_REAL, "Standard"))

txtExtrasFijas.Text = IIf(Val(!EXTRAS_FIJAS & "") = clsNull.NullNumerico, 0, Format(!EXTRAS_FIJAS, "Standard"))

txtDevengadoMes.Text = IIf(Val(!DEVENGADO_MES & "") = clsNull.NullNumerico, 0, Format(!DEVENGADO_MES, "Standard"))

chkCargaAsociacion.Tag = "N"
clsMensajes.CARGA_ASOCIACION = 0
If !CARGA_ASOCIACION & "" <> "" Then
    clsMensajes.CARGA_ASOCIACION = !CARGA_ASOCIACION
    If clsMensajes.CARGA_ASOCIACION > 0 Then
        chkCargaAsociacion.Tag = "S"
        chkCargaAsociacion.Value = Checked
    End If
End If
If chkCargaAsociacion.Tag = "N" Then
    chkCargaAsociacion.Value = Unchecked
Else
    chkCargaAsociacion.Value = Checked
End If

chkCargaFrap.Tag = "N"
If !CARGA_FRAP & "" <> "" Then
    clsMensajes.CARGA_FRAP = !CARGA_FRAP
    If clsMensajes.CARGA_FRAP > 0 Then
        chkCargaFrap.Tag = "S"
        chkCargaFrap.Value = Checked
    End If
End If

txtTotal_Cargas_CCSS.Text = IIf(Val(!TOTAL_CARGA_CCSS & "") = clsNull.NullNumerico, 0, Format(!TOTAL_CARGA_CCSS, "Standard"))

txtPorcSobreSalario.Text = IIf(Val(!PORCENTAJE_LIBRE & "") = clsNull.NullNumerico, 0, Format(!PORCENTAJE_LIBRE, "Standard"))

txtDeducciones.Text = IIf(Val(!DEDUCCIONES & "") = clsNull.NullNumerico, 0, Format(!DEDUCCIONES, "Standard"))

txtCrdTransitoCancelados.Text = IIf(Val(!CRD_TRANSITO_CANCELADOS & "") = clsNull.NullNumerico, 0, Format(!CRD_TRANSITO_CANCELADOS, "Standard"))

txtCrdTransitoXCobrar.Text = IIf(Val(!CRD_TRANSITO_XCOBRAR & "") = clsNull.NullNumerico, 0, Format(!CRD_TRANSITO_XCOBRAR, "Standard"))

txtSalarioLiquido.Text = IIf(Val(!SALARIO_LIQUIDO & "") = clsNull.NullNumerico, 0, Format(!SALARIO_LIQUIDO, "Standard"))

txtRefundiciones.Text = IIf(Val(!REFUNDICIONES & "") = clsNull.NullNumerico, 0, Format(!REFUNDICIONES, "Standard"))

If Val(txtRefundiciones.Text) = 0 Then
    txtRefundiciones.ToolTipText = Format(0, "Standard")

Else
    txtRefundiciones.ToolTipText = IIf(Val(!REFUNDICIONES_CUOTA & "") = Format(clsNull.NullNumerico, "Standard"), 0, Format(!REFUNDICIONES_CUOTA, "Standard"))
End If

txtDesembolsos.ToolTipText = Format(0, "Standard")
txtDesembolsos.Text = IIf(Val(!DESEMBOLSOS & "") = clsNull.NullNumerico, 0, Format(!DESEMBOLSOS, "Standard"))
If Val(txtDesembolsos.Text) = 0 Then

    txtSalarioLiquido.Text = IIf(Val(!SALARIO_LIQUIDO & "") = clsNull.NullNumerico, 0, Format(!SALARIO_LIQUIDO, "Standard"))
Else
    txtDesembolsos.ToolTipText = IIf(Val(!DESEMBOLSOS_CUOTA & "") = Format(clsNull.NullNumerico, "Standard"), 0, Format(!DESEMBOLSOS_CUOTA, "Standard"))
End If


txtTotalLiquido.Text = IIf(Val(!LIQUIDO_TOTAL & "") = clsNull.NullNumerico, 0, Format(!LIQUIDO_TOTAL, "Standard"))
txtTotalLiquidoGrupo.Text = IIf(Val(!LIQUIDO_TOTAL_GRUPO & "") = clsNull.NullNumerico, 0, Format(!LIQUIDO_TOTAL_GRUPO, "Standard"))


txtFianzas.Text = IIf(Val(!FIANZAS & "") = clsNull.NullNumerico, 0, Format(!FIANZAS, "Standard"))
txtLiquidezSinFianza.Text = IIf(Val(!LIQUIDEZ_SIMPLE & "") = clsNull.NullNumerico, 0, Format(!LIQUIDEZ_SIMPLE, "Standard"))
txtLiquidezConFianza.Text = IIf(Val(!LIQUIDEZ_CFIANZAS & "") = clsNull.NullNumerico, 0, Format(!LIQUIDEZ_CFIANZAS, "Standard"))
'Datos del frame
lblCargaCCSS.Caption = IIf(Val(!CARGA_CCSS & "") = clsNull.NullNumerico, 0, Format(!CARGA_CCSS, "Standard"))
lblCargaImpSalario.Caption = IIf(Val(!CARGA_IMPUESTO_SALARIO & "") = clsNull.NullNumerico, 0, Format(!CARGA_IMPUESTO_SALARIO, "Standard"))
lblCargaAsociacion.Caption = IIf(Val(!CARGA_ASOCIACION & "") = clsNull.NullNumerico, 0, Format(!CARGA_ASOCIACION, "Standard"))
lblCargaAsociacion.Caption = IIf(Val(!CARGA_FRAP & "") = clsNull.NullNumerico, 0, Format(!CARGA_FRAP, "Standard"))

'Obtienen las obsevaciones
'& "" esto me valida si tiene nulos
vObservacion(0) = Trim(!OBSERVACION_ANALISTA & "") 'Observaciones de Analisis de crédito
vObservacion(1) = Trim(!OBSERVACION_COMITE & "") 'Observaciones de Resolución del comité
vObservacion(2) = Trim(!OBSERVACION_JD & "") 'Observaciones de Junta directiva
txtCumplimientoNotas.Text = Trim(!CUMPLIMIENTO_NOTAS & "")

End With

Call SbCalculaCargasCCSS
Call sbCalculaPolizaDeVida
Call sbCalculaPolizaDeIncendio
Call sbCalculaPolizaDesempleo

Call sbToolBar(Me.tlb, "edicion")
Call SbAccionVentana(ModificarRegistro)


Call sbTraerNumFiadores

 m_CambioDatos = False


    Exit Sub
vError:
    MsgBox "Ocurrió un error al mostrar la información consultada. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Public Sub sbAsignaTipoSalario(pDato As String)
On Error GoTo vError

        cboSalario.Text = Trim(pDato)

Exit Sub

vError:
  cboSalario.AddItem pDato
  cboSalario.Text = pDato
  
End Sub


Private Function fxExistenFiadores() As Boolean
Dim m_Valor As String
Dim Indicador As String
Dim sql As String
Dim ExpedientePadre As String
Dim m_codGarantia As String

On Error GoTo vError


m_FiadoresRegistrador = 0
fxExistenFiadores = True

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    MsgBox "Debe selecionar un expediente maestro.", vbInformation, gMsgTitulo
    fxExistenFiadores = False
    Exit Function
Else
    ExpedientePadre = txtExpediente.Text
End If

sql = "Select count(*) as NumFiadores from CRD_PREA_PREANALISIS where COD_PREANALISIS_REF = " & fxFormatearValor(ExpedientePadre, caracter)

If clsEntidad.fxEjecutaSQL(sql) Then
    m_FiadoresRegistrador = glogon.Recordset!NumFiadores
End If
m_codGarantia = fxGarantia(cboGarantia.Text)

If (m_codGarantia = "F") And Val(cboCantidadFiadores.Text) > m_FiadoresRegistrador Then
    MsgBox "No se han registrado todos fiadores indicados en el expediente maestro.", vbInformation, gMsgTitulo
    fxExistenFiadores = False
    
End If


    Exit Function
vError:
    MsgBox "Ocurrió un error traer los sub expediente registrados. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Function


Private Function fxValidaNumFiadoresRegistrados(Optional MuestreMensaje As Boolean) As Boolean
Dim m_Valor As String
Dim Indicador As String
Dim sql As String
Dim ExpedientePadre As String

On Error GoTo vError
fxValidaNumFiadoresRegistrados = True
m_FiadoresRegistrador = 0

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    ExpedientePadre = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
Else
    ExpedientePadre = txtExpediente.Text
End If

sql = "Select count(*) as NumFiadores from CRD_PREA_PREANALISIS where COD_PREANALISIS_REF = " & fxFormatearValor(ExpedientePadre, caracter)

If clsEntidad.fxEjecutaSQL(sql) Then
    m_FiadoresRegistrador = glogon.Recordset!NumFiadores
    clsMensajes.NSUB_EXP = m_FiadoresRegistrador
End If
If (m_FiadoresRegistrador + 1) > Val(cboCantidadFiadores.Text) Then
    If MuestreMensaje Then
        MsgBox "Para agregar un sub expediente debe aumentar la cantidad de fiadores en el expediente maestro.", vbInformation, gMsgTitulo
    End If
    fxValidaNumFiadoresRegistrados = False
End If

    Exit Function
vError:
    MsgBox "Ocurrió un error traer los sub expediente registrados. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    fxValidaNumFiadoresRegistrados = False

End Function

Private Sub sbTraerNumFiadores()
Dim m_Valor As String
Dim Indicador As String
Dim sql As String
Dim ExpedientePadre As String

On Error GoTo vError


If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    ExpedientePadre = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
Else
    ExpedientePadre = txtExpediente.Text
End If

sql = "Select NSUB_EXP from CRD_PREA_PREANALISIS where   COD_PREANALISIS = " & fxFormatearValor(ExpedientePadre, caracter)

If clsEntidad.fxEjecutaSQL(sql) Then
cboCantidadFiadores.Text = glogon.Recordset!NSUB_EXP
End If

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto digitado. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    

End Sub
    
Private Sub sbTraerDatosExpediente()
    
On Error GoTo vError

Dim strSQL As String, rs As New ADODB.Recordset
Dim VcboCantidadFiadores As Integer
Dim m_Valor As String
Dim Indicador As String

Me.MousePointer = vbDefault

If Len(txtExpediente.Text) = 0 Then Exit Sub

    'carga combo de salarios por si se agregó algún tipo de salario no activo
    cboSalario.Clear
    Call sbLlenarComboTodos(cboSalario, "spCRDPreaTIPO_SALARIO", "TIPO_SALARIO", "DescTipoSalario")

clsEntidad.tablaName = "spCRDPreaPREANALISIS"
m_expedienteAnterior = txtExpediente.Text

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    m_Valor = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
    Indicador = "S"
Else
    m_Valor = txtExpediente.Text
    Indicador = "E"
End If
        
If clsEntidad.fxTraerUno(fxFormatearValor(txtExpediente.Text, caracter)) Then

    Set rs = glogon.Recordset
    m_Cargando = True
    txtExpediente.Locked = False
    If m_CargoCombo = False Then
        cboSubExpediente.Clear
        cboSubExpediente.AddItem m_Valor
        Call sbLlenarComboFiltrado(cboSubExpediente, "spCRDPreaPREANALISIS", "COD_PREANALISIS", "COD_PREANALISIS", "SubExpediente", "", fxFormatearValor(m_Valor, caracter))
        cboSubExpediente.AddItem "Nuevo Expediente"
        cboSubExpediente.ItemData(cboSubExpediente.NewIndex) = -1
        VcboCantidadFiadores = IIf(Val(cboCantidadFiadores.Text) = 0, 1, Val(cboCantidadFiadores.Text))
        
'        (fxGarantia(cboGarantia.Text) = "F" And
        If (m_FiadoresRegistrador + 1) <= VcboCantidadFiadores Then
        
            cboSubExpediente.AddItem "Nuevo SubExpediente"
            cboSubExpediente.ItemData(cboSubExpediente.NewIndex) = -2
            
        End If

    End If
    Call sbLigarDatos(rs)

    If Indicador = "S" Then
        Call sbBloquearControles(Me, SubExpediente)
    Else
       Call sbBloquearControles(Me, Expediente)
    End If
Else
    Call sbToolBar(Me.tlb, "edicion")
    Call sbInicializaComboExpediente
    Call tlb_ButtonClick(tlb.Buttons("nuevo"))
    txtExpediente.Locked = False
End If


m_Cargando = False

Call dtpCorte_Change

DoEvents
Call sbSeleccionarItemComboExp(cboSubExpediente, m_expedienteAnterior)


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto digitado. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Private Sub sbTraerMaxExpediente()

Dim vpadre As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo vError

Set vRecordset = Nothing

If m_valorComboExp = "Nuevo SubExpediente" Then
        If InStr(1, txtExpediente.Text, "-", vbTextCompare) <> 0 Then
            vpadre = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
        Else
            vpadre = txtExpediente.Text
        End If
         clsEntidad.tablaName = "spCRDPreaMaxSubExpediente"
        If clsEntidad.fxTraerUno(fxFormatearValor(vpadre, caracter)) Then
        Set vRecordset = glogon.Recordset
            If Trim(vRecordset(0) & "") <> "" Then
                txtExpediente.Text = Trim(vRecordset(0) & "")
                m_Expediente = Trim(vRecordset(0) & "")
            End If
         
        End If
ElseIf InStr(1, txtExpediente.Text, "-", vbTextCompare) = 0 Then
    clsEntidad.tablaName = "spCRDPreaPARAMETROS"
    m_Expediente = clsEntidad.fxTraerValor("VALOR", "'11'")
    If m_Expediente <> -1 Then
        txtExpediente.Text = m_Expediente
    End If
    If m_valorComboExp = "Nuevo SubExpediente" Then
         clsEntidad.tablaName = "spCRDPreaMaxSubExpediente"
        If clsEntidad.fxTraerUno(fxFormatearValor(m_Expediente, caracter)) Then
        Set vRecordset = glogon.Recordset
            If Trim(vRecordset(0) & "") <> "" Then
                txtExpediente.Text = Trim(vRecordset(0) & "")
                m_Expediente = Trim(vRecordset(0) & "")
            End If
         
        End If
    End If
Else
    vpadre = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
    clsEntidad.tablaName = "spCRDPreaMaxSubExpediente"
    If clsEntidad.fxTraerUno(fxFormatearValor(vpadre, caracter)) Then
    Set vRecordset = glogon.Recordset
        If Trim(vRecordset(0) & "") <> "" Then
            txtExpediente.Text = Trim(vRecordset(0) & "")
            m_Expediente = Trim(vRecordset(0) & "")
        End If
     
    End If
    
End If


    Exit Sub
vError:
    MsgBox "Ocurrió un error consultar número de experiente registrado. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Private Sub SbAccionVentana(ByVal Tipo As eVentanaEnModo)

m_CambioDatos = False
m_CambioCalculo = False
m_CambioObservaciones = False

Select Case True
  Case Tipo = eVentanaEnModo.NuevoRegistro
        m_ventanaEnModo = eVentanaEnModo.NuevoRegistro
        tcMain.Item(0).Selected = True
        m_FiadoresRegistrador = 0
  
  Case Tipo = ModificarRegistro
        m_ventanaEnModo = eVentanaEnModo.ModificarRegistro

End Select

End Sub
    
Private Sub sbCargarCombos()
Dim strSQL As String, rs As New ADODB.Recordset

'Call sbSTCargaCboGarantia(cboGarantia, FormatearValor("ADB2", Caracter))
Call sbSTCargaCboGarantia(cboGarantia, "-1")
'Call sbLlenarComboTodos(cboDestino, "spCRDPreaDestinos", "cod_destino", "DescDestino", "Seleccione un destino")
Call sbSeleccionaSexo(cboSexo, "F")
Call sbLlenarComboTodos(cboSalario, "spCRDPreaTIPO_SALARIO", "TIPO_SALARIO", "DescTipoSalario")

Call sbCargaCboComites

 
dtpCorte.Value = fxFechaServidor
dtpFecNac.Value = dtpCorte.Value

'Carga Garantias de Fondos
strSQL = "exec spCRDGarantiaFND"
Call OpenRecordSet(rs, strSQL)

cboFondo.Clear
If rs.EOF And rs.BOF Then
  MsgBox "No existen Garantias en Fondos [Creadas], verifique...", vbCritical
Else
 Do While Not rs.EOF
   cboFondo.AddItem Trim(rs!GARANTIA) & " - " & rs!DESCRIPCION
   rs.MoveNext
 Loop
 rs.MoveFirst
 cboFondo.Text = Trim(rs!GARANTIA) & " - " & rs!DESCRIPCION
End If
rs.Close


End Sub

Private Sub sbInicializaComboExpediente()
    m_Cargando = True
    cboSubExpediente.Clear
    cboSubExpediente.AddItem "Nuevo Expediente"
    cboSubExpediente.ItemData(cboSubExpediente.NewIndex) = -1
    cboSubExpediente.ListIndex = 0
    clsMensajes.Estado = "P"
     m_Cargando = False
     
End Sub

Function fxDescLineaCredito(ByVal strCodigo As String) As String
On Error GoTo vError

glogon.strSQL = "select descripcion from catalogo where codigo = '" & Trim(strCodigo) & "'"

If execSql(glogon.strSQL, True) Then
    fxDescLineaCredito = IIf(IsNull(glogon.Recordset!DESCRIPCION), "", glogon.Recordset!DESCRIPCION)
Else
    MsgBox "No se encontró la descripción del código de la linea de crédito digitada. - " & strCodigo, vbCritical
End If


    Exit Function
vError:
    MsgBox "Ocurrió un error validar información digitada. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Function

Private Sub sbBusqueda(ByVal Control As String)
'Set GLOBALES.gfrmFormulario = Me
gBusquedas.Resultado = ""
gBusquedas.Convertir = "N"

Select Case Control
  Case "txtLinea" 'Codigo de linea de credito
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
        gBusquedas.Filtro = " and Activo = 1 and Retencion = 'N'"
        frmBusquedas.Show vbModal
        txtLinea.Text = gBusquedas.Resultado
        If Len(Trim(txtLinea.Text)) > 0 Then
          txtDesLineaCredito.Text = fxDescLineaCredito(Trim(txtLinea.Text))
        End If
   
  Case "txtDesLineaCredito" 'Descripcion Linea Credito
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Columna = "descripcion"
        frmBusquedas.Show vbModal
        txtLinea.Text = gBusquedas.Resultado
        txtLinea.Text = gBusquedas.Resultado
        If Len(Trim(txtLinea.Text)) > 0 Then
          txtDesLineaCredito.Text = fxDescLineaCredito(Trim(txtLinea.Text))
        End If

    Case "txtCedula"
        
        gBusquedas.Convertir = "N"
        gBusquedas.Col1Name = "Cédula Colilla"
        gBusquedas.Col2Name = "Cédula Real"
        gBusquedas.Col3Name = "Nombre"
        gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
        gBusquedas.Orden = "cedula"
        gBusquedas.Columna = "cedula"
        frmBusquedas.Show vbModal
        txtCedula.Text = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          Call txtCedula_LostFocus
        End If
        
        
        
    Case "txtNombre"
        gBusquedas.Convertir = "N"
        gBusquedas.Col1Name = "Cédula Colilla"
        gBusquedas.Col2Name = "Cédula Real"
        gBusquedas.Col3Name = "Nombre"
        gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
        gBusquedas.Orden = "nombre"
        gBusquedas.Columna = "nombre"
        frmBusquedas.Show vbModal
        txtCedula.Text = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          Call txtCedula_LostFocus
        End If
    Case "txtExpediente"
      Call sbMostraVentanBusqueda
End Select

End Sub
Private Function fxSelectItemSubExpediente(ByVal ListIndex As String) As String
Select Case Trim(ListIndex)
    Case "-1"
        fxSelectItemSubExpediente = "E"
    Case "-2"
        fxSelectItemSubExpediente = "S"
End Select
End Function

Public Function fxSexoItemData(ByVal ListIndex As Integer) As String
Select Case Trim(ListIndex)
    Case "0"
        fxSexoItemData = "M"
    Case "1"
        fxSexoItemData = "F"
End Select
End Function

Private Sub sbSeleccionaSexo(ByVal Combo As ComboBox, ByVal ItemData As String)
Select Case ItemData
    Case "M"
        Combo.ListIndex = 0
    Case "F"
        Combo.ListIndex = 1
        
End Select
End Sub

 
Private Function fxValidaDatos(ByVal TabValidar As Integer) As Boolean
Dim m_Valor As String

On Error GoTo vError

fxValidaDatos = True

If m_ventanaEnModo = ModificarRegistro Then
   If Trim(txtExpediente.Text) <> Trim(cboSubExpediente.Text) Then
    MsgBox "No es posible realizar cambios al expediente seleccionado.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    Exit Function
   End If
End If

If ((clsMensajes.Estado = "A") Or (clsMensajes.Estado = "D")) Then
    m_estadoPreanalisis = clsMensajes.Estado
    MsgBox "No es posible realizar cambios al expediente seleccionado.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    Exit Function
Else
    clsMensajes.Estado = "R"
End If

If Val(cboCantidadFiadores.Text) < m_FiadoresRegistrador Then
    MsgBox "No es posible disminuir la cantidad de sub expedientes."
    cboCantidadFiadores.Text = m_FiadoresRegistrador
        fxValidaDatos = False
    Exit Function
'ElseIf Val(cboCantidadFiadores.Text) > m_FiadoresRegistrador Then
'    MsgBox "No es posible aumentar la cantidad de sub expedientes."
'    cboCantidadFiadores.Text = m_FiadoresRegistrador
'        fxValidaDatos = False
'    Exit Function
End If

m_estadoPreanalisis = clsMensajes.Estado
clsMensajes.Usuario = glogon.Usuario

'Validar de Cantidad Fiadores
clsMensajes.NSUB_EXP = cboCantidadFiadores.ListIndex

'Validar número de Salario
If cboSalario.ListCount = 0 Then
    clsMensajes.tipo_salario = clsNull.SetNull
Else
    clsMensajes.tipo_salario = SIFGlobal.fxCodText(cboSalario.Text)
End If
    
If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    m_Valor = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
Else
    m_Valor = txtExpediente.Text
End If

clsMensajes.cod_preanalisis = IIf(Len(txtExpediente.Text) = 0, clsNull.SetNull, txtExpediente.Text)

If fxSelectItemSubExpediente(cboSubExpediente.ItemData(cboSubExpediente.ListIndex)) = "E" Then

    clsMensajes.tipo_preanalisis = "E"
    clsMensajes.cod_preanalisis_ref = clsNull.SetNull
    
ElseIf fxSelectItemSubExpediente(cboSubExpediente.ItemData(cboSubExpediente.ListIndex)) = "S" Then

    If m_Valor = "" Then
        clsMensajes.cod_preanalisis_ref = vCodExpediente
        clsMensajes.cod_preanalisis = "" 'vCodExpediente
    Else
        clsMensajes.cod_preanalisis_ref = m_Valor
        clsMensajes.cod_preanalisis = txtExpediente.Text
    End If
    
    clsMensajes.tipo_preanalisis = "S"
    
    
Else
    If InStr(1, txtExpediente.Text, "-", vbTextCompare) = 0 Then
        clsMensajes.tipo_preanalisis = "E"
        clsMensajes.cod_preanalisis_ref = clsNull.SetNull
    Else
        clsMensajes.tipo_preanalisis = "S"
        clsMensajes.cod_preanalisis_ref = m_Valor
        clsMensajes.cod_preanalisis = txtExpediente.Text
    End If
End If

'If TabValidar = Datos Then 'Valida datos del tab del tab de datos
    'Validar número de cédula
clsMensajes.Cedula = Trim(txtCedula.Text)
If Len(txtCedula.Text) = 0 Then
    MsgBox "El Número de cédula es requerido.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    clsMensajes.Cedula = ""
    txtCedula.SetFocus
    Exit Function
End If
'Validar nombre
clsMensajes.Nombre = Trim(txtNombre.Text)
If Len(txtNombre.Text) = 0 Then
    MsgBox "El Nombre es requerido.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    clsMensajes.Nombre = ""
    txtNombre.SetFocus
    Exit Function
End If

clsMensajes.sexo = fxSexoItemData(cboSexo.ListIndex)
'Validar fecha de nacimiento
clsMensajes.fecha_nacimiento = dtpFecNac.Value
If Not (IsDate(dtpFecNac.Value)) Then
    MsgBox "Fecha de nacimiento no es válida.", vbInformation, gMsgTitulo
    tcMain.Item(0).Selected = True
    fxValidaDatos = False
    clsMensajes.fecha_nacimiento = Date
    dtpFecNac.SetFocus
    Exit Function
End If

'Validar Linea de credito
If Len(txtDesLineaCredito.Text) = 0 Then
    MsgBox "Línea de crédito es requerida.", vbInformation, gMsgTitulo
  '  tcmain.Item(0).Selected =True
    fxValidaDatos = False
    txtLinea.SetFocus
    Exit Function
Else
 clsMensajes.Cod_Linea = txtLinea.Text
End If

'Validar Linea de garantias
clsMensajes.GARANTIA = fxGarantia(cboGarantia.Text) 'cboGarantia.ItemData(cboGarantia.ListIndex)
If cboGarantia.ListCount = 0 Then
    MsgBox "Debe seleccionar una garantía.", vbInformation, gMsgTitulo
    'tcmain.Item(0).Selected =True
    fxValidaDatos = False
    clsMensajes.GARANTIA = clsNull.SetNull
    cboGarantia.SetFocus
    Exit Function
End If


'Validar Linea de Comite
If cboComite.ItemData(cboComite.ListIndex) = 0 Then
    MsgBox "Debe seleccionar un comite de aprobación.", vbInformation, gMsgTitulo
    'tcmain.Item(0).Selected =True
    fxValidaDatos = False
    Exit Function
Else
    clsMensajes.ID_COMITE = cboComite.ItemData(cboComite.ListIndex)
End If




'Validar  respaldo segun garantía

clsMensajes.GARANTIA_FND = clsNull.SetNull

If fxGarantia(cboGarantia.Text) = "Y" Then
    clsMensajes.GARANTIA_FND = SIFGlobal.fxCodText(cboFondo.Text)
    If cboFondo.ListCount = 0 Then
        MsgBox "Debe seleccionar un respaldo de la lista.", vbInformation, gMsgTitulo
        'tcmain.Item(0).Selected =True
        fxValidaDatos = False
        clsMensajes.GARANTIA_FND = clsNull.SetNull
        cboFondo.SetFocus
        Exit Function
    End If
End If


'Validar Destino credito


If (cboDestino.ListCount = 0) Or (SIFGlobal.fxCodText(cboDestino.Text) = "") Then
    MsgBox "Debe seleccionar un destino de crédito.", vbInformation, gMsgTitulo
    
    fxValidaDatos = False
    clsMensajes.cod_destino = ""
    cboDestino.SetFocus
    Exit Function
Else
    clsMensajes.cod_destino = SIFGlobal.fxCodText(cboDestino.Text)
End If

clsMensajes.APL_POLIZA_VIDA = IIf(chkPolizaVida.Value = 0, 0, 1)
clsMensajes.MONTO_POLIZA_VIDA = IIf(Len(txtPolizaVida.Text) = 0, clsNull.NullNumerico, CDbl(Val(txtPolizaVida.Text)))
clsMensajes.apl_poliza_incendio = IIf(chkPolizaIncendio.Value = 0, 0, 1)
clsMensajes.MONTO_POLIZA_INCENDIO = IIf(Len(txtPolizaIncendio.Text) = 0, clsNull.NullNumerico, CDbl(Val(txtPolizaIncendio.Text)))
clsMensajes.apl_primer_cuota = chkPrimerCuota.Value
'Validar Monto  credito

If Val(txtMonto.Text) = 0 Then
    MsgBox "Debe digitar el monto del crédito.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    txtMonto.SetFocus
    Exit Function
Else
    clsMensajes.Monto = CDbl(txtMonto.Text)
End If
'Validar plazo Monto  credito

If Val(txtPlazo.Text) = 0 Then
    MsgBox "Debe digitar el plazo para el monto del crédito.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    txtPlazo.SetFocus
    Exit Function
Else
    clsMensajes.Plazo = CInt(Val(txtPlazo.Text))
End If
'Validar tasa del Monto  credito

If Val(txtTasa.Text) = 0 Then
    clsMensajes.TASA = 0
Else
    clsMensajes.TASA = CDbl(Val(txtTasa.Text))
End If

'Validar cuota
If Val(txtCuota.Text) = 0 Then
    MsgBox "La cuota no fue calculada correctamente.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    txtTasa.SetFocus
    Exit Function
Else
    clsMensajes.Cuota = CDbl(txtCuota.Text)
End If

If Val(txtCompromiso.Text) = 0 Then
    clsMensajes.COMPROMISO = clsNull.NullNumerico
Else
   clsMensajes.COMPROMISO = CDbl(txtCompromiso.Text)
End If

If TabValidar = 1 Then 'Valida datos en el tab de calculo

If cboSalario.ListCount = 0 Then
    clsMensajes.tipo_salario = clsNull.SetNull
Else
    clsMensajes.tipo_salario = SIFGlobal.fxCodText(cboSalario.Text)
End If

End If

clsMensajes.FECHA_CORTE_COLIILA = dtpCorte.Value
 
clsMensajes.SALARIO_DEVENGADO_COLILLA = Val(txtSalarioDevengado.Text)
If Val(txtSalarioDevengado.Text) <> 0 Then
    clsMensajes.SALARIO_DEVENGADO_COLILLA = CDbl(txtSalarioDevengado.Text)
    clsMensajes.SALARIO_DEVENGADO_GRUPO = CDbl(txtSalarioDevengado.Text)
End If

clsMensajes.REBAJO_EXTRAS = Val(txtRebajoExtras.Text)
If Val(txtRebajoExtras.Text) <> 0 Then
    clsMensajes.REBAJO_EXTRAS = CCur(txtRebajoExtras.Text)
End If
clsMensajes.REBAJO_EXTRAS = Val(txtRebajoExtras.Text)
If Val(txtRebajoExtras.Text) <> 0 Then
    clsMensajes.REBAJO_EXTRAS = CDbl(txtRebajoExtras.Text)
End If
clsMensajes.REBAJO_EXTRAS = Val(txtRebajoExtras.Text)
If Val(txtRebajoExtras.Text) <> 0 Then
    clsMensajes.REBAJO_EXTRAS = CDbl(txtRebajoExtras.Text)
End If
clsMensajes.SALARIO_REAL = Val(txtSalarioReal.Text)
If Val(txtSalarioReal.Text) <> 0 Then
    clsMensajes.SALARIO_REAL = CDbl(txtSalarioReal.Text)
End If
clsMensajes.EXTRAS_FIJAS = Val(txtExtrasFijas.Text)
If Val(txtExtrasFijas.Text) <> 0 Then
    clsMensajes.EXTRAS_FIJAS = CDbl(txtExtrasFijas.Text)
End If

clsMensajes.DEVENGADO_MES = Val(txtDevengadoMes.Text)
If Val(txtDevengadoMes.Text) <> 0 Then
    clsMensajes.DEVENGADO_MES = CDbl(txtDevengadoMes.Text)
End If

clsMensajes.PORCENTAJE_LIBRE = 0
If Val(txtPorcSobreSalario.Text) <> 0 Then
    clsMensajes.PORCENTAJE_LIBRE = CDbl(txtPorcSobreSalario.Text)
End If
clsMensajes.DEDUCCIONES = 0
If Val(txtDeducciones.Text) <> 0 Then
    clsMensajes.DEDUCCIONES = CDbl(txtDeducciones.Text)
End If
clsMensajes.CRD_TRANSITO_CANCELADOS = 0
If Val(txtCrdTransitoCancelados.Text) <> 0 Then
    clsMensajes.CRD_TRANSITO_CANCELADOS = CDbl(txtCrdTransitoCancelados.Text)
End If

clsMensajes.CRD_TRANSITO_XCOBRAR = 0
If Val(txtCrdTransitoXCobrar.Text) <> 0 Then
    clsMensajes.CRD_TRANSITO_XCOBRAR = CDbl(txtCrdTransitoXCobrar.Text)
End If
clsMensajes.SALARIO_LIQUIDO = Val(txtSalarioLiquido.Text)
If Val(txtSalarioLiquido.Text) <> 0 Then
    clsMensajes.SALARIO_LIQUIDO = CDbl(txtSalarioLiquido.Text)
End If
clsMensajes.REFUNDICIONES = Val(txtRefundiciones.Text)
If Val(txtRefundiciones.Text) <> 0 Then
    clsMensajes.REFUNDICIONES = CDbl(txtRefundiciones.Text)
End If

clsMensajes.REFUNDICIONES_CUOTA = Val(txtRefundiciones.ToolTipText)
If Val(txtRefundiciones.ToolTipText) <> 0 Then
    clsMensajes.REFUNDICIONES_CUOTA = CDbl(txtRefundiciones.ToolTipText)
End If


clsMensajes.DESEMBOLSOS = Val(txtDesembolsos.Text)
If Val(txtDesembolsos.Text) <> 0 Then
    clsMensajes.DESEMBOLSOS = CDbl(txtDesembolsos.Text)
End If

clsMensajes.DESEMBOLSOS_CUOTA = Val(txtDesembolsos.ToolTipText)
If Val(txtDesembolsos.ToolTipText) <> 0 Then
    clsMensajes.DESEMBOLSOS_CUOTA = CDbl(txtDesembolsos.ToolTipText)
End If

clsMensajes.LIQUIDO_TOTAL = Val(txtTotalLiquido.Text)
If Val(txtTotalLiquido.Text) <> 0 Then
    clsMensajes.LIQUIDO_TOTAL = CDbl(txtTotalLiquido.Text)
End If


clsMensajes.FIANZAS = 0
If Val(txtFianzas.Text) <> 0 Then
    clsMensajes.FIANZAS = CDbl(txtFianzas.Text)
End If
clsMensajes.LIQUIDEZ_SIMPLE = 0
If Val(txtLiquidezSinFianza.Text) <> 0 Then
    clsMensajes.LIQUIDEZ_SIMPLE = CDbl(txtLiquidezSinFianza.Text)
End If

clsMensajes.LIQUIDEZ_CFIANZAS = 0
If Val(txtLiquidezConFianza.Text) <> 0 Then
    clsMensajes.LIQUIDEZ_CFIANZAS = CDbl(txtLiquidezConFianza.Text)
End If


clsMensajes.TOTAL_CARGA_CCSS = 0
clsMensajes.CARGA_ASOCIACION = 0
chkCargaAsociacion.Tag = "N"
If chkCargaAsociacion.Value = 1 Then
    If Val(lblCargaAsociacion.Caption) > 0 Then
        clsMensajes.CARGA_ASOCIACION = CDbl(lblCargaAsociacion.Caption)
    End If
    chkCargaAsociacion.Tag = "S"
End If

chkCargaFrap.Tag = "N"
clsMensajes.CARGA_FRAP = 0
If chkCargaFrap.Value = 1 Then
    If Val(lblCargaFrap.Caption) > 0 Then
        clsMensajes.CARGA_FRAP = CDbl(lblCargaFrap.Caption)
        chkCargaFrap.Tag = "S"
    End If
End If

clsMensajes.CARGA_CCSS = Val(lblCargaCCSS.Caption)
If chkCargaAsociacion.Value = 1 Then
    If clsMensajes.CARGA_CCSS > 0 Then
        clsMensajes.CARGA_CCSS = CDbl(lblCargaCCSS.Caption)
    End If
End If
clsMensajes.TOTAL_CARGA_CCSS = Val(txtTotal_Cargas_CCSS.Text)
If Val(txtTotal_Cargas_CCSS.Text) <> 0 Then
    clsMensajes.TOTAL_CARGA_CCSS = CDbl(txtTotal_Cargas_CCSS.Text)
End If
   
    
clsMensajes.CARGA_IMPUESTO_SALARIO = Val(lblCargaImpSalario.Caption)
If clsMensajes.CARGA_IMPUESTO_SALARIO > 0 Then
    clsMensajes.CARGA_IMPUESTO_SALARIO = CDbl(lblCargaImpSalario.Caption)
End If


'ElseIf TabValidar = Observaciones Then 'Valida datos en el tab de calculo
clsMensajes.OBSERVACION_ANALISTA = vObservacion(0)
clsMensajes.OBSERVACION_COMITE = vObservacion(1)
clsMensajes.OBSERVACION_JD = vObservacion(2)

clsMensajes.COD_OFICINA = GLOBALES.gOficinaTitular
clsMensajes.CUMPLIMIENTO_NOTAS = txtCumplimientoNotas.Text

Call LigarDatosClasificacion

Exit Function

vError:
    MsgBox "Ocurrió un error validar la información digitada. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    fxValidaDatos = False

End Function




Private Function fxValidaDatosBorrar() As Boolean
    Dim m_Valor As String
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
fxValidaDatosBorrar = True

If ((clsMensajes.Estado = "A") Or (clsMensajes.Estado = "D")) Then
    m_estadoPreanalisis = clsMensajes.Estado
    MsgBox "No es posible realizar cambios al expediente seleccionado.", vbInformation, gMsgTitulo
    fxValidaDatosBorrar = False
    Exit Function
End If

If Len(txtExpediente.Text) = 0 Then
   MsgBox "Debe seleccionar un expediente.", vbInformation, gMsgTitulo
    fxValidaDatosBorrar = False
    Exit Function
End If


    strSQL = "select count(*) from CRD_PREA_PREANALISIS where COD_PREANALISIS_REF = '" & Trim(txtExpediente.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        If rs.Fields(0) > 0 Then
            MsgBox "El expediente tiene subexpedientes asociados y debe de borrados antes de borrar el principal"
            fxValidaDatosBorrar = False
            Exit Function
        End If
    End If

clsMensajes.cod_preanalisis = txtExpediente.Text
If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    clsMensajes.cod_preanalisis_ref = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
Else
    clsMensajes.cod_preanalisis_ref = txtExpediente.Text
    
End If


    Exit Function
    
vError:
    MsgBox "Ocurrió un error validar la información para eliminar los datos seleccionados. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    fxValidaDatosBorrar = False

End Function


Private Sub LigarDatosClasificacion()
Dim i As Integer

vGrid.Col = 1
For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.Col = 2
    Select Case True
        Case InStr(1, vGrid.Text, "CAPACIDAD")
            vGrid.Col = 1
            clsMensajes.COD_CAPACIDAD = vGrid.Text
        
        Case InStr(1, vGrid.Text, "ENDEUDAMIENTO")
            vGrid.Col = 1
            clsMensajes.COD_ENDEUDAMIENTO = vGrid.Text
        
        Case InStr(1, vGrid.Text, "HISTORIAL")
            vGrid.Col = 1
            clsMensajes.COD_HISTORIAL = vGrid.Text
        
        Case InStr(1, vGrid.Text, "MOROSIDAD")
            vGrid.Col = 1
            clsMensajes.COD_MORA = vGrid.Text
        
        Case InStr(1, vGrid.Text, "GARANTIA")
            vGrid.Col = 1
            clsMensajes.COD_GARANTIA = vGrid.Text
        
    End Select
Next i

End Sub

Public Function fxAgregaColleccionBorrar(ByVal cod_preanalisis As String, ByVal cod_preanalisis_ref As String) As String
On Error GoTo error
Dim Vcoleccion As New Collection
With Vcoleccion
    .Add fxFormatearValor(cod_preanalisis, caracter)
    .Add fxFormatearValor(cod_preanalisis_ref, caracter)

End With
fxAgregaColleccionBorrar = fxFormatearValuesCollection(Vcoleccion)

Exit Function
error:
    MsgBox fxSys_Error_Handler(Err.Description)
End Function
Private Sub sbBorrar()

On Error GoTo vError

If Not fxValidaDatosBorrar Then Exit Sub

clsEntidad.tablaName = "spCRDPreaPREANALISIS"

If m_ventanaEnModo = ModificarRegistro Then
    If (MsgBox("¿ Desea borrar la información seleccionada?", vbQuestion + vbYesNo, gMsgTitulo) = vbYes) Then
        Call clsEntidad.fxRemover(fxAgregaColleccionBorrar(clsMensajes.cod_preanalisis, clsMensajes.cod_preanalisis_ref))
            MsgBox "La información fue borrada correctamente.", vbInformation, gMsgTitulo
            cboSubExpediente.ListIndex = 0
            Call tlb_ButtonClick(tlb.Buttons("deshacer"))
        End If
Else
    MsgBox "La información no se encuentra almacenada.", vbInformation, gMsgTitulo
End If


    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)

End Sub



Private Function fxGuardar() As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset
Dim vMensaje As String


On Error GoTo vError

vMensaje = ""

fxGuardar = fxValidaDatos(m_PreviousTab)

If Not fxGuardar Then Exit Function


Screen.MousePointer = vbHourglass

clsEntidad.tablaName = "spCRDPreaPREANALISIS"
  
'Nuevo
If Len(vMensaje) = 0 Then
  strSQL = "exec spCrdFormaliza_Valida_Rangos '" & txtCedula.Text & "','" & txtLinea.Text & "'," _
         & CCur(txtMonto.Text) & "," & CCur(txtTasa.Text) & "," & CInt(txtPlazo.Text) _
         & ",'" & SIFGlobal.fxCodText(cboDestino.Text) & "','" & fxGarantia(cboGarantia.Text) _
         & "',0"
  Call OpenRecordSet(rsX, strSQL)
  If Len(rsX!Mensaje) > 0 Then
      vMensaje = vMensaje & vbCrLf & rsX!Mensaje
  End If
  rsX.Close
End If

'Fix
If Not IsNumeric(txtPolizaDesempleo.Text) Then
    Call sbCalculaPolizaDesempleo
End If


txtCumplimientoNotas.Text = vMensaje
clsMensajes.CUMPLIMIENTO_NOTAS = txtCumplimientoNotas.Text
clsMensajes.APL_POLIZA_DESEMPLEO = chkPolizaDesempleo.Value
clsMensajes.MONTO_POLIZA_DESEMPLEO = txtPolizaDesempleo.Text
  
Select Case True
  Case m_ventanaEnModo = NuevoRegistro
    
    If clsEntidad.fxAgregar(clsMensajes.fxConcatenaColleccion) Then
             
        Call sbTraerMaxExpediente
        Call txtExpediente_LostFocus
        Call SbAccionVentana(ModificarRegistro)
        
        If m_MuestraMensaje = True Then
            MsgBox "La información fue registrada correctamente.", vbInformation, gMsgTitulo
        End If


        'Inicializa datos vinculados
        If gPreAnalisis.Expediente = "" Then
            MsgBox "El número de expediente no fue cargado en las variables globales.", vbInformation, gMsgTitulo
        Else
        
            'Refundiciones (Lista)
            glogon.strSQL = "spCRDPreaRefundiciones " & fxFormatearValor(gPreAnalisis.Expediente, caracter) & "," & "'I'"
            If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
                MsgBox "Ocurrió un error al inicializar refundiciones.", vbInformation, gMsgTitulo
            End If
            
            'Fianzas (Lista)
            glogon.strSQL = "spCRDPreaFianzas " & fxFormatearValor(gPreAnalisis.Expediente, caracter) & "," & "'I'"
            If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
                MsgBox "Ocurrió un error al inicializar fianzas.", vbInformation, gMsgTitulo
            End If
        End If 'Else
    
    End If
    
 Case m_ventanaEnModo = ModificarRegistro

Call clsEntidad.fxModificar(clsMensajes.fxConcatenaColleccion)
    Call SbAccionVentana(ModificarRegistro)
    Call txtExpediente_LostFocus
    
    If m_MuestraMensaje = True Then
        MsgBox "La información fue actualizada correctamente.", vbInformation, gMsgTitulo
    End If
    
    
End Select

'FIX TEMPORAL DE COLUMNAS NUEVAS
With glogon
   If Not IsNumeric(txtPolizaDesempleo.Text) Then
    txtPolizaDesempleo.Text = "0"
   End If

    .strSQL = "UPDATE CRD_PREA_PREANALISIS SET CUMPLIMIENTO_NOTAS = '" & txtCumplimientoNotas.Text _
            & "', MONTO_POLIZA_DESEMPLEO = " & CCur(txtPolizaDesempleo.Text) _
            & " , APL_POLIZA_DESEMPLEO = " & chkPolizaDesempleo.Value _
            & " where cod_Preanalisis = '" & txtExpediente.Text & "'"
   Call ConectionExecute(.strSQL)
End With


m_CargoSalario = False
m_MuestraMensaje = False


Screen.MousePointer = vbDefault

Exit Function

vError:
    Screen.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbMostraVentanBusqueda()
On Error GoTo vError
tcMain.Item(0).Selected = True
frmPreaConsultaExpeditentes.Show vbModal
If frmPreaConsultaExpeditentes.m_Expediente <> "" Then
    txtExpediente.SetFocus
    txtExpediente.Text = frmPreaConsultaExpeditentes.m_Expediente
    Call txtExpediente_LostFocus
    txtCedula.SetFocus
    
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description)
    
End Sub


Private Sub btnComiteCambio_Click()
Dim strSQL As String

If txtExpediente.Text = "" Then Exit Sub

On Error GoTo vError


strSQL = "exec spCrd_Prea_Expediente_Comite_Cambio '" & txtExpediente.Text & "','" & glogon.Usuario & "'," & cboComite.ItemData(cboComite.ListIndex)
Call ConectionExecute(strSQL)

If glogon.error Then
    MsgBox "No fue posible realizar el cambio de estado, verifique!", vbExclamation
End If

'Actualiza Expediente
Call txtExpediente_LostFocus

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnCopiar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If txtExpediente.Text = "" Then Exit Sub

On Error GoTo vError


strSQL = "exec spCrd_Prea_Expediente_Copia '" & txtExpediente.Text & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If glogon.error Then
    MsgBox "No fue posible realizar la copia del Expediente, verifique!", vbExclamation
    Exit Sub
End If

txtExpediente.Text = rs!Expediente

rs.Close

'Actualiza Expediente
Call txtExpediente_LostFocus

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnDetalle_Click(Index As Integer)
Dim vProceso As String

gPreAnalisis.Expediente = txtExpediente.Text
GLOBALES.gTag = txtExpediente.Text
GLOBALES.gTag2 = dtpCorte.Value


  Select Case Index
    Case 0 'Rebajo de Extras
          m_curValor_Anterior = txtRebajoExtras
          
          frmPreaSubExtras.Show vbModal
          txtRebajoExtras.Text = GLOBALES.gTag
          
          If txtRebajoExtras <> m_curValor_Anterior Then
                Call sbEstructuraActualiza(1, False)
                m_CambioCalculo = True
          End If
          
    Case 1 'Cargas Sociales
         
         m_curValor_Anterior = txtTotal_Cargas_CCSS
    
         fraDCargas.Visible = True
         Call sbPosicionFrameCargas
         Call SbCalculaCargasCCSS
         
        
         
    Case 2 'Deducciones
    
          m_curValor_Anterior = txtDeducciones
    
          frmPreaSubDeducciones.Show vbModal
          
          txtDeducciones.ToolTipText = GLOBALES.gTag
          txtDeducciones.Text = GLOBALES.gTag2
         
          If txtDeducciones <> m_curValor_Anterior Then
                Call sbEstructuraActualiza(3, False)
                m_CambioCalculo = True
          End If
          
          txtCrdTransitoCancelados.SetFocus

                     
    Case 3 'Refundiciones
    
          m_curValor_Anterior = txtRefundiciones
          
          frmPreaSubRefundicionesNew.Show vbModal
          
          txtRefundiciones.ToolTipText = GLOBALES.gTag
          txtRefundiciones.Text = GLOBALES.gTag2
          
            
          If txtRefundiciones <> m_curValor_Anterior Then
                Call sbEstructuraActualiza(4, True)
                m_CambioCalculo = True
          End If
            
          txtRefundiciones.SetFocus
            
    Case 4 'Desembolsos
    
          m_curValor_Anterior = txtDesembolsos
    
          frmPreaSubDesembolsos.Show vbModal
          
          txtDesembolsos.ToolTipText = Format(GLOBALES.gTag, "Standard")
          txtDesembolsos.Text = Format(GLOBALES.gTag2, "Standard")
                
          If txtDesembolsos <> m_curValor_Anterior Then
              Call sbEstructuraActualiza(4, True)
              m_CambioCalculo = True
          End If
         
          txtDesembolsos.SetFocus
           
    Case 5 'Fianzas
         'Ojo solo si la garantia es Fiduciaria
         
          m_curValor_Anterior = txtFianzas.Text
    
          frmPreaSubFianzas.Show vbModal
          
          txtFianzas.Text = GLOBALES.gTag
          txtFianzas.ToolTipText = "Saldo..: " & GLOBALES.gTag2
          
          If CCur(txtFianzas) <> m_curValor_Anterior Then
                Call sbEstructuraActualiza(5, False)
                m_CambioCalculo = True
          End If
              
          txtFianzas.SetFocus
          
    Case 6 'Creditos Cancelados en Tramite
        'Tipo = C
        
        m_curValor_Anterior = txtCrdTransitoCancelados.Text
        
        vProceso = DatePart("yyyy", dtpCorte.Value)
        If DatePart("m", dtpCorte.Value) < 10 Then
            vProceso = vProceso & "0" & DatePart("m", dtpCorte.Value)
        Else
            vProceso = vProceso & DatePart("m", dtpCorte.Value)
        End If
        
        'Inicializa Créditos en transito
        glogon.strSQL = "spCRDPreaCreditosTransito  " & fxFormatearValor(gPreAnalisis.Expediente, caracter) & "," & "'I'" & "," & vProceso
        If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
            MsgBox "Ocurrió un error al inicializar créditos en transito.", vbInformation, gMsgTitulo
        End If

        glogon.strSQL = "exec spCRDPreaCreditosTransito '" & txtExpediente & "','D',0"
        gPreAnalisis.Tag1 = "C"
        
        frmPreaSubCreditosEnTramite.Show vbModal
        
        txtCrdTransitoCancelados.Text = GLOBALES.gTag
        
        If txtCrdTransitoCancelados <> m_curValor_Anterior Then
            Call sbEstructuraActualiza(3, False)
            m_CambioCalculo = True
        End If
        
        txtCrdTransitoCancelados.SetFocus
        
    Case 7 'Creditos x Cobrar
        'Tipo = A

        m_curValor_Anterior = txtCrdTransitoXCobrar.Text

        vProceso = DatePart("yyyy", dtpCorte.Value)
        If DatePart("m", dtpCorte.Value) < 10 Then
            vProceso = vProceso & "0" & DatePart("m", dtpCorte.Value)
        Else
            vProceso = vProceso & DatePart("m", dtpCorte.Value)
        End If
        
        'Inicializa Créditos en transito
        glogon.strSQL = "spCRDPreaCreditosTransito  " & fxFormatearValor(gPreAnalisis.Expediente, caracter) & "," & "'I'" & "," & vProceso
        If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
            MsgBox "Ocurrió un error al inicializar créditos en transisto.", vbInformation, gMsgTitulo
        End If
        
        gPreAnalisis.Tag1 = "A"
        frmPreaSubCreditosEnTramite.Show vbModal
        txtCrdTransitoXCobrar.Text = GLOBALES.gTag
    
        If txtCrdTransitoXCobrar <> m_curValor_Anterior Then
            Call sbEstructuraActualiza(3, False)
            m_CambioCalculo = True
        End If
        
        txtCrdTransitoXCobrar.SetFocus
  End Select

tcMain.Item(1).Selected = True

'Aplica todas las Formulas
Call sbAplicarFormulas(eFormulas.eAplicarTodas)

End Sub

Private Sub btnGestion_Click(Index As Integer)

If Trim(txtExpediente.Text) = "" Then Exit Sub

'Verifica y Guarda en caso de cambios


'TODO: Call ssTab_Click(ssTab.Tab)
Call sbTabChange(tcMain.SelectedItem)

Select Case Index

    Case 0 'Causas
        If Len(txtExpediente) > 0 Then
            frmPreaSeguimientoCausas.mCod_linea = Trim(txtLinea)
            frmPreaSeguimientoCausas.Show
       End If
       
    Case 1 'Tags
        If Len(txtExpediente) > 0 Then
            frmPrea_SeguimientoEtiquetas.mId_Solicitud = Trim(txtAsignado)
            frmPrea_SeguimientoEtiquetas.Show
        End If


    Case 2 'gestion
            If fxExistenFiadores Then
                
                frmPreaEstadoPreanalisis.m_estadoPreanalisis = m_estadoPreanalisis
                
                gPreAnalisis.Expediente = txtExpediente.Text
                frmPreaEstadoPreanalisis.Show vbModal
                m_estadoPreanalisis = frmPreaEstadoPreanalisis.m_estadoPreanalisis
                
                tcMain.Item(0).Selected = True
                
                Select Case m_estadoPreanalisis
                Case "P"
                    lblEstado.Caption = "Pendiente"
                Case "R"
                    lblEstado.Caption = "Recibido"
                Case "A"
                    lblEstado.Caption = "Aprobado"
                Case "D"
                    lblEstado.Caption = "Denegado"
                End Select

            Else
            
                tcMain.Item(0).Selected = True
            End If
                      
    Case 3 'Solicitud
    
            If fxExistenFiadores Then
                gPreAnalisis.Expediente = txtExpediente.Text
                gPreAnalisis.Tag1 = txtCedula
                frmPreaSubCredito.Show vbModal
                tcMain.Item(0).Selected = True
                
                txtAsignado = Trim(frmPreaSubCredito.m_Id_Solicitud)
            Else
                tcMain.Item(0).Selected = True
            End If

End Select

Me.MousePointer = vbDefault

End Sub

Private Sub btnSolicitado_Click()
Dim strSQL As String

If txtExpediente.Text = "" Then Exit Sub

On Error GoTo vError


strSQL = "exec spCrd_Prea_Estado_Solicitado '" & txtExpediente.Text & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

If glogon.error Then
    MsgBox "No fue posible realizar el cambio de estado, verifique!", vbExclamation
End If

'Actualiza Expediente
Call txtExpediente_LostFocus

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboCantidadFiadores_Change()
    m_CambioDatos = True
End Sub

Private Sub cboCantidadFiadores_Click()
m_CambioDatos = True
If cboSubExpediente.ListCount = 1 Then Exit Sub
If m_DesplegoMensaje Then Exit Sub

m_DesplegoMensaje = True
If m_FiadoresRegistrador > cboCantidadFiadores.Text Then
    MsgBox "No es posible disminuir la cantidad de sub expedientes."
    cboCantidadFiadores.Text = m_FiadoresRegistrador
End If
m_DesplegoMensaje = False
End Sub
     
Private Sub cboComite_Change()
        m_CambioDatos = True
End Sub

Private Sub cboComite_Click()
    m_CambioDatos = True
End Sub

Private Sub cboDestino_Click()
   m_CambioDatos = True
   Call sbAplicaPrimeraCta
End Sub
Private Sub sbAplicaPrimeraCta()

On Error GoTo vError

chkPrimerCuota.Value = 0

If (cboDestino.ListCount = 0) Or (SIFGlobal.fxCodText(cboDestino.Text) = "") Then Exit Sub
If Len(txtDesLineaCredito.Text) = 0 Then Exit Sub

clsEntidad.tablaName = "spCRDPreaDestinos"
If clsEntidad.fxTraerFiltrado("AplicaPrimCta", "'" & Trim(txtLinea.Text) & "','" & SIFGlobal.fxCodText(cboDestino.Text) & "'") Then
    chkPrimerCuota.Value = glogon.Recordset.Fields!PRIMER_CUOTA
Else
    chkPrimerCuota.Value = 0
End If

 
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Private Sub cboFondo_Click()
    Dim strSQL As String, rs As New ADODB.Recordset
    
    If vPasoCarga Then Exit Sub
    If cboFondo.ListCount <= 0 Then Exit Sub
    If cboFondo.Text = "" Then Exit Sub
    
    If fxGarantia(cboGarantia.Text) <> "Y" Then Exit Sub
    
    strSQL = "exec spCRDGarantiaFNDCalculo '" & Trim(txtCedula.Text) & "','" & SIFGlobal.fxCodText(cboFondo.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    txtMonto.Text = Format(rs!Disponible, "Standard")
    If rs!AplicaTasa = 1 Then
        txtTasa.Text = rs!TASA
    End If
    
    If rs!AplicaPlazo = 1 Then
        txtPlazo.Text = rs!Plazo
    End If
    
    LblTasa.Tag = rs!AplicaTasa
    LblPlazo.Tag = rs!AplicaPlazo
    
    rs.Close
    
    m_CambioDatos = True

End Sub

Private Sub cboGarantia_Change()
    m_CambioDatos = True
End Sub

Private Sub cboGarantia_Click()
Dim strSQL As String, rs As New ADODB.Recordset

m_CambioDatos = True

chkPolizaVida.Value = vbUnchecked
chkPolizaVida.Enabled = True
cboFondo.Enabled = False

cboCantidadFiadores.Enabled = False

Dim pGarantia As String, pGarantiaForm As String

pGarantia = fxGarantia(cboGarantia.Text)

strSQL = "select FORMULARIO  From CRD_GARANTIA_TIPOS" _
       & " where garantia = '" & pGarantia & "'"
Call OpenRecordSet(rs, strSQL)
 pGarantiaForm = Trim(rs!Formulario)
rs.Close


Select Case pGarantiaForm
    
    
    Case "F01" 'Sobre Ahorros
        strSQL = "select dbo.fxCrdGarantiaPatMnt('" & txtCedula.Text & "','A', 'M') as 'Monto'"
        Call OpenRecordSet(rs, strSQL)
          txtMonto.Text = Format(rs!Monto, "Standard")
        rs.Close
    
    Case "F02" 'Fiduciaria
    
        cboCantidadFiadores.Enabled = True
    
    Case "F03" 'Hipotecaria
    
        cboCantidadFiadores.Enabled = True
        
        chkPolizaVida.Value = vbChecked
        Call chkPolizaVida_Click
        
        chkPolizaVida.Enabled = False
    
    Case "F05" 'Fondos de Ahorros
            If vPasoCarga Then Exit Sub
            If cboGarantia.ListCount <= 0 Then Exit Sub
            
             cboFondo.Enabled = True
             Call cboFondo_Click
            Exit Sub
    
    Case "F06" 'Adelanto de Salario
        strSQL = "select dbo.fxCrdDisponibleAdelantoSalario_Estudio('" & txtCedula.Text & "', 'M') as 'Monto'"
        Call OpenRecordSet(rs, strSQL)
          txtMonto.Text = Format(rs!Monto, "Standard")
        rs.Close
    

End Select


'Corrección
Select Case pGarantia
    Case "F", "H"
      'Nada
    Case Else
        If Val(cboCantidadFiadores.Text) > 0 Then
            m_FiadoresRegistrador = 0
            cboCantidadFiadores.Text = 0
            cboCantidadFiadores.Enabled = False
        End If
End Select



m_curValor_Anterior = 0
If IsNumeric(txtMonto.Text) Then
    If CCur(txtMonto.Text) > 0 Then
      Call sbCalcularCuota("txtMonto")
    End If
End If

End Sub

'---------------------Registro de cliente
Private Sub sbCrearEncabezado()
On Error GoTo error
ltvSalarios.ListItems.Clear
ltvSalarios.ColumnHeaders.Clear
ltvSalarios.ColumnHeaders.Add , , "Salario", 2000 '0
ltvSalarios.ColumnHeaders.Add , , "Fecha", 2000 '1


Exit Sub
error:
    cMensaje.deError ("Ocurrió un error  al establecer la columnas de lista de empresas. Error:" & Err.Description)
End Sub
Private Function fxBorrarExtras() As Boolean

On Error GoTo vError
Me.MousePointer = vbHourglass
 
fxBorrarExtras = False

'If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
'    GoTo salir
'End If

clsEntidad.tablaName = "spCRDPreaExtrasXPreAnalisis"
If clsEntidad.fxRemover("'" & gPreAnalisis.Expediente & "'") Then
    fxBorrarExtras = True
End If


    Me.MousePointer = vbDefault
    Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
    
    
End Function



Private Sub cboSalario_Click()
Dim Item As String
Dim Codigo As String
Dim IndiceTipoSalario As String
Dim sql As String

On Error GoTo vError

    If vPaso Then Exit Sub

    'Solo estos codigos permiten a, b , c y f seleccionar el salario
    
    Me.MousePointer = vbHourglass

    Codigo = SIFGlobal.fxCodText(cboSalario.Text)
    Item = Left(Right(cboSalario.Text, 2), 1)
    
'    Call sbHabilitaFechaColilla(Trim(Item))

    IndiceTipoSalario = Right(cboSalario.Text, 3)
    
    'sbCRDPreaSalario(@Expediente varchar(15), @TipoSalario varchar(2) = '')
    sql = "spCRDPreaSalario " & fxFormatearValor(gPreAnalisis.Expediente, caracter) & ","
    sql = sql & fxFormatearValor(Trim(Codigo), caracter)
    
    If m_CargoSalario Then
        m_CargoSalario = False
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        Call sbActCtlConstExternos(True, IndiceTipoSalario)

        Select Case Trim(Item)
        Case "a", "b", "c", "f"
            
            If clsEntidad.fxEjecutaSQL(sql) Then
                If glogon.Recordset.RecordCount = 1 Then
                    txtSalarioDevengado.Text = Format(glogon.Recordset!Salario, "Standard")
'                    dtpCorte = CDate(glogon.Recordset!Fecha)
                Else
                    m_SoloVerSalarios = False
                    Call sbPosicionFrameSalarios
                    FrameSalarios.Caption = cboSalario.Text
                    Call sbLlenarListSalario
                End If
            End If
            
        Case "g"
        
            Call sbActCtlConstExternos(False, IndiceTipoSalario)
        
        Case Else
            
            txtSalarioDevengado.Text = 0
    
        End Select
    
    End If
    Call SbBloquearTxtSalario(Trim(Codigo))

    m_CambioCalculo = True
    
    Call sbEstructuraActualiza(1, False)
    
    If Item <> "e" Then
     '   Call fxBorrarExtras
    End If
    
    
    
    Me.MousePointer = vbDefault
Exit Sub
vError:
    Me.MousePointer = vbDefault
    cMensaje.deError ("Ocurrió un error . Error:" & Err.Description)

End Sub


Private Sub sbHabilitaFechaColilla(ByVal Tipo As String)

    If Tipo = "f" Or Tipo = "a" Then
        dtpCorte.Enabled = False
    Else
        dtpCorte.Enabled = True
    End If

End Sub

Private Sub sbActCtlConstExternos(ByVal Activar As Boolean, ByVal IndiceTipoSalario As String)
Dim vkey As String
On Error GoTo vError
Me.MousePointer = vbHourglass
If Activar Then
    txtSalarioLiquido.BackColor = &HE0E0E0
    txtTotal_Cargas_CCSS.BackColor = &HFFFFFF
    txtDeducciones.BackColor = &HFFFFFF
Else
    txtSalarioLiquido.BackColor = &HFFFFFF
    txtSalarioLiquido.Locked = False
    txtTotal_Cargas_CCSS.BackColor = &HE0E0E0
    txtDeducciones.BackColor = &HE0E0E0
End If

 If IndiceTipoSalario = "(g)" Then 'codigo que Corresponde al tipo de salario
    
    btnDetalle.Item(1).Visible = False
    btnDetalle.Item(2).Visible = False
    
    txtTotal_Cargas_CCSS.Text = Format(0, "Standard")
    txtDeducciones.Text = Format(0, "Standard")
Else
    btnDetalle.Item(1).Visible = True
    btnDetalle.Item(2).Visible = True
End If


Me.MousePointer = vbDefault
Exit Sub

vError:
cMensaje.deError ("Ocurrió un error  activando o deshablitando controles. Error:" & Err.Description)


End Sub

Private Sub sbLlenarListSalario()
Dim vkey As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbCrearEncabezado

FrameSalarios.Visible = True

With glogon.Recordset
  While Not .EOF
  vkey = "(IN)" & .Fields!Salario & "(SA)" & .Fields!fecha & "(id)"
        Set litem = ltvSalarios.ListItems.Add(, vkey, Format(.Fields("Salario"), "Standard"))
        litem.SubItems(1) = Format(.Fields("Fecha"), "dd-mm-yyyy")
        .MoveNext
    Wend
  End With
  

Me.MousePointer = vbDefault

Exit Sub


vError:
cMensaje.deError ("Ocurrió un error al mostrar la lista de salarios. Error:" & Err.Description)

End Sub





Private Sub cboSexo_Change()
    m_CambioDatos = True
End Sub

Private Sub cboSexo_Click()
    Call sbCalcularPlazoMaximo
    m_CambioDatos = True
End Sub

Private Sub cboSubExpediente_Click()
   
    If m_Cargando Then Exit Sub
    If txtExpediente.Text = cboSubExpediente.Text Then
        Exit Sub
    End If

    Me.MousePointer = vbHourglass

    m_valorComboExp = cboSubExpediente.Text
        
    '' Proceso pregunta si desea guardar los datos.
    If cboSubExpediente.Text = "Nuevo SubExpediente" Or m_valorComboExp = "Nuevo Expediente" Then
        If Trim(txtExpediente.Text) <> "" Then
            If (m_CambioDatos = True) Or (m_CambioCalculo = True) Or (m_CambioObservaciones = True) Then
                If (MsgBox("NO GUARDO los cambios que hizo en el expediente, ¿Desea continuar SIN GUARDAR los cambios?.", vbQuestion + vbYesNo + vbDefaultButton2, gMsgTitulo) = vbNo) Then
                    cboSubExpediente.Text = Trim(txtExpediente.Text)
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    End If
    
    DoEvents

    Call sbcboSubExpediente_Validate

    If cboSubExpediente.Text = "Nuevo SubExpediente" Then
        vCodExpediente = txtExpediente.Text
        Me.MousePointer = vbDefault
        If fxValidaNumFiadoresRegistrados(True) = False Then Exit Sub
    End If

    DoEvents

    If Right(cboSalario.Text, 3) = "(g)" Then
        Call sbActCtlConstExternos(False, "(g)")
    Else
        Call sbActCtlConstExternos(True, "(d)")
    End If

    If m_valorComboExp = "Nuevo SubExpediente" Or m_valorComboExp = "Nuevo Expediente" Then
        txtCedula.SetFocus
    End If

    sbActivarMontoGirar
    
    
    Me.MousePointer = vbDefault

End Sub

Private Sub sbActivarMontoGirar()
    If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
        txtMontoGirar.Visible = False
        lblMontoGirar.Item(34).Visible = False
    Else
        txtMontoGirar.Visible = True
        lblMontoGirar.Item(34).Visible = True
    End If
End Sub

Private Sub sbcboSubExpediente_Validate()
Dim vControl As Control
txtExpediente.Locked = True

If m_valorComboExp = "Nuevo SubExpediente" Then
'If fxSelectItemSubExpediente(cboSubExpediente.ItemData(cboSubExpediente.ListIndex)) = "S" Then
   ' m_expediente = txtExpediente.Text
    For Each vControl In Me
'        MsgBox TypeName(vControl) & "  :" & vControl.Name
        Select Case TypeName(vControl)
        Case "TextBox", "FlatEdit"
            If Not ((vControl.Name = "txtExpediente") _
                Or (vControl.Name = "txtLinea") _
                Or (vControl.Name = "txtDesLineaCredito") _
                Or (vControl.Name = "txtPolizaVida") _
                Or (vControl.Name = "txtPolizaIncendio") _
                Or (vControl.Name = "txtPolizaDesempleo") _
                Or (vControl.Name = "txtMonto") _
                Or (vControl.Name = "txtPlazo") _
                Or (vControl.Name = "txtCuota") _
                Or (vControl.Name = "txtCompromiso") _
                Or (vControl.Name = "txtTasa")) _
                Then
                
                    vControl.Text = ""
            End If
            txtExpediente.Locked = False
            
        Case "CheckBox"
            If Not ((vControl.Name = "chkPolizaVida") _
                Or (vControl.Name = "chkPolizaIncendio") _
                Or (vControl.Name = "chkPolizaDesempleo")) _
                Or (vControl.Name = "chkPrimerCuota") _
            Then
                'vControl.Value = 0
                If vControl.Name <> "chkCargaFrap" Then
                    vControl.Enabled = False
                End If
                
            End If
        End Select
    Next vControl
    
    chkCargaAsociacion.Enabled = True
    Call SbAccionVentana(NuevoRegistro)
    Call sbBloquearControles(Me, SubExpediente)
    
    chkCargaAsociacion.Value = vbUnchecked
    chkCargaFrap.Value = vbUnchecked

ElseIf m_valorComboExp = "Nuevo Expediente" Then
    Call sbLimipiaControles(Me, True)
    
    dtpCorte.Value = fxFechaServidor
    
    If cboSubExpediente.ListCount > 1 Then
        Call sbInicializaComboExpediente
    End If
    
    Call SbAccionVentana(NuevoRegistro)
    Call sbBloquearControles(Me, Expediente)
    txtExpediente.Locked = False

ElseIf m_valorComboExp = "Nuevo Expediente" Or m_valorComboExp = "Nuevo SubExpediente" Then
    Else
        txtExpediente.Text = m_valorComboExp
        m_CargoCombo = True
        Call txtExpediente_LostFocus
End If



End Sub

Private Sub sbCalcularCompromiso()

    txtCompromiso.Text = ""
    If Not IsNumeric(txtCuota.Text) Then
        txtCuota.Text = 0
    End If
    
    If txtPolizaVida = "" Then
        txtPolizaVida = 0
    End If
    
    If txtPolizaIncendio = "" Then
        txtPolizaIncendio = 0
    End If
    
    If txtPolizaDesempleo = "" Then
        txtPolizaDesempleo = 0
    End If
    
    
    If (Len(txtCuota.Text) > 0 And Len(txtPolizaVida.Text) > 0 And Len(txtPolizaIncendio.Text) > 0 And Len(txtPolizaDesempleo.Text) > 0) Then
        txtCompromiso.Text = Format((CDbl(txtCuota.Text) _
                + CDbl(txtPolizaVida.Text) _
                + CDbl(txtPolizaIncendio.Text) _
                + CDbl(txtPolizaDesempleo.Text) _
                ), "Standard")
    End If


End Sub

Private Sub sbCalcularCuota(ByVal Control As String)
Dim mBono As Double, mPlazo As Integer

On Error GoTo vError
 
m_CambioDatos = True

Select Case Control
    Case "txtMonto"
    
       If Val(txtMonto.Text) > 0 Then
       
        If fxGarantia(cboGarantia.Text) = "Y" And cboFondo.ListCount > 0 Then
            If LblPlazo.Tag = "0" Then txtPlazo.Text = fxCatalogoRango(Trim(txtLinea.Text), Format(txtMonto.Text, "Standard"), "P", SIFGlobal.fxCodText(cboDestino.Text), fxGarantia(cboGarantia.Text))
            If LblTasa.Tag = "0" Then txtTasa.Text = fxCatalogoRango(txtLinea.Text, txtMonto.Text, "I", SIFGlobal.fxCodText(cboDestino.Text), fxGarantia(cboGarantia.Text))
        Else
            If Len(txtDesLineaCredito.Text) > 0 Then
                txtPlazo.Text = fxCatalogoRango(txtLinea.Text, CDbl(txtMonto.Text), "P", SIFGlobal.fxCodText(cboDestino.Text), fxGarantia(cboGarantia.Text))
                txtTasa.Text = fxCatalogoRango(txtLinea.Text, CDbl(txtMonto.Text), "I", SIFGlobal.fxCodText(cboDestino.Text), fxGarantia(cboGarantia.Text))
            End If
        End If
        
        ''Modifica tasa para aplicar bonos por membresia
        If clsMensajes.Estado = "R" Or clsMensajes.Estado = "P" Then
            mBono = fxBonoMembresia(Trim(txtCedula.Text), txtLinea.Text, fxGarantia(cboGarantia.Text))
            mPlazo = fxBonoPlazoMembresia(Trim(txtCedula.Text), fxGarantia(cboGarantia.Text))
            If mBono > 0 Then
                txtTasa.Text = CDbl(txtTasa.Text) - mBono
                clsMensajes.TASA_PTS_BONO = mBono
            Else
                txtTasa.ToolTipText = Empty
                txtTasa.Tag = 0
                clsMensajes.TASA_PTS_BONO = 0
            End If
        
            If mPlazo > 0 Then
               txtPlazo.Text = mPlazo
            End If
        
        End If
        
        If clsMensajes.TASA_PTS_BONO > 0 Then
            txtTasa.ToolTipText = "Bono por Membresia de " & clsMensajes.TASA_PTS_BONO
        Else
            txtTasa.ToolTipText = Empty
        End If
        
        
        If (Val(txtPlazo.Text) > 0 And Val(txtTasa.Text) >= 0) And IsNumeric(txtMonto.Text) Then
                txtCuota.Text = Format(fxCalcula_Cuota(CDbl(txtMonto.Text), Val(txtPlazo.Text), CDbl(txtTasa.Text), mFrecuenciaPago), "Standard")
            Else
                txtCuota.Text = 0
            End If
        Else
            txtPlazo.Text = 0
            txtTasa.Text = 0
            txtCuota.Text = 0
       End If
       
    Case "txtPlazo", "txtTasa"
   
        If (Val(txtMonto.Text) > 0 And Val(txtPlazo.Text) > 0) Then 'And Val(txtTasa.Text) > 0) Then
            
            txtCuota.Text = Format(fxCalcula_Cuota(CDbl(txtMonto.Text), CDbl(txtPlazo.Text), CDbl(txtTasa.Text), mFrecuenciaPago), "Standard")
        Else
            txtCuota.Text = 0
        End If
        
End Select
Call sbCalcularCompromiso
Call sbEstructuraActualiza(5, False)

    Exit Sub
vError:
    MsgBox "Ocurrió un error al calcular la cuota. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub CerrarFrameSalarios_Click()

On Error GoTo vError
    If m_SoloVerSalarios = False Then
        Call sbEstructuraActualiza(1, False)
    End If
    FrameSalarios.Visible = False
    Exit Sub
vError:
    MsgBox "Ocurrió un error validar datos de Cargas sociales . " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub





Private Sub chkCargaAsociacion_Click()
m_CambioCalculo = True
End Sub

Private Sub chkCargaFrap_Click()
m_CambioCalculo = True
End Sub

Private Sub chkPolizaDesempleo_Click()
  Call sbCalculaPolizaDesempleo
  Call sbCalcularCompromiso
  m_CambioDatos = True
End Sub

Private Sub chkPolizaIncendio_Click()
  Call sbCalculaPolizaDeIncendio
  Call sbCalcularCompromiso
  m_CambioDatos = True
End Sub



Private Sub chkPolizaVida_Click()
    Call sbCalculaPolizaDeVida
    Call sbCalcularCompromiso
    m_CambioDatos = True
End Sub

Private Sub sbCalculaPolizaDeIncendio()
txtPolizaIncendio.Text = 0
If Val(txtMonto.Text) = 0 Then Exit Sub
If chkPolizaIncendio.Value = Checked Then
    glogon.strSQL = "select  dbo.fxCrd_Prea_Poliza_Incendio(" & fxFormatearValor(CDbl(txtMonto.Text), Numerico) & " )"
    If execSql(glogon.strSQL) Then
        If Trim(glogon.Recordset(0) & "") <> "" Then
            txtPolizaIncendio.Text = Format(glogon.Recordset(0), "Standard")
        End If
    End If
End If
End Sub


Private Sub sbCalculaPolizaDesempleo()
txtPolizaDesempleo.Text = 0

If Val(txtCuota.Text) = 0 Then Exit Sub

Dim pMonto As Currency

If Not IsNumeric(txtPolizaVida.Text) Then
        txtPolizaVida.Text = "0"
End If

If Not IsNumeric(txtPolizaIncendio.Text) Then
        txtPolizaIncendio.Text = "0"
End If

pMonto = CCur(txtCuota.Text) + CCur(txtPolizaVida.Text) + CCur(txtPolizaIncendio.Text)

If chkPolizaDesempleo.Value = Checked Then
    glogon.strSQL = "select  dbo.fxCrd_Prea_Poliza_Desempleo(" & fxFormatearValor(pMonto, Numerico) & " )"
    If execSql(glogon.strSQL) Then
        If Trim(glogon.Recordset(0) & "") <> "" Then
            txtPolizaDesempleo.Text = Format(glogon.Recordset(0), "Standard")
        End If
    End If
End If
End Sub

Private Sub sbCalculaPolizaDeVida()
txtPolizaVida.Text = 0
If Val(txtMonto.Text) = 0 Then Exit Sub
If chkPolizaVida.Value = Checked Then
    glogon.strSQL = "select  dbo.fxCrd_Prea_Poliza_Vida(" & fxFormatearValor(CDbl(txtMonto.Text), Numerico) & " )"
    '                         dbo.fxCRDCuotaPolizaVida(Monto,cod_linea,garantia)
    If execSql(glogon.strSQL) Then
        If Trim(glogon.Recordset(0) & "") <> "" Then
            txtPolizaVida.Text = Format(glogon.Recordset(0), "Standard")
        End If
    End If
End If
End Sub

Private Sub chkPrimerCuota_Click()
    m_CambioDatos = True
    Call sbEstructuraActualiza(1000, True)
End Sub


Public Function fxAgregaColleccion(ByVal Expediente As String, ByVal pObsAnalisista As String, ByVal pObsComite As String, ByVal pObsJuntaDirectiva As String) As String
On Error GoTo error
Dim Vcoleccion As New Collection
With Vcoleccion
    .Add fxFormatearValor(Expediente, caracter)
    .Add fxFormatearValor(pObsAnalisista, caracter)
    .Add fxFormatearValor(pObsComite, caracter)
    .Add fxFormatearValor(pObsJuntaDirectiva, caracter)
End With
fxAgregaColleccion = fxFormatearValuesCollection(Vcoleccion)

Exit Function
error:
    MsgBox fxSys_Error_Handler(Err.Description)
End Function

Private Sub sbGuardaObservaciones()
 
If Len(vObservacion(0)) + Len(vObservacion(1)) + Len(vObservacion(2)) = 0 Then Exit Sub
 
On Error GoTo vError
 


If ((clsMensajes.Estado = "A") Or (clsMensajes.Estado = "D")) Then
    MsgBox "No es posible realizar cambios en las observaciones del expediente seleccionado.", vbInformation, gMsgTitulo
    Exit Sub
End If


Me.MousePointer = vbHourglass

clsEntidad.tablaName = "spCRDPreaObservaciones"
If clsEntidad.fxModificar(fxAgregaColleccion(gPreAnalisis.Expediente, vObservacion(0), vObservacion(1), vObservacion(2))) Then
    MsgBox "La información se registro correctamente.", vbExclamation
    m_CambioObservaciones = False
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox "Ocurrió un error en el proceso de guardar observaciones. " & "err: " & Err.Description, vbCritical
End Sub

Private Sub cmdGuardaObservaciones_Click()
If m_CambioObservaciones Then
    Call sbGuardaObservaciones
End If
    
End Sub

Private Sub cmdScrollBar(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

  strSQL = "select top 5 * FROM  CRD_PREA_PREANALISIS"
  m_Expediente = ""
  
  Select Case Index
  
    Case 0
        strSQL = strSQL & " where cod_preanalisis_Ref  is null and  cod_preanalisis < '" & txtExpediente.Text & "' order by cod_preanalisis desc"
    Case 1
        strSQL = strSQL & " where cod_preanalisis_Ref is null  and cod_preanalisis > '" & txtExpediente.Text & "' order by cod_preanalisis asc"
        
  End Select
  
  Call OpenRecordSet(rs, strSQL)

  If Not rs.EOF And Not rs.BOF Then

  txtExpediente.Text = rs!cod_preanalisis
   Call txtExpediente_LostFocus
  
'  Call SbTraerDatosExpediente
    rs.Close
  End If



Me.MousePointer = vbDefault

Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbNumPagos_Update()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not m_Cargando And txtNombre.Text <> "" Then
  
    strSQL = "exec spCrd_Prea_NumPagos '" & txtCedula.Text & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"
    Call OpenRecordSet(rs, strSQL)
        m_NumPagos = rs!Num_Pagos
    rs.Close
End If

Exit Sub

vError:

End Sub


Private Sub dtpCorte_Change()
m_CambioCalculo = True
Call sbNumPagos_Update

End Sub


Private Sub dtpCorte_LostFocus()
'    tcMain.Item(1).Selected = True
'    txtSalarioDevengado.SetFocus
End Sub

Private Sub dtpFecNac_Change()
    Call sbCalcularPlazoMaximo
    m_CambioDatos = True
End Sub


Private Sub Form_Activate()
    vModulo = 3 'Modulo de Credito
 
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.ActiveControl.Name <> "txtObservaciones" Then
    If (KeyCode = vbKeyReturn) Then 'Or KeyCode = vbKeyTab) Then
'        Call sbCalcularCuota(Me.ActiveControl.Name)
    Call gsbPulsarTecla(vbKeyTab)
    
    ElseIf KeyCode = vbKeyF4 Then
        Call sbBusqueda(Me.ActiveControl.Name)
    End If
End If
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo vError
Select Case Me.ActiveControl.Name
Case "txtMonto"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtMonto.Text), KeyAscii)
Case "txtPolizaVida"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtPolizaVida.Text), KeyAscii)
Case "txtPolizaIncendio"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtPolizaIncendio.Text), KeyAscii)
Case "txtPolizaDesempleo"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtPolizaDesempleo.Text), KeyAscii)

Case "txtPlazo"
    KeyAscii = fxPermiteSoloDigitos(KeyAscii)
    txtPlazo.Text = fxValidaLargoZeroIzq(Trim$(txtPlazo.Text))
Case "txtTasa"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtTasa.Text), KeyAscii)
    txtTasa.Text = fxValidaLargoZeroIzq(Trim$(txtTasa.Text))
Case "txtSalarioDevengado"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtSalarioDevengado.Text), KeyAscii)
'Validar si esta se quedan por esta siempre bloqueadas
Case "txtRebajoExtras"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtRebajoExtras.Text), KeyAscii)
Case "txtSalarioReal"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtSalarioReal.Text), KeyAscii)
Case "txtExtrasFijas"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtExtrasFijas.Text), KeyAscii)
Case "txtDevengadoMes"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtDevengadoMes.Text), KeyAscii)
Case "txtPorcSobreSalario"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtPorcSobreSalario.Text), KeyAscii)
Case "txtDeducciones"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtDeducciones.Text), KeyAscii)
Case "txtCrdTransitoCancelados"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtCrdTransitoCancelados.Text), KeyAscii)
Case "txtCrdTransitoXCobrar"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtCrdTransitoXCobrar.Text), KeyAscii)
Case "txtSalarioLiquido"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtSalarioLiquido.Text), KeyAscii)
Case "txtRefundiciones"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtRefundiciones.Text), KeyAscii)
Case "txtDesembolsos"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtDesembolsos.Text), KeyAscii)
Case "txtTotalLiquido"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtTotalLiquido.Text), KeyAscii)
Case "txtFianzas"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtFianzas.Text), KeyAscii)
Case "txtLiquidezSinFianza"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtLiquidezSinFianza.Text), KeyAscii)
Case "txtLiquidezPorcSinFianza"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtLiquidezPorcSinFianza.Text), KeyAscii)
Case "txtLiquidezConFianza"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtLiquidezConFianza.Text), KeyAscii)
Case "txtLiquidezPorcConFianza"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtLiquidezPorcConFianza.Text), KeyAscii)
Case "txtMontoGirar"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtMontoGirar.Text), KeyAscii)
'Fin/Validar si esta se quedan por esta siempre bloqueadas

Case "txtCargaCCSS"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(lblCargaCCSS.Caption), KeyAscii)
Case "txtCargaImpSalario"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(lblCargaImpSalario.Caption), KeyAscii)
Case "txtCargaAsociacion"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(lblCargaAsociacion.Caption), KeyAscii)
Case "txtCargaFrap"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(lblCargaFrap.Caption), KeyAscii)
End Select


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar la información de los formatos. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub
  


'Eventos del forms
Private Sub Form_Load()
 
vModulo = 3 'Modulo de Credito
 
m_NumPagos = 2
 
txtSalarioDevengado.BackColor = RGB(187, 215, 247)
txtExtrasFijas.BackColor = RGB(187, 215, 247)

txtPlMax.BackColor = RGB(187, 215, 247)
txtAsignado.BackColor = RGB(187, 215, 247)
txtClasificacion.BackColor = RGB(187, 215, 247)
txtCuota.BackColor = RGB(187, 215, 247)
txtCompromiso.BackColor = RGB(187, 215, 247)

txtExpediente.Height = cboSubExpediente.Height

 mFrecuenciaPago = "M"

Call sbInicializaGlobales
lblPorcentajeSalario.Caption = str(GlobalPorcLiquidezLibre) & " % Sobre Salario "
'txtCedula.SetFocus
 
'Inicializa Barra
Call sbToolBarIconos(tlb)
Call sbToolBar(tlb, "nuevo")


cboCantidadFiadores.Clear
cboCantidadFiadores.AddItem "0"
cboCantidadFiadores.AddItem "1"
cboCantidadFiadores.AddItem "2"
cboCantidadFiadores.AddItem "3"
cboCantidadFiadores.AddItem "4"
cboCantidadFiadores.Text = "0"

'Inicializa Seguridad
Call Formularios(Me)
Call RefrescaTags(Me)

vPaso = True

Call sbCargarCombos
Call SbAccionVentana(NuevoRegistro)
Call sbInicializaComboExpediente

tcMain.Item(0).Selected = True

Call sbBloquearTab

m_CargoSalario = True
m_FECHA_CREACION = "-1"
cboCantidadFiadores.ListIndex = 0
m_CambioDatos = False
m_CambioCalculo = False
m_CambioObservaciones = False

vPaso = False

Me.MousePointer = vbDefault

End Sub


Private Sub Form_Unload(Cancel As Integer)

    sbDeseaGuardar
    
    Set clsMensajes = Nothing
    Set clsEntidad = Nothing
    Set clsNull = Nothing

End Sub


Private Sub sbDeseaGuardar()

    If (m_CambioDatos = True) Or (m_CambioCalculo = True) Or (m_CambioObservaciones = True) Then
        If (MsgBox("¿Desea guardar los cambios efectuados en el expediente seleccionado. ?", vbQuestion + vbYesNo, gMsgTitulo) = vbYes) Then
            Call fxGuardar
            Call sbGuardaObservaciones
        End If
    End If

End Sub







Private Sub imgFraCerrar_Click()
Dim CargasCCSS As Double
On Error GoTo vError
    CargasCCSS = 0
    If Val(lblCargaImpSalario.Caption) > 0 Then
        CargasCCSS = lblCargaImpSalario.Caption
    End If
    
    
   If chkCargaAsociacion.Value = Checked Then
        If IsNumeric(GlobalPorcCCSS) Then
          lblCargaCCSS.Caption = Format((GlobalPorcCCSS * txtDevengadoMes.Text) / 100, "Standard")
        End If
        CargasCCSS = CDbl(lblCargaCCSS.Caption) + CDbl(lblCargaImpSalario.Caption)
        CargasCCSS = CargasCCSS + CDbl(lblCargaAsociacion.Caption)
        txtTotal_Cargas_CCSS.Text = Format(CargasCCSS, "Standard")
        chkCargaAsociacion.Tag = "S"
   Else
        chkCargaAsociacion.Tag = "N"
        CargasCCSS = CDbl(lblCargaCCSS.Caption) + CDbl(lblCargaImpSalario.Caption)
   End If
   
   If chkCargaFrap.Value = Checked Then
    CargasCCSS = CargasCCSS + CDbl(lblCargaFrap.Caption)
    
        chkCargaFrap.Tag = "S"
    Else
        chkCargaFrap.Tag = "N"
   End If
   
  txtTotal_Cargas_CCSS.Text = Format(CargasCCSS, "Standard")
  
   If CCur(txtTotal_Cargas_CCSS.Text) <> m_curValor_Anterior Then
        Call sbEstructuraActualiza(3, False)
        m_CambioCalculo = True
'        Call fxGuardar
   End If
  
  txtPorcSobreSalario.SetFocus
  fraDCargas.Visible = False

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar datos de Cargas sociales . " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub





















Private Sub lblSalarioDevengado_DblClick(Index As Integer)
    Dim sql As String
    
    sql = "spCRDPreaSalariosLista " & fxFormatearValor(gPreAnalisis.Expediente, caracter)
        
    If clsEntidad.fxEjecutaSQL(sql) Then
        m_SoloVerSalarios = True
        Call sbPosicionFrameSalarios
        FrameSalarios.Caption = cboSalario.Text
        Call sbLlenarListSalario
    End If

        
End Sub



Private Sub ltvSalarios_DblClick()
    If Not Item_Seleccionado Is Nothing Then
        If m_SoloVerSalarios = False Then
            txtSalarioDevengado.Text = Format(CDbl(fxDeCodificaPrimaryKey(Item_Seleccionado.Key, 5, "(SA)")), "Standard")
            txtRebajoExtras.SetFocus
            Call sbEstructuraActualiza(1, False)
        End If
        FrameSalarios.Visible = False
    End If
End Sub




Private Sub ltvSalarios_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    Set Item_Seleccionado = Item
End Sub

Private Sub optCausas_Click(Index As Integer)
Call sbCargarListaCausas
End Sub

Private Sub optObservacion_Click(Index As Integer)
On Error GoTo vError

Select Case Index
 Case 0
  txtObservaciones.Text = vObservacion(0)
 Case 1
  txtObservaciones.Text = vObservacion(1)
 Case 2
  txtObservaciones.Text = vObservacion(2)
End Select

txtObservaciones.SetFocus
m_CambioObservaciones = False

vError:

End Sub







Private Sub rbActas_Click(Index As Integer)
Call sbEstudio_Comite_Resolucion_Load(txtExpediente.Text)
End Sub

Private Sub Tbl_Desplazamiento_ButtonClick(ByVal Button As MSComctlLib.Button)

If m_CambioDatos And Trim(txtExpediente.Text) <> "" Then
    m_MuestraMensaje = True
    If Not fxGuardar Then Exit Sub
End If

Select Case UCase(Button.Key)
    Case UCase("anterior")
        Call cmdScrollBar(0)
        Call sbActivarMontoGirar
        tcMain.Item(0).Selected = True
        
    Case UCase("Siguiente")
        Call cmdScrollBar(1)
        Call sbActivarMontoGirar
        tcMain.Item(0).Selected = True
End Select

End Sub

Private Sub sbEstudio_Comite_Resolucion_Load(pExpediente As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim pTipo As String, itmX As ListViewItem


On Error GoTo vError

Me.MousePointer = vbHourglass

lswAutorizadores.Checkboxes = False
lswAutorizadores.ListItems.Clear

With lswAutorizadores.ColumnHeaders
    .Clear
    
    Select Case True
        Case rbActas.Item(0).Value
          pTipo = "RES"
          .Add , , "Estado", 1800, vbCenter
          .Add , , "Fecha", 1800
          .Add , , "Usuario", 2100, vbCenter
          .Add , , "Notas", 3100
        
        Case rbActas.Item(1).Value
          pTipo = "AUT"
          .Add , , "Estado", 1800, vbCenter
          .Add , , "Fecha", 1800
          .Add , , "Identificación", 1800
          .Add , , "Nombre", 3800
          .Add , , "Usuario", 2100, vbCenter
          .Add , , "Notas", 2100
          
        
        Case rbActas.Item(2).Value
          pTipo = "ASI"
          .Add , , "Fecha", 1800
          .Add , , "Identificación", 1800
          .Add , , "Nombre", 3800
          .Add , , "Usuario", 2100, vbCenter
    End Select
    

End With

strSQL = "exec spCrd_Estudio_Resolucion_Detalle '" & pExpediente & "', '" & pTipo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    txtActa.Text = rs!Acta
    txtActaFecha.Text = Format(rs!Acta_Fecha, "dd/MM/yyyy")
End If

Do While Not rs.EOF

    Select Case True
     Case rbActas.Item(0).Value 'Resoluciones
        
        Set itmX = lswAutorizadores.ListItems.Add(, , rs!Estado)
            itmX.SubItems(1) = rs!Registro_Fecha
            itmX.SubItems(2) = rs!Registro_Usuario
            itmX.SubItems(3) = rs!Notas
     
     Case rbActas.Item(1).Value 'Autorizaciones
        
        Set itmX = lswAutorizadores.ListItems.Add(, , rs!Estado)
            itmX.SubItems(1) = rs!Registro_Fecha
            itmX.SubItems(2) = rs!Cedula
            itmX.SubItems(3) = rs!Nombre
            itmX.SubItems(4) = rs!Registro_Usuario
            itmX.SubItems(5) = rs!Notas
            
          
        
     Case rbActas.Item(2).Value 'Asistencia
   
        Set itmX = lswAutorizadores.ListItems.Add(, , rs!Registro_Fecha)
            itmX.SubItems(1) = rs!Cedula
            itmX.SubItems(2) = rs!Nombre
            itmX.SubItems(3) = rs!Registro_Usuario
    
    End Select


  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 1 Then
    rbActas.Item(0).Value = True
  Call sbEstudio_Comite_Resolucion_Load(txtExpediente.Text)
End If
End Sub


Private Sub sbTabChange(Index As Integer)
On Error GoTo vError

Me.MousePointer = vbHourglass

m_PreviousTab = Index
m_MuestraMensaje = False
m_CargoSalario = False

FrameSalarios.Visible = False

lblPorcentajeSalario.Caption = str(GlobalPorcLiquidezLibre) & " % Sobre Salario "

Select Case Index
    Case 0
            
            
            txtCedula.SetFocus
            
            If m_CambioDatos Then
                Call fxGuardar
            End If
    
    Case 1
    
            cboSalario.SetFocus
        
        If Len(txtExpediente.Text) = 0 Then
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        If m_CambioCalculo Then
            If m_valorComboExp <> "Nuevo SubExpediente" Then
                
                If ((clsMensajes.Estado = "R") Or (clsMensajes.Estado = "P")) Then
                        Call fxGuardar
                End If
            End If
        End If
    
    
    Case 2
        tcAux.Item(0).Selected = True
        If ((clsMensajes.Estado = "R") Or (clsMensajes.Estado = "P")) Then
            If m_CambioObservaciones Then
                Call sbGuardaObservaciones
            End If
        End If
End Select


If Index = 3 Then
    DoEvents
  '  Call CargarGrid
End If

'cboSalario.SetFocus
m_CambioCalculo = False
m_CambioObservaciones = False
m_CambioDatos = False

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    txtMontoGirar.Visible = False
    lblMontoGirar.Item(34).Visible = False
Else
    txtMontoGirar.Visible = True
    lblMontoGirar.Item(34).Visible = True
End If


Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)


If txtExpediente.Text = "" And CCur(txtMonto.Text) > 0 And pTabBefore > 0 Then
    pTabBefore = 0
End If

Call sbTabChange(pTabBefore)
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

'Call sbTabChange(Item.Index)
pTabBefore = Item.Index


Select Case Item.Index
    Case 1
        FrameSalarios.Visible = False
        fraDCargas.Visible = False

    Case 2 'Observaciones
        optObservacion(0).Value = True
        txtObservaciones.Text = vObservacion(0)
    
    Case 4 'Tags
       Call sbCargarListaTags
    
    Case 5 'Causas
       Call sbCargarListaCausas

End Select

End Sub

Public Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

m_MuestraMensaje = False

Select Case UCase(Button.Key)

    Case "INSERTAR", "NUEVO"
      Call SbAccionVentana(NuevoRegistro)
      Call sbLimipiaControles(Me, True)
      Call sbToolBar(Me.tlb, "edicion")
      Call sbInicializaComboExpediente
      txtExpediente.Locked = False
      
    Case "MODIFICAR", "EDITAR"
      Call sbToolBar(Me.tlb, "edicion")
      txtCedula.SetFocus
      Call sbBloquearTab
    
    Case "BORRAR"
         Call sbBorrar
         
    Case "GUARDAR", "SALVAR"
        m_MuestraMensaje = True
        If tcMain.Item(0).Selected = True Then
            If m_CambioDatos = False Then Exit Sub
        ElseIf tcMain.Item(1).Selected = True Then
            If m_CambioCalculo = False Then Exit Sub
        ElseIf tcMain.Item(2).Selected = True Then
            If m_CambioObservaciones = False Then Exit Sub
        End If
        Call fxGuardar
        
    Case "DESHACER"
      
      Call SbAccionVentana(NuevoRegistro)
      'Call sbToolBar(Me.tlb, "nuevo")
      m_DesplegoMensaje = True
      Call sbLimipiaControles(Me, True)
      Call sbInicializaComboExpediente
      Call sbBloquearControles(Me, Expediente)
      tcMain.Item(0).Selected = True
      txtCedula.SetFocus
      m_CambioDatos = False
      m_CambioCalculo = False
      m_CambioObservaciones = False
      'TxtExpediente.Locked = True
      
      Call sbToolBar(Me.tlb, "nuevo")
         
    Case "CONSULTAR"
        Call sbMostraVentanBusqueda
        
    Case "REPORTES"
        If m_ventanaEnModo = ModificarRegistro Then
           
           If Trim(txtExpediente.Text) <> Trim(cboSubExpediente.Text) Then
            MsgBox "Debe selecionar un expediente o sub expediente válido.", vbInformation, gMsgTitulo
            Exit Sub
           End If
           
           If Len(txtExpediente.Text) = 0 Then Exit Sub
            'Actualiza Datos
            Call fxGuardar
            
            'Verifica y Guarda en caso de cambios
            'TODO: Call ssTab_Click(ssTab.Tab)
            Call sbTabChange(tcMain.SelectedItem)

            gPreAnalisis.Expediente = txtExpediente.Text
            frmPreaSubReporte.Show vbModal
        End If
         
End Select

Call RefrescaTags(Me)


End Sub


Private Sub txtCedula_Change()
    txtNombre.Text = ""
    lblEdad.Caption = ""
    
     Call sbBloquearTab
     m_CambioDatos = True
     If cboSubExpediente.Text = "Nuevo Expediente" Then
        cboCantidadFiadores.Enabled = True
     End If
     
lblEstadoSocio.Visible = False
lblEstadoSocio.Caption = ""
End Sub





Private Function fxExiteCedulaEnSubExpediente() As Boolean
On Error GoTo vError

fxExiteCedulaEnSubExpediente = False

Dim vExpPadre As String

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    vExpPadre = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
Else
    vExpPadre = txtExpediente.Text
End If


glogon.strSQL = "select nombre from CRD_PREA_PREANALISIS where COD_PREANALISIS = '" & vExpPadre & "' and cedula = '" & txtCedula.Text & "'"
If execSql(glogon.strSQL, True) Then
    
    MsgBox glogon.Recordset!Nombre & " con numero de cedula " & txtCedula.Text & " ya existe como un expediente Maestro, verifique e intente de nuevo.", vbInformation, gMsgTitulo
    fxExiteCedulaEnSubExpediente = True
End If


    Exit Function
vError:
    MsgBox "Ocurrió un error al validar que el numero de cedula. " & "-" & Err.Description, vbCritical, gMsgTitulo
    
End Function


Public Sub txtCedula_LostFocus()
On Error GoTo vError

Dim RsTemp As ADODB.Recordset

If Len(txtCedula.Text) = 0 Then Exit Sub

txtNombre.Text = ""
gPreAnalisis.Institucion = "-1"
gPreAnalisis.Socio = "N"

If (cboSubExpediente.Text = "Nuevo SubExpediente" Or InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0) Then
    If fxExiteCedulaEnSubExpediente Then Exit Sub
End If

glogon.strSQL = "select S.nombre, S.cod_Institucion, S.ESTADOACTUAL, S.FECHA_NAC, S.sexo" _
              & ",dbo.MyGetdate() as 'FechaSistema', isnull(E.descripcion,'') as 'EstadoPersona'" _
              & ", isnull(I.Frecuencia,'M') as 'Frecuencia_Id'" _
              & " from socios S left join AFI_ESTADOS_PERSONA E on S.EstadoActual = E.cod_Estado" _
              & " left join Instituciones I on S.cod_institucion = I.cod_Institucion" _
              & " where S.cedula = '" & txtCedula.Text & "'"


mFrecuenciaPago = "M"
If execSql(glogon.strSQL, True) Then

    Set RsTemp = glogon.Recordset
    
    
    mFrecuenciaPago = RsTemp!Frecuencia_ID
    
    txtNombre.Text = IIf(IsNull(RsTemp!Nombre), "", RsTemp!Nombre)
    gPreAnalisis.Institucion = IIf(IsNull(Trim(RsTemp!cod_institucion & "")), "-1", Trim(RsTemp!cod_institucion & ""))
    
    gPreAnalisis.Socio = IIf(Trim(RsTemp!EstadoActual & "") = "", "N", Trim(RsTemp!EstadoActual))
    
    Call sbSeleccionaSexo(cboSexo, Trim(RsTemp!sexo & ""))
    
    If Trim(RsTemp!fecha_nac & "") <> "" Then
        dtpFecNac.Value = RsTemp!fecha_nac
    Else
        dtpFecNac.Value = Format(RsTemp!FechaSistema, "dd/mm/yyyy")
    End If

    lblEstadoSocio.Visible = True
    lblEstadoSocio.Caption = Trim(RsTemp!EstadoPersona)

    txtNombre.Locked = True
    
    Call sbBloquearTab
Else
    glogon.strSQL = "select nombre,FECHA_NACIMIENTO,sexo,dbo.MyGetdate() as 'FechaSistema', 'No Socio' as 'EstadoPersona'" _
                  & " from CRD_PREA_PREANALISIS" _
                  & " where cedula = '" & txtCedula.Text & "'"
    
    If execSql(glogon.strSQL, True) Then
    
        Set RsTemp = glogon.Recordset
        txtNombre.Text = IIf(IsNull(RsTemp!Nombre), "", RsTemp!Nombre)
        gPreAnalisis.Socio = "N"
        
        Call sbSeleccionaSexo(cboSexo, Trim(RsTemp!sexo & ""))
        
        If Trim(RsTemp!fecha_nacimiento & "") <> "" Then
            dtpFecNac.Value = RsTemp!fecha_nacimiento
        Else
            dtpFecNac.Value = Format(RsTemp!FechaSistema, "dd/mm/yyyy")
        End If
    
    lblEstadoSocio.Visible = True
    lblEstadoSocio.Caption = Trim(RsTemp!EstadoPersona)
    
    Else
        txtNombre.Locked = False
    End If
txtNombre.Locked = False
    
End If


Call sbNumPagos_Update

Call sbCalcularPlazoMaximo


    Exit Sub
vError:
    MsgBox "Ocurrió un error al traer los datos del expediente. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub


Private Sub txtDesLineaCredito_Change()
 m_CambioDatos = True
End Sub


Private Sub TxtExpediente_Change()

If cboSubExpediente.ListIndex <> -1 Then
    If fxSelectItemSubExpediente(cboSubExpediente.ItemData(cboSubExpediente.ListIndex)) = "S" Then
        Call SbAccionVentana(NuevoRegistro)
    End If
End If
End Sub


Public Sub txtExpediente_LostFocus()
If Len(txtExpediente.Text) = 0 Then Exit Sub

vPaso = True
    Call sbTraerNumFiadores
    Call fxValidaNumFiadoresRegistrados
    Call sbTraerDatosExpediente

    m_CargoSalario = True
    m_CargoCombo = False
    m_CambioDatos = False
    m_CambioCalculo = False
    m_CambioObservaciones = False
    
vPaso = False

Call sbAplicarFormulas(eFormulas.eAplicarTodas)

End Sub


Private Sub txtExtrasFijas_GotFocus()
    
    If txtExtrasFijas = "" Then
        txtExtrasFijas = 0
    End If
    
    m_curValor_Anterior = txtExtrasFijas
    
    txtExtrasFijas.SelStart = 0
    txtExtrasFijas.SelLength = Len(txtExtrasFijas.Text)
    
End Sub

Private Sub txtExtrasFijas_LostFocus()
    
    If txtExtrasFijas = "" Then
        txtExtrasFijas = 0
    End If
    
    If txtExtrasFijas <> m_curValor_Anterior Then
        Call sbEstructuraActualiza(2, False)
        m_CambioCalculo = True
    End If

End Sub

Private Sub txtLinea_Change()
    txtDesLineaCredito.Text = ""
    chkPrimerCuota.Value = 0
    m_CambioDatos = True
    cboGarantia.Clear
End Sub

Private Sub txtLinea_LostFocus()

If Len(txtLinea.Text) = 0 Then Exit Sub
txtDesLineaCredito.Text = fxDescLineaCredito(Trim(txtLinea.Text))

If Len(Trim(txtDesLineaCredito.Text)) > 0 Then
    Call sbCalcularCuota("txtMonto")
    cboDestino.Clear
    Call sbLlenarComboFiltrado(cboDestino, "spCRDPreaDestinos", "cod_destino", "DescDestino", "Linea", "Seleccione un destino", fxFormatearValor(txtLinea.Text, caracter))
        
    If Len(txtDesLineaCredito.Text) > 0 Then
        Call sbSTCargaCboGarantia(cboGarantia, txtLinea.Text)
    End If
End If

'Garantia en Fondos
If fxGarantia(cboGarantia.Text) = "Y" Then
   Call cboFondo_Click
   cboFondo.Enabled = True
Else
   cboFondo.Enabled = False
End If
End Sub


Private Sub txtMonto_GotFocus()
    
    If Not IsNumeric(txtMonto.Text) Then
        txtMonto.Text = 0
    End If
    
    If m_curValor_Anterior <> 0 Then
        m_curValor_Anterior = txtMonto
    End If
    
    txtMonto.SelStart = 0
    txtMonto.SelLength = Len(txtMonto.Text)
    
End Sub

Private Sub txtMonto_LostFocus()
        
    If Not IsNumeric(txtMonto.Text) Then
        txtMonto.Text = 0
    End If
        
    If CCur(txtMonto.Text) <> m_curValor_Anterior Then
        Call sbCalculaPolizaDeVida
        Call sbCalculaPolizaDeIncendio
        
        Call sbCalcularCuota("txtMonto")
        
        Call sbCalculaPolizaDesempleo
        
        Call sbEstructuraActualiza(1000, True)
        m_CambioDatos = True
        m_curValor_Anterior = CCur(txtMonto.Text)
    End If
    
End Sub



Private Sub txtNombre_LostFocus()
Call sbBloquearTab
End Sub

Private Sub txtObservaciones_Change()

If optObservacion.Item(0) = True Then
   vObservacion(0) = txtObservaciones.Text
   m_CambioObservaciones = True
ElseIf optObservacion.Item(1) = True Then
   vObservacion(1) = txtObservaciones.Text
   m_CambioObservaciones = True
ElseIf optObservacion.Item(2) = True Then
   vObservacion(2) = txtObservaciones.Text
   m_CambioObservaciones = True
Else
m_CambioObservaciones = False
End If

End Sub



Private Sub txtPlazo_GotFocus()
    txtPlazo.SelStart = 0
    txtPlazo.SelLength = Len(txtPlazo.Text)
End Sub

Private Sub txtPlazo_LostFocus()
If IsNumeric(txtPlazo.Text) Then
    
    If Val(txtPlazo.Text) = 0 Then Exit Sub
    Call sbCalcularCuota("txtPlazo")
    Call sbCalcularPlazoMaximo
    Call sbEstructuraActualiza(1000, True)
    m_CambioDatos = True
    
End If
End Sub

Private Sub txtPlazo_Validate(Cancel As Boolean)
On Error GoTo vError
If Val(txtPlazo.Text) = 0 Then
    txtPlazo.Text = 0
    txtCuota.Text = 0
End If
 
 

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto digitado. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtMonto_Validate(Cancel As Boolean)
On Error GoTo vError
If Val(txtMonto.Text) = 0 Then
    txtMonto.Text = Format(0, "Standard")
Else
    txtMonto.Text = Format(txtMonto.Text, "Standard")
End If



    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto digitado. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub



Private Sub txtPolizaIncendio_Validate(Cancel As Boolean)
On Error GoTo vError
txtPolizaIncendio.Text = Format(txtPolizaIncendio.Text, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto de poliza de incendio. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub



Private Sub txtPolizaVida_Validate(Cancel As Boolean)
On Error GoTo vError
txtPolizaVida.Text = Format(txtPolizaVida.Text, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto de poliza de vida. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub






Private Sub txtSalarioDevengado_GotFocus()
    
    If txtSalarioDevengado = "" Then
        txtSalarioDevengado = 0
    End If

    m_curValor_Anterior = txtSalarioDevengado
    
    txtSalarioDevengado.SelStart = 0
    txtSalarioDevengado.SelLength = Len(txtSalarioDevengado.Text)
    
End Sub



Private Sub txtSalarioDevengado_LostFocus()

On Error GoTo vError

    If txtSalarioDevengado = "" Then
        txtSalarioDevengado = 0
    End If

    If txtSalarioDevengado <> m_curValor_Anterior Then
        Call sbEstructuraActualiza(1, False)
        m_CambioCalculo = True
    End If

    txtSalarioDevengado.Text = Format(txtSalarioDevengado.Text, "Standard")
    
    txtRebajoExtras.SetFocus

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtSalarioLiquido_GotFocus()
On Error GoTo vError
    m_curValor_Anterior = txtSalarioLiquido

Exit Sub

vError:
m_curValor_Anterior = 0
End Sub

Private Sub txtSalarioLiquido_LostFocus()
    ''Call sbAplicarFormulas(eFormulas.eTotalLiquido)
    If txtSalarioLiquido = "" Then
        txtSalarioLiquido = 0
    End If
    
    If txtSalarioDevengado <> m_curValor_Anterior Then
        Call sbEstructuraActualiza(4, False)
        m_CambioCalculo = True
    End If
    
End Sub





'/Fin Tab Datos****************************************************************************


'Tab Calculo****************************************************************************

Private Sub txtRebajoExtras_Validate(Cancel As Boolean)
On Error GoTo vError
txtRebajoExtras.Text = Format(txtRebajoExtras.Text, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el Rebajo Horas Extras. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub


Private Sub txtSalarioReal_Validate(Cancel As Boolean)
On Error GoTo vError
txtSalarioReal.Text = Format(txtSalarioReal.Text, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el Salario Real. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub
Private Sub txtExtrasFijas_Validate(Cancel As Boolean)
On Error GoTo vError
txtExtrasFijas.Text = Format(txtExtrasFijas.Text, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el (+) Extras Fijas . " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtDevengadoMes_Validate(Cancel As Boolean)
On Error GoTo vError
txtDevengadoMes.Text = Format(txtDevengadoMes.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el Devengado del Mes. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtTasa_GotFocus()

    txtTasa.SelStart = 0
    txtTasa.SelLength = Len(txtTasa.Text)

End Sub

Private Sub txtTasa_LostFocus()
On Error GoTo vError
    
        If IsNumeric(txtTasa.Text) Then
            Call sbCalcularCuota("txtTasa")
        Else
            txtTasa.Text = 0
        End If
        m_CambioDatos = True
        Call sbEstructuraActualiza(1000, True)
    
    Exit Sub
vError:
        MsgBox "Ocurrió un error al realizar el calculo de la cuota. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Private Sub txtTotal_Cargas_CCSS_Validate(Cancel As Boolean)
On Error GoTo vError
txtTotal_Cargas_CCSS.Text = Format(txtTotal_Cargas_CCSS.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) Cargas. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtPorcSobreSalario_Validate(Cancel As Boolean)
On Error GoTo vError
txtPorcSobreSalario.Text = Format(txtPorcSobreSalario.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (%?) Sobre Salario. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub


Private Sub txtCrdTransitoCancelados_Validate(Cancel As Boolean)
On Error GoTo vError
txtCrdTransitoCancelados.Text = Format(txtCrdTransitoCancelados.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (+) Créditos Cancelados. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtCrdTransitoXCobrar_Validate(Cancel As Boolean)
On Error GoTo vError
txtCrdTransitoXCobrar.Text = Format(txtCrdTransitoXCobrar.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) Créditos x Cobrar. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtSalarioLiquido_Validate(Cancel As Boolean)
On Error GoTo vError
txtSalarioLiquido.Text = Format(txtSalarioLiquido.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Salario Liquido. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub

Private Sub txtRefundiciones_Validate(Cancel As Boolean)
On Error GoTo vError
txtRefundiciones.Text = Format(txtRefundiciones.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (+) Refundiciones. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub

Private Sub txtDesembolsos_Validate(Cancel As Boolean)
On Error GoTo vError
txtDesembolsos.Text = Format(txtDesembolsos.Text, "Standard")
If Val(txtDesembolsos.Text) = 0 Then
    txtDesembolsos.ToolTipText = Format(0, "Standard")
End If



    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (+) Desembolsos. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtTotalLiquido_Validate(Cancel As Boolean)
On Error GoTo vError
txtTotalLiquido.Text = Format(txtTotalLiquido.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Total Liquido. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtFianzas_Validate(Cancel As Boolean)
On Error GoTo vError
txtFianzas.Text = Format(txtFianzas.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Total Fianzas. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtLiquidezPorcSinFianza_Validate(Cancel As Boolean)
On Error GoTo vError
txtLiquidezPorcSinFianza.Text = Format(txtLiquidezPorcSinFianza.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Total [%] Sin Fianzas. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtLiquidezConFianza_Validate(Cancel As Boolean)
On Error GoTo vError
txtLiquidezConFianza.Text = Format(txtLiquidezConFianza.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Total Con Fianzas. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub

Private Sub txtLiquidezPorcConFianza_Validate(Cancel As Boolean)
On Error GoTo vError
txtLiquidezPorcConFianza.Text = Format(txtLiquidezPorcConFianza.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar [%] Con Fianzas. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtMontoGirar_Validate(Cancel As Boolean)
On Error GoTo vError
txtMontoGirar.Text = Format(txtMontoGirar.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Monto a Girar. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtCargaCCSS_Validate(Cancel As Boolean)
On Error GoTo vError
 lblCargaCCSS.Caption = Format(lblCargaCCSS.Caption, "Standard")
 If fraDCargas.Visible = True Then
    'txtCargaImpSalario.SetFocus
 End If
 

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) C.C.S.S.. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Private Sub txtCargaImpSalario_Validate(Cancel As Boolean)
On Error GoTo vError
lblCargaImpSalario.Caption = Format(lblCargaImpSalario.Caption, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) Imp.Salario. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtCargaAsociacion_Validate(Cancel As Boolean)
On Error GoTo vError
 lblCargaAsociacion.Caption = Format(lblCargaAsociacion.Caption, "Standard")
 If fraDCargas.Visible = True Then
    'txtCargaFrap.SetFocus
 End If
 

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) Asociación. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
    
End Sub
Private Sub txtCargaFrap_Validate(Cancel As Boolean)
On Error GoTo vError
 lblCargaFrap.Caption = Format(lblCargaFrap.Caption, "Standard")
 If fraDCargas.Visible = True Then
    'txtCargaCCSS.SetFocus
 End If
 

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) FAP/FRAP. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub




Private Sub txtMontoGirar_Change()

If Val(txtMontoGirar.Text) > 0 Then
    Timer.Enabled = True
Else
    Timer.Enabled = False
End If

End Sub








'Fin Tab Calculo *****************************************************************************

'Inicio Tab Clasificacion ********************************************************************
Private Sub CargarGrid()
Dim sql As String

On Error GoTo error

vGrid.MaxCols = 3
        
If Val(txtLiquidezPorcConFianza.Text) = 0 Then
    txtLiquidezPorcConFianza.Text = 0
End If

sql = "exec spCRDPreaClasificacionNew " & fxFormatearValor(txtCedula.Text, caracter) & "," & fxFormatearValor(CDbl(txtLiquidezPorcConFianza.Text), Numerico) & ", " & fxFormatearValor(gPreAnalisis.Expediente, caracter)

'Call clsEntidad.fxEjecutaSQL(sql)

Call OpenRecordSet(glogon.Recordset, sql)
Call sbCargaGridLocal(vGrid, 3)

    
    Exit Sub
error:
    Call cMensaje.deError("Ocurrió un erro  al traer la información solicitada. Error " & Err.Description)
    
End Sub

Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer)

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 0

With glogon.Recordset
    Do While Not .EOF
       vGrid.MaxRows = vGrid.MaxRows + 1
       
       vGrid.Row = vGrid.MaxRows
       vGrid.Col = 1
       vGrid.Text = !Codigo
       
       vGrid.Col = 2
       vGrid.Text = !DESCRIPCION
       
       vGrid.Col = 3
       vGrid.Text = !Razon
    
       vGrid.Col = 1
        Select Case LCase(!Color)
            Case "rojo"
                 vGrid.BackColor = &HFF&
            Case "verde"
                 vGrid.BackColor = &H80FF80
            Case "amarillo"
                vGrid.BackColor = &HFFFF&
        End Select
    
      .MoveNext
    Loop
    .Close
End With


End Sub


Private Function fxColorCell(ByRef vGrid As Object, _
                             ByVal Row As Integer, _
                             ByVal Col As Integer, _
                             ByVal strcolor As String) As String
vGrid.Row = Row
vGrid.Col = Col
Select Case LCase(strcolor)
    Case "rojo"
         vGrid.BackColor = &HFF&
    Case "verde"
         vGrid.BackColor = &H80FF80
    Case "amarillo"
        vGrid.BackColor = &HFFFF&
End Select
End Function
'Fin tab Clasificación**********************************************************

Private Sub sbCargaCboComites(Optional pComite As String)
Dim strSQL As String, rs As New ADODB.Recordset
    
    strSQL = "Select id_comite,descripcion from comites where estado = 1"
    Call OpenRecordSet(rs, strSQL)
    cboComite.Clear
    If rs.EOF And rs.BOF Then
        MsgBox "No existen Comités creados...(Debe Crearlos)", vbCritical
    Else
    
        cboComite.AddItem " "
        cboComite.ItemData(cboComite.NewIndex) = 0
        
        Do While Not rs.EOF
            cboComite.AddItem Trim(rs!DESCRIPCION & "")
            cboComite.ItemData(cboComite.NewIndex) = rs!ID_COMITE
         rs.MoveNext
        Loop
'        rs.MoveFirst
'        cboComite.Text = Trim(rs!Descripcion)
        cboComite.Text = " "
    End If
    rs.Close

End Sub

Private Function fxComite(ByVal pId_Comite As Integer) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As String

    strSQL = "Select id_comite,descripcion as ItemX from comites where  id_comite = " & pId_Comite
    Call OpenRecordSet(rs, strSQL)
    If rs.EOF And rs.BOF Then
        vResultado = " "
    Else
        vResultado = Trim(rs!iTemX)
    End If
    rs.Close

    fxComite = vResultado

End Function

Private Sub sbCargarListaTags()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

vGridTags.MaxRows = 0
vGridTags.MaxCols = 5

If Val(Trim(txtAsignado)) > 0 Then
    strSQL = "select O.*,T.descripcion as Etiqueta" _
           & " from CRD_OPERACION_TAGS O inner join Crd_Tags T on O.Tag_codigo = T.Tag_Codigo" _
           & " where O.id_solicitud = " & Trim(txtAsignado.Text) & " order by O.registro_fecha "
Else
    strSQL = "select O.*,T.descripcion as Etiqueta" _
           & " from CRD_PREA_TAGS O inner join Crd_Tags T on O.Tag_codigo = T.Tag_Codigo" _
           & " where O.COD_PREANALISIS = '" & Trim(txtExpediente.Text) & "' order by O.registro_fecha "
End If
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  vGridTags.MaxRows = vGridTags.MaxRows + 1
  vGridTags.Row = vGridTags.MaxRows
  
  
  For i = 1 To vGridTags.MaxCols
    vGridTags.Col = i
    Select Case i
        Case 1
            vGridTags.Tag = rs!Linea
            vGridTags.Text = rs!Registro_Fecha & ""
        Case 2
            vGridTags.Text = rs!Registro_Usuario & ""
        Case 3
            vGridTags.Text = rs!Etiqueta & ""
        Case 4
            vGridTags.Text = rs!Notas & ""
        Case 5
            vGridTags.Text = rs!Asignado_A & ""
    End Select
  Next i
  vGridTags.RowHeight(vGridTags.Row) = vGridTags.MaxTextRowHeight(vGridTags.Row)
  
 rs.MoveNext
Loop
rs.Close
 
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCargarListaCausas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, pTipo As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
  Case optCausas.Item(0).Value
    pTipo = "D"
  Case optCausas.Item(1).Value
    pTipo = "P"
End Select

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Código", 1200
lsw.ColumnHeaders.Add , , "Descripción", 3200
lsw.ColumnHeaders.Add , , "Fecha", 2800
lsw.ColumnHeaders.Add , , "Usuario", 2800

strSQL = "select Pa.*, Cg.DESCRIPCION " _
       & " from CRD_PREA_GESTION Pa inner join OPERACION_CAUSAS Cg on Pa.COD_CAUSAS = Cg.COD_CAUSAS and Pa.TIPO = Cg.TIPO" _
       & " where Pa.COD_PREANALISIS = '" & Trim(txtExpediente.Text) & "' and Pa.TIPO = '" & pTipo & "'" _
       & " order by REGISTRO_FECHA"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Cod_Causas)
     itmX.SubItems(1) = rs!DESCRIPCION
     itmX.SubItems(2) = rs!Registro_Fecha & ""
     itmX.SubItems(3) = rs!Registro_Usuario & ""
    
 rs.MoveNext
Loop
rs.Close
 
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


