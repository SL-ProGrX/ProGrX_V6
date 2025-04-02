VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmRH_Informes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "RRHH: Informes del Módulo"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13095
   LinkTopic       =   "Form8"
   ScaleHeight     =   7590
   ScaleWidth      =   13095
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   720
   End
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4692
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   4572
      _Version        =   1441793
      _ExtentX        =   8064
      _ExtentY        =   8276
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
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4692
      Left            =   4920
      TabIndex        =   2
      Top             =   1680
      Width           =   8052
      _Version        =   1441793
      _ExtentX        =   14203
      _ExtentY        =   8276
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
      Item(0).Caption =   "Filtros"
      Item(0).ControlCount=   24
      Item(0).Control(0)=   "cboNomina"
      Item(0).Control(1)=   "Label15(1)"
      Item(0).Control(2)=   "txtCentroCod"
      Item(0).Control(3)=   "txtCentroDesc"
      Item(0).Control(4)=   "txtDeptCodigo"
      Item(0).Control(5)=   "txtDeptDesc"
      Item(0).Control(6)=   "txtSecCodigo"
      Item(0).Control(7)=   "txtSecDesc"
      Item(0).Control(8)=   "txtPuestoCod"
      Item(0).Control(9)=   "txtPuestoDesc"
      Item(0).Control(10)=   "Label21(0)"
      Item(0).Control(11)=   "lblSeccion"
      Item(0).Control(12)=   "lblDepartamento"
      Item(0).Control(13)=   "Label10(12)"
      Item(0).Control(14)=   "dtpIngreso(0)"
      Item(0).Control(15)=   "Label1(2)"
      Item(0).Control(16)=   "chkIngreso"
      Item(0).Control(17)=   "dtpLiquida(0)"
      Item(0).Control(18)=   "Label1(3)"
      Item(0).Control(19)=   "dtpLiquida(1)"
      Item(0).Control(20)=   "chkLiquida"
      Item(0).Control(21)=   "cboEstado"
      Item(0).Control(22)=   "Label15(7)"
      Item(0).Control(23)=   "dtpIngreso(1)"
      Item(1).Caption =   "Add 1"
      Item(1).ControlCount=   17
      Item(1).Control(0)=   "cboContrato"
      Item(1).Control(1)=   "cboJornada"
      Item(1).Control(2)=   "cboVacaciones"
      Item(1).Control(3)=   "Label15(2)"
      Item(1).Control(4)=   "Label15(3)"
      Item(1).Control(5)=   "Label15(5)"
      Item(1).Control(6)=   "Label1(1)"
      Item(1).Control(7)=   "dtpCntVence(0)"
      Item(1).Control(8)=   "dtpCntVence(1)"
      Item(1).Control(9)=   "chkCntVence"
      Item(1).Control(10)=   "cboTipoId"
      Item(1).Control(11)=   "Label15(8)"
      Item(1).Control(12)=   "cboDivisa"
      Item(1).Control(13)=   "Label15(9)"
      Item(1).Control(14)=   "Label15(10)"
      Item(1).Control(15)=   "cboBancos"
      Item(1).Control(16)=   "cboFormaPago"
      Item(2).Caption =   "Add 2"
      Item(2).ControlCount=   27
      Item(2).Control(0)=   "cboSexo"
      Item(2).Control(1)=   "cboEstadoCivil"
      Item(2).Control(2)=   "cboNacionalidad"
      Item(2).Control(3)=   "Label14"
      Item(2).Control(4)=   "Label15(0)"
      Item(2).Control(5)=   "Label15(6)"
      Item(2).Control(6)=   "dtpNacimiento(0)"
      Item(2).Control(7)=   "Label1(0)"
      Item(2).Control(8)=   "dtpNacimiento(1)"
      Item(2).Control(9)=   "chkNacimiento"
      Item(2).Control(10)=   "cboProvincia"
      Item(2).Control(11)=   "cboCanton"
      Item(2).Control(12)=   "cboDistrito"
      Item(2).Control(13)=   "Label7(0)"
      Item(2).Control(14)=   "Label7(1)"
      Item(2).Control(15)=   "Label7(2)"
      Item(2).Control(16)=   "txtProfesionCod"
      Item(2).Control(17)=   "txtProfesionDesc"
      Item(2).Control(18)=   "cboNivel"
      Item(2).Control(19)=   "txtJefeCod"
      Item(2).Control(20)=   "txtJefeDesc"
      Item(2).Control(21)=   "Label21(1)"
      Item(2).Control(22)=   "Label15(4)"
      Item(2).Control(23)=   "Label9"
      Item(2).Control(24)=   "chkProvincia"
      Item(2).Control(25)=   "chkCantones"
      Item(2).Control(26)=   "chkDistritos"
      Begin XtremeSuiteControls.CheckBox chkNacimiento 
         Height          =   252
         Left            =   -65440
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.ComboBox cboSexo 
         Height          =   312
         Left            =   -65200
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4048
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
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoCivil 
         Height          =   312
         Left            =   -68440
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5106
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
      End
      Begin XtremeSuiteControls.ComboBox cboNacionalidad 
         Height          =   312
         Left            =   -68440
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5106
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
      End
      Begin XtremeSuiteControls.DateTimePicker dtpNacimiento 
         Height          =   312
         Index           =   0
         Left            =   -68440
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.DateTimePicker dtpNacimiento 
         Height          =   312
         Index           =   1
         Left            =   -67000
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.ComboBox cboProvincia 
         Height          =   312
         Left            =   -68440
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5106
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
      End
      Begin XtremeSuiteControls.ComboBox cboCanton 
         Height          =   312
         Left            =   -68440
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5106
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
      End
      Begin XtremeSuiteControls.ComboBox cboDistrito 
         Height          =   312
         Left            =   -68440
         TabIndex        =   20
         Top             =   2520
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5106
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
      End
      Begin XtremeSuiteControls.CheckBox chkProvincia 
         Height          =   252
         Left            =   -65440
         TabIndex        =   24
         Top             =   1800
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkCantones 
         Height          =   252
         Left            =   -65440
         TabIndex        =   25
         Top             =   2160
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkDistritos 
         Height          =   252
         Left            =   -65440
         TabIndex        =   26
         Top             =   2520
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtProfesionCod 
         Height          =   312
         Left            =   -68440
         TabIndex        =   27
         Top             =   3120
         Visible         =   0   'False
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtProfesionDesc 
         Height          =   312
         Left            =   -67720
         TabIndex        =   28
         Top             =   3120
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1441793
         _ExtentX        =   8488
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
      Begin XtremeSuiteControls.ComboBox cboNivel 
         Height          =   312
         Left            =   -68440
         TabIndex        =   29
         Top             =   3600
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9763
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
      End
      Begin XtremeSuiteControls.FlatEdit txtJefeCod 
         Height          =   312
         Left            =   -68440
         TabIndex        =   30
         Top             =   4080
         Visible         =   0   'False
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtJefeDesc 
         Height          =   312
         Left            =   -67720
         TabIndex        =   31
         Top             =   4080
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1441793
         _ExtentX        =   8488
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
      Begin XtremeSuiteControls.ComboBox cboNomina 
         Height          =   312
         Left            =   1680
         TabIndex        =   35
         Top             =   480
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9763
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
      End
      Begin XtremeSuiteControls.ComboBox cboContrato 
         Height          =   312
         Left            =   -68440
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9763
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
      End
      Begin XtremeSuiteControls.ComboBox cboJornada 
         Height          =   312
         Left            =   -68440
         TabIndex        =   38
         Top             =   1440
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9763
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
      End
      Begin XtremeSuiteControls.ComboBox cboVacaciones 
         Height          =   312
         Left            =   -68440
         TabIndex        =   39
         Top             =   1920
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9763
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
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCntVence 
         Height          =   312
         Index           =   0
         Left            =   -68440
         TabIndex        =   43
         Top             =   840
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.DateTimePicker dtpCntVence 
         Height          =   312
         Index           =   1
         Left            =   -67000
         TabIndex        =   45
         Top             =   840
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.CheckBox chkCntVence 
         Height          =   252
         Left            =   -65440
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCentroCod 
         Height          =   312
         Left            =   1680
         TabIndex        =   47
         Top             =   1080
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCentroDesc 
         Height          =   312
         Left            =   2400
         TabIndex        =   48
         Top             =   1080
         Width           =   4812
         _Version        =   1441793
         _ExtentX        =   8488
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
      Begin XtremeSuiteControls.FlatEdit txtDeptCodigo 
         Height          =   312
         Left            =   1680
         TabIndex        =   49
         Top             =   1440
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDeptDesc 
         Height          =   312
         Left            =   2400
         TabIndex        =   50
         Top             =   1440
         Width           =   4812
         _Version        =   1441793
         _ExtentX        =   8488
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
      Begin XtremeSuiteControls.FlatEdit txtSecCodigo 
         Height          =   312
         Left            =   1680
         TabIndex        =   51
         Top             =   1800
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSecDesc 
         Height          =   312
         Left            =   2400
         TabIndex        =   52
         Top             =   1800
         Width           =   4812
         _Version        =   1441793
         _ExtentX        =   8488
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
      Begin XtremeSuiteControls.FlatEdit txtPuestoCod 
         Height          =   312
         Left            =   1680
         TabIndex        =   53
         Top             =   2280
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPuestoDesc 
         Height          =   312
         Left            =   2400
         TabIndex        =   54
         Top             =   2280
         Width           =   4812
         _Version        =   1441793
         _ExtentX        =   8488
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
      Begin XtremeSuiteControls.DateTimePicker dtpIngreso 
         Height          =   312
         Index           =   0
         Left            =   1680
         TabIndex        =   59
         Top             =   2880
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.DateTimePicker dtpIngreso 
         Height          =   312
         Index           =   1
         Left            =   3120
         TabIndex        =   61
         Top             =   2880
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.CheckBox chkIngreso 
         Height          =   252
         Left            =   4680
         TabIndex        =   62
         Top             =   2880
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpLiquida 
         Height          =   312
         Index           =   0
         Left            =   1680
         TabIndex        =   63
         Top             =   3360
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.DateTimePicker dtpLiquida 
         Height          =   312
         Index           =   1
         Left            =   3120
         TabIndex        =   65
         Top             =   3360
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.CheckBox chkLiquida 
         Height          =   252
         Left            =   4680
         TabIndex        =   66
         Top             =   3360
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   1680
         TabIndex        =   67
         Top             =   3960
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9763
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
      End
      Begin XtremeSuiteControls.ComboBox cboTipoId 
         Height          =   315
         Left            =   -68440
         TabIndex        =   69
         Top             =   2640
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
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
      End
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   315
         Left            =   -68440
         TabIndex        =   71
         Top             =   3240
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
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
      End
      Begin XtremeSuiteControls.ComboBox cboBancos 
         Height          =   330
         Left            =   -68440
         TabIndex        =   73
         Top             =   3840
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
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
      End
      Begin XtremeSuiteControls.ComboBox cboFormaPago 
         Height          =   330
         Left            =   -65560
         TabIndex        =   75
         Top             =   3840
         Visible         =   0   'False
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
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
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Metodo de Pago"
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
         Index           =   10
         Left            =   -69880
         TabIndex        =   74
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Divisa"
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
         Index           =   9
         Left            =   -69880
         TabIndex        =   72
         Top             =   3240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Identificación"
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
         Index           =   8
         Left            =   -69880
         TabIndex        =   70
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del Empleado"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   7
         Left            =   240
         TabIndex        =   68
         Top             =   3960
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidación"
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
         Index           =   3
         Left            =   240
         TabIndex        =   64
         Top             =   3360
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso"
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
         Left            =   240
         TabIndex        =   60
         Top             =   2880
         Width           =   1212
      End
      Begin VB.Label Label10 
         Caption         =   "Centro"
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
         Left            =   240
         TabIndex        =   58
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label lblDepartamento 
         Caption         =   "Departamento"
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
         Left            =   240
         TabIndex        =   57
         Top             =   1440
         Width           =   1332
      End
      Begin VB.Label lblSeccion 
         Caption         =   "Sección"
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
         Left            =   240
         TabIndex        =   56
         Top             =   1800
         Width           =   1572
      End
      Begin VB.Label Label21 
         Caption         =   "Puesto"
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
         Left            =   240
         TabIndex        =   55
         Top             =   2280
         Width           =   1572
      End
      Begin VB.Label Label1 
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
         Height          =   252
         Index           =   1
         Left            =   -69880
         TabIndex        =   44
         Top             =   840
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Régimen de Vacaciones"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   5
         Left            =   -69880
         TabIndex        =   42
         Top             =   1920
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Jornada"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   -69880
         TabIndex        =   41
         Top             =   1440
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Contrato"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   -69880
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nómina"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label9 
         Caption         =   "Profesión"
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
         Left            =   -69880
         TabIndex        =   34
         Top             =   3120
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Académico"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Index           =   4
         Left            =   -69880
         TabIndex        =   33
         Top             =   3480
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label21 
         Caption         =   "Jefe/Superior"
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
         Index           =   1
         Left            =   -69880
         TabIndex        =   32
         Top             =   4080
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito"
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
         Left            =   -69880
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantón"
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
         Index           =   1
         Left            =   -69880
         TabIndex        =   22
         Top             =   2160
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
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
         Left            =   -69880
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   1452
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
         Index           =   0
         Left            =   -69880
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidad"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   6
         Left            =   -69880
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   -69880
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label14 
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
         Left            =   -65200
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   852
      End
   End
   Begin XtremeSuiteControls.GroupBox gbReporte 
      Height          =   1092
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   12852
      _Version        =   1441793
      _ExtentX        =   22669
      _ExtentY        =   1926
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkResumen 
         Height          =   372
         Left            =   9240
         TabIndex        =   7
         Top             =   360
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Resumen"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   612
         Left            =   10800
         TabIndex        =   6
         Top             =   240
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Picture         =   "frmRH_Informes.frx":0000
         ImageAlignment  =   0
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption scReporte 
      Height          =   372
      Left            =   4920
      TabIndex        =   5
      Top             =   1200
      Width           =   8052
      _Version        =   1441793
      _ExtentX        =   14203
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption scLista 
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8488
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Informes Disponibles"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   6732
      _Version        =   1441793
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Informes de Recursos Humanos"
      ForeColor       =   16777215
      BackColor       =   -2147483633
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
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13212
   End
End
Attribute VB_Name = "frmRH_Informes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, mSubTitulo As String

Private Function fxFechaReportes(pInicio As Date, pCorte As Date) As String

fxFechaReportes = " in Date(" & Format(pInicio, "yyyy,mm,dd") & ")" _
                & " to Date(" & Format(pCorte, "yyyy,mm,dd") & ")"

End Function

Private Function fxFiltros(Optional pTipo As String = "") As String
Dim vFiltro As String, vTabla As String

vFiltro = ""

mSubTitulo = ""

vFiltro = "{vRH_Personas.COD_NOMINA} = '" & cboNomina.ItemData(cboNomina.ListIndex) & "'"

mSubTitulo = "Nómina: " & cboNomina.ItemData(cboNomina.ListIndex)

If Len(txtCentroCod.Text) > 0 Then
    vFiltro = vFiltro & " AND {vRH_Personas.COD_CENTRO} = '" & txtCentroCod.Text & "'"
    mSubTitulo = mSubTitulo & "¦ Centro: " & txtCentroCod.Text
End If
If Len(txtDeptCodigo.Text) > 0 Then
    vFiltro = vFiltro & " AND {vRH_Personas.COD_DEPARTAMENTO} = '" & txtDeptCodigo.Text & "'"
    mSubTitulo = mSubTitulo & "¦ Dept.: " & txtDeptCodigo.Text
End If
If Len(txtSecCodigo.Text) > 0 Then
    vFiltro = vFiltro & " AND {vRH_Personas.COD_SECCION} = '" & txtSecCodigo.Text & "'"
    mSubTitulo = mSubTitulo & "¦ Secc.: " & txtSecCodigo.Text
End If
If Len(txtPuestoCod.Text) > 0 Then
    vFiltro = vFiltro & " AND {vRH_Personas.COD_PUESTO} = '" & txtPuestoCod.Text & "'"
    mSubTitulo = mSubTitulo & "¦ Puesto: " & txtPuestoCod.Text
End If

If chkIngreso.Value = xtpUnchecked Then
    vFiltro = vFiltro & " AND {vRH_Personas.FECHA_INGRESO} " & fxFechaReportes(dtpIngreso(0).Value, dtpIngreso(1).Value)
    mSubTitulo = mSubTitulo & "¦ Fec.Ing.: " & Format(dtpIngreso(0).Value, "dd-mm-yyyy") & " a " & Format(dtpIngreso(1).Value, "dd-mm-yyyy")
End If

If chkLiquida.Value = xtpUnchecked Then
    vFiltro = vFiltro & " AND {vRH_Personas.LIQUIDA_FECHA} " & fxFechaReportes(dtpLiquida(0).Value, dtpLiquida(1).Value)
    mSubTitulo = mSubTitulo & "¦ Fec.Liq.: " & Format(dtpLiquida(0).Value, "dd-mm-yyyy") & " a " & Format(dtpLiquida(1).Value, "dd-mm-yyyy")
End If

If chkCntVence.Value = xtpUnchecked Then
    vFiltro = vFiltro & " AND {vRH_Personas.CONTRATO_VENCIMIENTO} " & fxFechaReportes(dtpCntVence(0).Value, dtpCntVence(1).Value)
    mSubTitulo = mSubTitulo & "¦ Cnt. Vence: " & Format(dtpCntVence(0).Value, "dd-mm-yyyy") & " a " & Format(dtpCntVence(1).Value, "dd-mm-yyyy")
End If

If cboEstado.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.ESTADO_PERSONA} = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Est.Per.: " & cboEstado.Text
End If

If cboContrato.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.CONTRATO_TIPO} = '" & cboContrato.ItemData(cboContrato.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Contrato: " & cboContrato.Text
End If

If cboJornada.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.JORNADA_TIPO} = '" & cboJornada.ItemData(cboJornada.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Jornada: " & cboJornada.Text
End If

If cboVacaciones.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.COD_VACA_REGIMEN} = '" & cboVacaciones.ItemData(cboVacaciones.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Vaca.Reg.: " & cboVacaciones.ItemData(cboVacaciones.ListIndex)
End If

If cboTipoId.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.TIPO_PERSONERIA} = '" & cboTipoId.ItemData(cboTipoId.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Tipo Id.: " & cboTipoId.ItemData(cboTipoId.ListIndex)
End If

If cboDivisa.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.COD_DIVISA} = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Divisa: " & cboDivisa.ItemData(cboDivisa.ListIndex)
End If

If cboEstadoCivil.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.ESTADO_CIVIL} = '" & cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Est.Civil: " & cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex)
End If

If cboNacionalidad.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.COD_NACIONALIDAD} = '" & cboNacionalidad.ItemData(cboNacionalidad.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Nacional.: " & cboNacionalidad.Text
End If

If chkNacimiento.Value = xtpUnchecked Then
    vFiltro = vFiltro & " AND {vRH_Personas.FECHA_NACIMIENTO} " & fxFechaReportes(dtpNacimiento(0).Value, dtpNacimiento(1).Value)
    mSubTitulo = mSubTitulo & "¦ Fec.Nac.: " & Format(dtpNacimiento(0).Value, "dd-mm-yyyy") & " a " & Format(dtpNacimiento(1).Value, "dd-mm-yyyy")
End If

If chkProvincia.Value = xtpUnchecked Then
    vFiltro = vFiltro & " AND {vRH_Personas.PROVINCIA} '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Provincia: " & cboProvincia.Text
End If
If chkCantones.Value = xtpUnchecked Then
    vFiltro = vFiltro & " AND {vRH_Personas.CANTON} '" & cboCanton.ItemData(cboCanton.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Cantón: " & cboCanton.Text
End If
If chkDistritos.Value = xtpUnchecked Then
    vFiltro = vFiltro & " AND {vRH_Personas.DISTRITO} '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Distrito: " & cboDistrito.Text
End If

If UCase(cboSexo.Text) <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.SEXO} = '" & Mid(cboSexo.Text, 1, 1) & "'"
    mSubTitulo = mSubTitulo & "¦ Sexo: " & cboSexo.Text
End If

If cboNivel.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.NIVEL_ACADEMICO} = '" & cboNivel.ItemData(cboNivel.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Niv.Acade.: " & cboNivel.ItemData(cboNivel.ListIndex)
End If


If cboBancos.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.COD_BANCO} = " & cboBancos.ItemData(cboBancos.ListIndex)
    mSubTitulo = mSubTitulo & "¦ Pago en: " & cboBancos.ItemData(cboBancos.ListIndex)
End If

If cboFormaPago.Text <> "TODOS" Then
    vFiltro = vFiltro & " AND {vRH_Personas.FORMA_PAGO} = '" & cboFormaPago.ItemData(cboFormaPago.ListIndex) & "'"
    mSubTitulo = mSubTitulo & "¦ Forma Pago: " & cboFormaPago.ItemData(cboFormaPago.ListIndex)
End If

If Len(txtProfesionCod.Text) > 0 Then
    vFiltro = vFiltro & " AND {vRH_Personas.COD_PROFESION} = '" & txtProfesionCod.Text & "'"
    mSubTitulo = mSubTitulo & "¦ Prof.Id: " & txtProfesionCod.Text
End If

If Len(txtJefeCod.Text) > 0 Then
    vFiltro = vFiltro & " AND {vRH_Personas.JEFE_ID} = '" & txtJefeCod.Text & "'"
    mSubTitulo = mSubTitulo & "¦ Jefe Id: " & txtJefeCod.Text
End If


fxFiltros = vFiltro

End Function


Private Sub sbReporte()

Dim vTitulo As String, vReporte As String

On Error GoTo vError

If scReporte.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass

vTitulo = scReporte.Caption


Select Case scReporte.Tag
    Case Is = "R001" 'Patron de Personas"
        vReporte = "RRHH_List_Padron_Personas"
    Case Is = "R002" 'Puestos y Salarios
        vReporte = "RRHH_List_Puestos_Salarios"

    Case Is = "R003" 'Nóminas y Salarios
        vReporte = "RRHH_List_Nominas_Salarios"
    Case Is = "R004" 'Centros, Departamentos
        vReporte = "RRHH_List_Centros_Dept"
    Case Is = "R005" 'Informe de Cesantía
        vReporte = "RRHH_List_Censantía"
    
    Case Is = "R006.0" 'Informe de Vacaciones
        vReporte = "RRHH_List_Vacaciones"
    Case Is = "R006.1" 'Informe de Permisos
        vReporte = "RRHH_List_Permisos"
    Case Is = "R006.2" 'Informe de Incapacidades
        vReporte = "RRHH_List_Incapacidadeas"
    Case Is = "R006.3" 'Informe de Personal a Cargo
        vReporte = "RRHH_List_Personal_aCargo"

    Case Is = "R007" 'Tipos de Contratos
        vReporte = "RRHH_List_Contratos"
    Case Is = "R008" 'Jornadas de Trabajo
        vReporte = "RRHH_List_Jornadas"
    Case Is = "R009" 'Nacionalidades
        vReporte = "RRHH_List_Nacionalidades"
    Case Is = "R009.1" 'Estado Civil
        vReporte = "RRHH_List_Estado_Civil"
    Case Is = "R009.2" 'Genero
        vReporte = "RRHH_List_Genero"
        
        
    Case Is = "R010" 'Informe para Estudio de Mercado
        vReporte = "RRHH_List_Estudio_Mercado"
    Case Is = "R011" 'Liquidaciones de Personal
        vReporte = "RRHH_List_Liquidaciones"
    Case Is = "R012.1" 'Gasto en Nómina por Centro Costo
        vReporte = "RRHH_List_Gastos_Centro_Costo"
    Case Is = "R012.2" 'Conceptos de Nómina aplicados
        vReporte = "RRHH_List_Conceptos_Aplicados"
    Case Is = "R013" 'Preliminar para Aguinaldos
        vReporte = "RRHH_List_Aguinaldos"

End Select


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes de RRHH"
 
 .Connect = glogon.ConectRPT
 
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='" & vTitulo & "'"
 .Formulas(4) = "fxSubTitulo='" & Mid(mSubTitulo, 1, 250) & "'"
 
 
 strSQL = fxFiltros("")
 
 If chkResumen.Value = xtpChecked Then
    .ReportFileName = SIFGlobal.fxPathReportes(vReporte & "_Resumen.rpt")
 Else
    .ReportFileName = SIFGlobal.fxPathReportes(vReporte & "_Detalle.rpt")
 End If
 .SelectionFormula = strSQL

 .PrintReport
End With

Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnReporte_Click()

Call sbReporte

End Sub

Private Sub cboCanton_Click()
If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
            & "' and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistrito, strSQL, False, True)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistrito.AddItem " "
cboDistrito.Text = " "
End Sub

Private Sub cboProvincia_Click()
If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click
End Sub

Private Sub Form_Load()
vModulo = 23


Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

tcMain.Item(0).Selected = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Informe:", lsw.Width - 100
End With

Dim itmX As ListViewItem

Set itmX = lsw.ListItems.Add(, , "Padron de Personas")
    itmX.Tag = "R001"
Set itmX = lsw.ListItems.Add(, , "Puestos y Salarios")
    itmX.Tag = "R002"
Set itmX = lsw.ListItems.Add(, , "Nóminas y Salarios")
    itmX.Tag = "R003"
Set itmX = lsw.ListItems.Add(, , "Centros, Departamentos")
    itmX.Tag = "R004"
Set itmX = lsw.ListItems.Add(, , "Informe de Cesantía")
    itmX.Tag = "R005"
Set itmX = lsw.ListItems.Add(, , "Informe de Vacaciones")
    itmX.Tag = "R006.0"
Set itmX = lsw.ListItems.Add(, , "Informe de Permisos")
    itmX.Tag = "R006.1"
Set itmX = lsw.ListItems.Add(, , "Informe de Incapacidades")
    itmX.Tag = "R006.2"
Set itmX = lsw.ListItems.Add(, , "Informe de Personal a Cargo")
    itmX.Tag = "R006.3"

Set itmX = lsw.ListItems.Add(, , "Tipos de Contratos")
    itmX.Tag = "R007"
Set itmX = lsw.ListItems.Add(, , "Jornadas de Trabajo")
    itmX.Tag = "R008"
Set itmX = lsw.ListItems.Add(, , "Nacionalidades")
    itmX.Tag = "R009"

Set itmX = lsw.ListItems.Add(, , "Estado Civil")
    itmX.Tag = "R009.1"
Set itmX = lsw.ListItems.Add(, , "Genero")
    itmX.Tag = "R009.2"

Set itmX = lsw.ListItems.Add(, , "Informe para Estudio de Mercado")
    itmX.Tag = "R010"
Set itmX = lsw.ListItems.Add(, , "Liquidaciones de Personal")
    itmX.Tag = "R011"
Set itmX = lsw.ListItems.Add(, , "Gasto en Nómina por Centro Costo")
    itmX.Tag = "R012.1"
Set itmX = lsw.ListItems.Add(, , "Conceptos de Nómina aplicados")
    itmX.Tag = "R012.2"

Set itmX = lsw.ListItems.Add(, , "Preliminar para Aguinaldos")
    itmX.Tag = "R013"



Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
  
  scReporte.Tag = Item.Tag
  scReporte.Caption = Item.Text
  
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub


Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaActual As Date


On Error GoTo vError

vFechaActual = Format(fxFechaServidor, "dd/mm/yyyy")

dtpIngreso.Item(0).Value = vFechaActual
dtpIngreso.Item(1).Value = vFechaActual

dtpNacimiento.Item(0).Value = vFechaActual
dtpNacimiento.Item(1).Value = vFechaActual

dtpLiquida.Item(0).Value = vFechaActual
dtpLiquida.Item(1).Value = vFechaActual

dtpCntVence.Item(0).Value = vFechaActual
dtpCntVence.Item(1).Value = vFechaActual

cboSexo.Clear
cboSexo.AddItem "Todos"
cboSexo.AddItem "Masculino"
cboSexo.AddItem "Femenino"
cboSexo.Text = "Todos"


strSQL = "select cod_nacionalidad as 'IdX', Descripcion as 'ItmX' from sys_nacionalidades" _
       & " where Activo = 1" _
       & " order by Omision desc, Descripcion asc"
Call sbCbo_Llena_New(cboNacionalidad, strSQL, True, True)

strSQL = "select Estado_Civil as 'IdX', Descripcion as 'ItmX' from SYS_ESTADO_CIVIL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoCivil, strSQL, True, True)

'Carga Tipos de Identificacion
vPaso = True
strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, True, True)
vPaso = False



vPaso = True

'Provincias
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)

'Nivel Academico
    strSQL = "select NIVEL_ACADEMICO as Idx, rtrim(Descripcion) as ItmX from RH_NIVEL_ACADEMICO"
    Call sbCbo_Llena_New(cboNivel, strSQL, True, True)

'Nomina
    strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX from RH_NOMINAS_CATALOGO"
    Call sbCbo_Llena_New(cboNomina, strSQL, False, True)

'Divisa
    strSQL = "select COD_DIVISA as Idx, rtrim(Descripcion) as ItmX from vSys_Divisas"
    Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)

'Jornada
    strSQL = "select JORNADA_TIPO as Idx, rtrim(Descripcion) as ItmX from RH_JORNADAS_TIPOS"
    Call sbCbo_Llena_New(cboJornada, strSQL, True, True)

'Contratos
    strSQL = "Select CONTRATO_TIPO as Idx, rtrim(Descripcion) as ItmX from RH_CONTRATOS_TIPOS"
    Call sbCbo_Llena_New(cboContrato, strSQL, True, True)

'Vacaciones
    strSQL = "Select COD_VACA_REGIMEN as Idx, rtrim(Descripcion) as ItmX from RH_VACACIONES_REGIMEN"
    Call sbCbo_Llena_New(cboVacaciones, strSQL, True, True)

'Estados de la Persona
    strSQL = "Select ESTADO_PERSONA as Idx, rtrim(Descripcion) as ItmX from RH_ESTADOS_TIPOS"
    Call sbCbo_Llena_New(cboEstado, strSQL, True, True)

'Bancos Autorizados

    strSQL = "exec spRH_Bancos_Autorizados"
    Call sbCbo_Llena_New(cboBancos, strSQL, True, True)

cboFormaPago.Clear
cboFormaPago.AddItem "TODOS"
cboFormaPago.AddItem "Transferencia"
cboFormaPago.ItemData(cboFormaPago.ListCount - 1) = "TE"
cboFormaPago.AddItem "Cheque"
cboFormaPago.ItemData(cboFormaPago.ListCount - 1) = "CK"

cboFormaPago.Text = "TODOS"



vPaso = False

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCentroCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentroDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_CENTRO,descripcion,desc_Corta from RH_CENTRO_TRABAJO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCentroCod.Text = Trim(gBusquedas.Resultado)
    txtCentroDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub


Private Sub txtCentroDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptCodigo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_CENTRO,descripcion,desc_Corta from RH_CENTRO_TRABAJO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCentroCod.Text = Trim(gBusquedas.Resultado)
    txtCentroDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub


Private Sub txtDeptCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptDesc.SetFocus

If KeyCode = vbKeyF4 Then

    gBusquedas.Columna = "cod_departamento"
    gBusquedas.Orden = "cod_departamento"
    gBusquedas.Consulta = "select cod_departamento,descripcion from RH_Departamentos"
    gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text & "'"
  
   
  
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If
End Sub


Private Sub txtDeptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecCodigo.SetFocus
If KeyCode = vbKeyF4 Then

    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_departamento,descripcion from RH_Departamentos"
    gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text & "'"

  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If

End Sub




Private Sub txtJefeCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtJefeDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Identificacion"
   gBusquedas.Orden = "Identificacion"
   gBusquedas.Consulta = "Select Identificacion,Empleado_ID,Nombre_Completo From Rh_Personas"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   
   txtJefeCod.Text = Trim(gBusquedas.Resultado)
   txtJefeDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtProfesionCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProfesionDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_profesion,descripcion from RH_Profesiones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtProfesionCod.Text = Trim(gBusquedas.Resultado)
    txtProfesionDesc.Text = gBusquedas.Resultado2
  End If
End If

End Sub

Private Sub txtPuestoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPuestoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "COD_PUESTO"
  gBusquedas.Orden = "COD_PUESTO"
  gBusquedas.Consulta = "select COD_PUESTO,descripcion from Rh_Puestos"
  gBusquedas.Filtro = ""
        
  frmBusquedas.Show vbModal
  txtPuestoCod.Text = gBusquedas.Resultado
  txtPuestoDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtPuestoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "COD_PUESTO"
  gBusquedas.Orden = "COD_PUESTO"
  gBusquedas.Consulta = "select COD_PUESTO,descripcion from Rh_Puestos"
  gBusquedas.Filtro = ""
        
  frmBusquedas.Show vbModal
  txtPuestoCod.Text = gBusquedas.Resultado
  txtPuestoDesc.Text = gBusquedas.Resultado2
End If

End Sub


