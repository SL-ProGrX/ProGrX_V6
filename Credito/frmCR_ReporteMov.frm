VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_ReportesMov 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Movimientos a  Créditos"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10020
   Icon            =   "frmCR_ReporteMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   10020
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5532
      Left            =   120
      TabIndex        =   42
      Top             =   1560
      Width           =   4092
      _Version        =   1441793
      _ExtentX        =   7218
      _ExtentY        =   9758
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
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl tcFiltros 
      Height          =   3735
      Left            =   4320
      TabIndex        =   11
      Top             =   3360
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
      _ExtentY        =   6588
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
      Item(0).ControlCount=   15
      Item(0).Control(0)=   "Label1(14)"
      Item(0).Control(1)=   "Label1(15)"
      Item(0).Control(2)=   "Label1(18)"
      Item(0).Control(3)=   "cboDestino"
      Item(0).Control(4)=   "cboRecurso"
      Item(0).Control(5)=   "cboInstitucion"
      Item(0).Control(6)=   "chkLineas"
      Item(0).Control(7)=   "txtCodigo"
      Item(0).Control(8)=   "txtDescripcion"
      Item(0).Control(9)=   "Label1(7)"
      Item(0).Control(10)=   "chkCanceladas"
      Item(0).Control(11)=   "Label1(39)"
      Item(0).Control(12)=   "Label1(38)"
      Item(0).Control(13)=   "cboEspecial"
      Item(0).Control(14)=   "Label1(9)"
      Item(1).Caption =   "Add No 1"
      Item(1).ControlCount=   18
      Item(1).Control(0)=   "cboSigno(0)"
      Item(1).Control(1)=   "cboSigno(1)"
      Item(1).Control(2)=   "txtUltMov"
      Item(1).Control(3)=   "txtPrideduc"
      Item(1).Control(4)=   "chkPriDeduc"
      Item(1).Control(5)=   "chkUltMov"
      Item(1).Control(6)=   "Label1(34)"
      Item(1).Control(7)=   "Label1(36)"
      Item(1).Control(8)=   "cboOficina"
      Item(1).Control(9)=   "Label1(26)"
      Item(1).Control(10)=   "cboGarantia"
      Item(1).Control(11)=   "Label1(16)"
      Item(1).Control(12)=   "Label1(8)"
      Item(1).Control(13)=   "Label1(11)"
      Item(1).Control(14)=   "txtIdentificacion"
      Item(1).Control(15)=   "txtDocumento"
      Item(1).Control(16)=   "cboDivisa"
      Item(1).Control(17)=   "Label1(10)"
      Item(2).Caption =   "Add No 2"
      Item(2).ControlCount=   6
      Item(2).Control(0)=   "cboAseguradoras"
      Item(2).Control(1)=   "cboPolizas"
      Item(2).Control(2)=   "Label1(3)"
      Item(2).Control(3)=   "Label1(12)"
      Item(2).Control(4)=   "cboCargos"
      Item(2).Control(5)=   "Label1(13)"
      Begin XtremeSuiteControls.ComboBox cboDestino 
         Height          =   330
         Left            =   1080
         TabIndex        =   12
         Top             =   1320
         Width           =   4455
         _Version        =   1441793
         _ExtentX        =   7858
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
      Begin XtremeSuiteControls.ComboBox cboRecurso 
         Height          =   330
         Left            =   1080
         TabIndex        =   13
         Top             =   1680
         Width           =   4455
         _Version        =   1441793
         _ExtentX        =   7858
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
      Begin XtremeSuiteControls.ComboBox cboInstitucion 
         Height          =   330
         Left            =   1080
         TabIndex        =   14
         Top             =   2040
         Width           =   4455
         _Version        =   1441793
         _ExtentX        =   7858
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
      Begin XtremeSuiteControls.CheckBox chkLineas 
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   480
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas   "
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
         TextAlignment   =   1
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   330
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   330
         Left            =   1080
         TabIndex        =   17
         Top             =   840
         Width           =   4455
         _Version        =   1441793
         _ExtentX        =   7858
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
      Begin XtremeSuiteControls.ComboBox cboSigno 
         Height          =   315
         Index           =   0
         Left            =   -67360
         TabIndex        =   23
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.ComboBox cboSigno 
         Height          =   315
         Index           =   1
         Left            =   -67360
         TabIndex        =   24
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
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
      Begin XtremeSuiteControls.FlatEdit txtUltMov 
         Height          =   330
         Left            =   -66520
         TabIndex        =   25
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.FlatEdit txtPrideduc 
         Height          =   330
         Left            =   -66520
         TabIndex        =   26
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.CheckBox chkPriDeduc 
         Height          =   255
         Left            =   -65440
         TabIndex        =   27
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkUltMov 
         Height          =   255
         Left            =   -65440
         TabIndex        =   28
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   330
         Left            =   -68680
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin XtremeSuiteControls.ComboBox cboGarantia 
         Height          =   330
         Left            =   -68680
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin XtremeSuiteControls.CheckBox chkCanceladas 
         Height          =   255
         Left            =   1680
         TabIndex        =   36
         Top             =   3240
         Width           =   3735
         _Version        =   1441793
         _ExtentX        =   6583
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Operaciones Canceladas   "
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
         TextAlignment   =   1
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
         Height          =   315
         Left            =   -68680
         TabIndex        =   40
         Top             =   2160
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   315
         Left            =   -68680
         TabIndex        =   41
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin XtremeSuiteControls.ComboBox cboEspecial 
         Height          =   330
         Left            =   1080
         TabIndex        =   45
         Top             =   2760
         Width           =   4455
         _Version        =   1441793
         _ExtentX        =   7858
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
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   330
         Left            =   -68680
         TabIndex        =   47
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin XtremeSuiteControls.ComboBox cboAseguradoras 
         Height          =   330
         Left            =   -68680
         TabIndex        =   51
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin XtremeSuiteControls.ComboBox cboPolizas 
         Height          =   330
         Left            =   -68680
         TabIndex        =   52
         Top             =   1680
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin XtremeSuiteControls.ComboBox cboCargos 
         Height          =   330
         Left            =   -68680
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin VB.Label Label1 
         Caption         =   "Tipos de Cargos"
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
         Index           =   13
         Left            =   -69880
         TabIndex        =   56
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Aseguradoras"
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
         Index           =   12
         Left            =   -69880
         TabIndex        =   54
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tipos de Pólizas"
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
         Index           =   3
         Left            =   -69880
         TabIndex        =   53
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
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
         Height          =   252
         Index           =   10
         Left            =   -69880
         TabIndex        =   48
         Top             =   1320
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Especial"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   9
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cédula"
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
         Index           =   11
         Left            =   -69880
         TabIndex        =   39
         Top             =   2160
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
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
         Left            =   -69880
         TabIndex        =   38
         Top             =   1800
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   7
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Garantía"
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
         Left            =   -69880
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Oficina"
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
         Index           =   26
         Left            =   -69880
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Ult.Mov."
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
         Index           =   36
         Left            =   -68680
         TabIndex        =   30
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Primer Deduc."
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
         Index           =   34
         Left            =   -68680
         TabIndex        =   29
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Deductora"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   39
         Left            =   -1.26608e5
         TabIndex        =   22
         Top             =   3240
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   38
         Left            =   -1.26608e5
         TabIndex        =   21
         Top             =   2160
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Recurso"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Institución"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   14
         Left            =   120
         TabIndex        =   18
         Top             =   525
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   255
      Left            =   9600
      TabIndex        =   2
      ToolTipText     =   "Todas las Fechas"
      Top             =   2040
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   444
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Alignment       =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   8160
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Top             =   1680
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
   Begin XtremeSuiteControls.ComboBox cboTransaccion 
      Height          =   315
      Left            =   5640
      TabIndex        =   10
      Top             =   2760
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
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   495
      Left            =   6720
      TabIndex        =   43
      Top             =   7320
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Reporte"
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
      Picture         =   "frmCR_ReporteMov.frx":000C
   End
   Begin XtremeSuiteControls.PushButton btnCubo 
      Height          =   495
      Left            =   8280
      TabIndex        =   44
      Top             =   7320
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cubo"
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
      Picture         =   "frmCR_ReporteMov.frx":0713
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   4320
      TabIndex        =   50
      Top             =   1080
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Filtros:"
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
   Begin XtremeShortcutBar.ShortcutCaption lblReporte 
      Height          =   372
      Left            =   0
      TabIndex        =   49
      Top             =   1080
      Width           =   4332
      _Version        =   1441793
      _ExtentX        =   7641
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "...."
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transacción"
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
      Left            =   4440
      TabIndex        =   37
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   6
      Left            =   7560
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   5
      Left            =   5640
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
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
      Index           =   4
      Left            =   4440
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Reporte"
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
      Left            =   4440
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
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
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   4320
      X2              =   7080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   120
      X2              =   3720
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos a Operaciones de crédito"
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
      Height          =   492
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   7572
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmCR_ReportesMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mModoSif As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub sbInicializa()

Me.MousePointer = vbHourglass

tcFiltros.Item(0).Selected = True


strSQL = "select cod_institucion as 'IdX',rtrim(descripcion) as 'ItmX' from instituciones order by descripcion"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)

strSQL = "select rtrim(Garantia) as 'IdX' , rtrim(descripcion) as Itmx" _
       & " from crd_garantia_tipos order by descripcion"
Call sbCbo_Llena_New(cboGarantia, strSQL, True, True)


strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as Itmx" _
       & " from SIF_Oficinas order by descripcion"
Call sbCbo_Llena_New(cboOficina, strSQL, True, True)


strSQL = "select COD_DIVISA AS 'IdX', DESCRIPCION as 'ItmX'" _
       & " From vSys_Divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)


strSQL = "select rtrim(Tipo_Documento) as  'IdX' , rtrim(Descripcion) + space(5) + '[' + rtrim(Tipo_Documento) + ']' as ItmX" _
       & " from sif_documentos" _
       & " Where Tipo_Documento in('FRM','ND','NC','RE','LIQ','RLIQ','PLA','AFR','CBR','TRA','REA','CAJA','CAJARE','CA')" _
       & " order by Descripcion"
Call sbCbo_Llena_New(cboTransaccion, strSQL, True, True)

Call chkFechas_Click
Call chkLineas_Click

strSQL = "select COD_CARGO as 'IdX',rtrim(descripcion) as 'ItmX' from vCrd_Cargos_Unificados_Lista order by descripcion"
Call sbCbo_Llena_New(cboCargos, strSQL, True, True)

strSQL = "select COD_ASEGURADORA as 'IdX',rtrim(NOMBRE) as 'ItmX' from CRD_POLIZAS_ASEGURADORAS order by NOMBRE"
Call sbCbo_Llena_New(cboAseguradoras, strSQL, True, True)

strSQL = "select COD_POLIZA as 'IdX',rtrim(descripcion) as 'ItmX' from CRD_CATALOGO_POLIZAS order by descripcion"
Call sbCbo_Llena_New(cboPolizas, strSQL, True, True)



Me.MousePointer = vbDefault

End Sub



Private Sub btnCubo_Click()
lblStatus.Visible = True
Call sbCubo
End Sub

Private Sub btnReporte_Click()
lblStatus.Visible = False
Call sbReporteCRD
End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub


Private Sub chkLineas_Click()

If chkLineas.Value = vbChecked Then
  
  txtCodigo.Enabled = False
  
  strSQL = "select cod_grupo as 'IdX',  rtrim(descripcion) as 'ItmX'" _
         & " from  catalogo_grupos order by Descripcion"
  Call sbCbo_Llena_New(cboRecurso, strSQL, True, True)
  
  strSQL = "select cod_destino as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  catalogo_destinos order by descripcion"
  Call sbCbo_Llena_New(cboDestino, strSQL, True, True)
  
Else
  txtCodigo.Enabled = True

  strSQL = "select (R.cod_grupo) as 'IdX', rtrim(R.descripcion) as 'ItmX'" _
         & " from catalogo_grupos R inner join catalogo_AsignaGrp A on R.cod_grupo = A.cod_grupo" _
         & " where A.codigo = '" & txtCodigo & "' order by R.descripcion"
  Call sbCbo_Llena_New(cboRecurso, strSQL, True, True)
  
  strSQL = "select (R.cod_destino) as 'IdX', rtrim(R.descripcion) as 'ItmX'" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtCodigo & "' order by R.descripcion"
  Call sbCbo_Llena_New(cboDestino, strSQL, True, True)

End If

End Sub


Private Sub chkPriDeduc_Click()
If chkPriDeduc.Value = vbChecked Then
   txtPrideduc.Enabled = False
Else
   txtPrideduc.Enabled = True
End If
End Sub

Private Sub chkUltMov_Click()
If chkUltMov.Value = vbChecked Then
   txtUltMov.Enabled = False
Else
   txtUltMov.Enabled = True
End If
End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()


vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture


lblReporte.Tag = ""
lblReporte.Caption = ">>> Seleccione Un Reporte <<<"


cboTipo.Clear
cboTipo.AddItem "Detallado"
cboTipo.AddItem "Resumen"
cboTipo.Text = "Detallado"


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
chkFechas.Value = vbUnchecked
chkLineas.Value = vbChecked

cboEspecial.Clear
cboEspecial.AddItem "TODOS"
cboEspecial.AddItem "Cartera Interna"
cboEspecial.AddItem "Cartera Administrada"
cboEspecial.AddItem "Recaudos & Retenciones"
cboEspecial.Text = "TODOS"


With lsw.ColumnHeaders
    .Clear
    .Add , , "Reporte", 3600
End With

'Llena lsw
lsw.ListItems.Clear
lsw.ListItems.Add , "x00", "Movimientos General"
lsw.ListItems.Add , "x12", "Movimientos por Tipo Documento"
lsw.ListItems.Add , "x13", "Movimientos por Tipo, Línea"
lsw.ListItems.Add , "x14", "Movimientos por Tipo, Garantía"
lsw.ListItems.Add , "x01", "Movimientos por Línea"
lsw.ListItems.Add , "x02", "Movimientos por Línea, Destino"
lsw.ListItems.Add , "x03", "Movimientos por Destino, Tipo"
lsw.ListItems.Add , "x04", "Movimientos por Linea, Tipo"
lsw.ListItems.Add , "x05", "Movimientos por Linea, Destino, Tipo"
lsw.ListItems.Add , "x06", "Movimientos por Usuarios, Linea, Destino"
lsw.ListItems.Add , "x07", "Movimientos por Institución"
lsw.ListItems.Add , "x08", "Movimientos por Institución (Estadística)"
lsw.ListItems.Add , "x09", "Movimientos por Oficina"
lsw.ListItems.Add , "x10", "Movimientos por Oficina (Estadística)"
lsw.ListItems.Add , "x11", "Movimientos por Cliente"

lsw.ListItems.Add , "x20", "Pólizas aplicadas"
lsw.ListItems.Add , "x21", "Cargos aplicados"


cboSigno(0).Text = "="
cboSigno(1).Text = "="

txtPrideduc.Text = GLOBALES.glngFechaCR
txtUltMov.Text = GLOBALES.glngFechaCR

chkPriDeduc_Click
chkUltMov_Click


Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbReporteCRD()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String, vTemp As String

On Error GoTo vError

If lblReporte.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass


vTitulo = UCase(lblReporte.Caption & " : " & cboTipo.Text)
vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Créditos"
 
 .Connect = glogon.ConectRPT
  
 If chkFechas.Value = vbUnchecked Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        vSubTitulo = "Fechas: " & Format(dtpInicio.Value, "dd-mm-yyyy") & " al " & Format(dtpCorte.Value, "dd-mm-yyyy")
       
        strSQL = strSQL & "{vCRDsReportesMov.Fecha}" _
               & " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" _
               & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
 Else
   vSubTitulo = "Historico"
 End If
 
 
 If cboDivisa.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.COD_DIVISA} = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
 End If
 vSubTitulo = vSubTitulo & " ¦ Divisa: " & cboDivisa.Text


 If cboEspecial.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   
   Select Case cboEspecial.Text
        Case "Cartera Interna"
            strSQL = strSQL & "{vCRDsReportesMov.Linea_Interna} = 1"
        Case "Cartera Administrada"
            strSQL = strSQL & "{vCRDsReportesMov.Linea_Interna} = 0"
        Case "Recaudos & Retenciones"
            strSQL = strSQL & "({vCRDsReportesMov.CAT_RETENCION} = 'S' OR {vCRDsReportesMov.CAT_POLIZA} = 'S')"
   End Select
 End If
 vSubTitulo = vSubTitulo & " ¦ Listado: " & cboEspecial.Text
 
 
 If cboOficina.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.cod_oficina_R} = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
 End If
 vSubTitulo = vSubTitulo & " ¦ Oficina: " & cboOficina.Text
 
 
 If chkLineas.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.Codigo} = '" & Trim(txtCodigo) & "'"
   vFiltro = vFiltro & "¦ línea: " & UCase(txtCodigo)
 Else
   vFiltro = vFiltro & "¦ Todas la Líneas"
 End If
 
 If cboRecurso.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.cod_grupo} = '" & cboRecurso.ItemData(cboRecurso.ListIndex) & "'"
 End If
 vFiltro = vFiltro & "¦ Recurso: " & cboRecurso.Text
 
 If cboDestino.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.cod_destino} = '" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
 End If
 vFiltro = vFiltro & "¦ Destino: " & cboDestino.Text

 If cboInstitucion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.cod_institucion} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ""
 End If
 vFiltro = vFiltro & "¦ Empresa: " & cboInstitucion.Text
 
 'No. Documento
 If Trim(txtDocumento.Text) <> "" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.Ncon} = '" & txtDocumento.Text & "'"
    
   vFiltro = vFiltro & "¦ No.Doc. " & txtDocumento.Text
 
 End If
 
 
 'No. Cédula (Identificacion)
 If Trim(txtIdentificacion.Text) <> "" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.Cedula} = '" & txtIdentificacion.Text & "'"
    
   vFiltro = vFiltro & "¦ No.Id. " & txtIdentificacion.Text
 End If
 
 
 'Primer Deducción
 If chkPriDeduc.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.PriDeduc} " & cboSigno(0).Text & txtPrideduc.Text
    
   vFiltro = vFiltro & "¦ Pri.Deduc. " & cboSigno(0).Text & " " & txtPrideduc.Text
 End If
 
 'Ultimo Movimiento
 If chkUltMov.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.FecUlt} " & cboSigno(1).Text & txtUltMov.Text
    
   vFiltro = vFiltro & "¦ Ult.Mov. " & cboSigno(1).Text & " " & txtUltMov.Text
 End If
 
 'Listar Operaciones Canceladas con el Movimiento
 If chkCanceladas.Value = vbChecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.SALDO} <= 1 "
    
   vFiltro = vFiltro & "¦ Ops. Canceladas "
 End If
 
 
 
 If cboGarantia.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.Garantia} = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
 End If
 vFiltro = vFiltro & "¦ Garantia: " & cboGarantia.Text
 
 If cboTransaccion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDsReportesMov.Tcon} = '" & cboTransaccion.ItemData(cboTransaccion.ListIndex) & "'"
 End If
 vFiltro = vFiltro & "¦ Transacción: " & cboTransaccion.Text
 
 
 .Formulas(0) = "fxFecha='Fecha: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='Usuario: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='" & vTitulo & "'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(5) = "fxFiltro='" & Mid(vFiltro, 1, 250) & "'"
 
 Select Case lblReporte.Tag
     Case "x00" 'Movimientos General
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovGeneral.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovGeneralRsm.rpt")
         End If
 
      Case "x01" 'Movimientos x Línea
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovLinea.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovLineaRsm.rpt")
         End If
     
     
     Case "x02" 'Movimientos x Línea x Destino
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovLineaDestino.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovLineaDestinoRsm.rpt")
         End If
     Case "x03" 'Movimientos x Destino x Tipo
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovDestinoTipo.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovDestinoTipoRsm.rpt")
         End If
     Case "x04" 'Movimientos x Linea x Tipo
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovLineaTipo.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovLineaTipoRsm.rpt")
         End If
     Case "x05" 'Movimientos x Linea x Destino x Tipo
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovLineaDestinoTipo.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovLineaDestinoTipoRsm.rpt")
         End If
     Case "x06" 'Movimientos x Usuarios x Linea x Destino
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovUsuarioLineaDestino.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovUsuarioLineaDestinoRsm.rpt")
         End If
 
     Case "x07" 'Movimientos x Institución
          If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovInstitucion.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovInstitucionRsm.rpt")
         End If
     Case "x08" 'Movimientos x Institución (Estadística)
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovInstitucionEstadistica.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovInstitucionEstadisticaRsm.rpt")
         End If
     Case "x09" 'Movimientos x Oficina
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovOficina.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovOficinaRsm.rpt")
         End If
     Case "x10" 'Movimientos x Oficina (Estadística)
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovOficinaEstadistica.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovOficinaEstadisticaRsm.rpt")
         End If
 
     Case "x11" 'Movimientos por Cliente
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovCliente.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovClienteRsm.rpt")
         End If
 
      Case "x12" 'Movimientos x Tipo de Documento
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovTipo.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovTipoRsm.rpt")
         End If
     
      Case "x13" 'Movimientos x Tipo x Línea
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovTipoLinea.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovTipoLineaRsm.rpt")
         End If
     
      Case "x14" 'Movimientos x Tipo x Garantía
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovTipoGarantia.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovTipoGarantiaRsm.rpt")
         End If
 
 
 
 
 
 
      Case "x20" 'Detalle de Pólizas
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovPolizas.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovPolizasRsm.rpt")
         End If
 
        If cboAseguradoras.Text <> "TODOS" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{vCrd_Polizas_Aplicadas.COD_ASEGURADORA} = '" & cboAseguradoras.ItemData(cboAseguradoras.ListIndex) & "'"
        End If
        vFiltro = vFiltro & "¦ Aseguradora: " & cboAseguradoras.Text
 
        If cboPolizas.Text <> "TODOS" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{vCrd_Polizas_Aplicadas.COD_POLIZA} = '" & cboPolizas.ItemData(cboPolizas.ListIndex) & "'"
        End If
        vFiltro = vFiltro & "¦ Tipo Póliza: " & cboTransaccion.Text
 
 
      Case "x21" 'Detalle de Cargos
         If Mid(cboTipo.Text, 1, 1) = "D" Then
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovCargos.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_MovCargosRsm.rpt")
         End If
 
        If cboCargos.Text <> "TODOS" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{vCrd_Cargos_Aplicados.COD_CARGO_UNI} = '" & cboCargos.ItemData(cboCargos.ListIndex) & "'"
        End If
        vFiltro = vFiltro & "¦ Tipo Cargo: " & cboCargos.Text
 
 
 End Select
 
 .SelectionFormula = strSQL

 .PrintReport

End With

Me.MousePointer = vbDefault

Call Bitacora("Imprime", lblReporte.Caption)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCubo()
Dim vFechaInicio As Date, vFechaCorte As Date
Dim vMensaje As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lblStatus.Caption = "Procesando Información Espere!....Este proceso puede durar varios minutos."
lblStatus.Refresh

vMensaje = "Credito_Movimientos"

If chkFechas.Value = vbChecked Then
  vFechaInicio = "1900/01/01"
  vFechaCorte = fxFechaServidor
Else
  vFechaInicio = dtpInicio.Value
  vFechaCorte = dtpCorte.Value
End If

strSQL = "exec spCrdMovAnalisisCubo '" & Format(vFechaInicio, "yyyy/mm/dd") & "','" & Format(dtpCorte, "yyyy/mm/dd") & "'"
Call ConectionExecute(strSQL)

lblStatus.Caption = "Proceso Concluido con éxito, la información puede ser utilizada desde la base de datos de análisis, cubo: " & vMensaje

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

lblReporte.Tag = Item.Key
lblReporte.Caption = Item.Text

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then cboDestino.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select Codigo, Descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCodigo_LostFocus()
 If Len(Trim(txtCodigo)) > 0 Then txtDescripcion.Text = fxDescribeCodigo(Trim(txtCodigo.Text))
 Call chkLineas_Click
End Sub

