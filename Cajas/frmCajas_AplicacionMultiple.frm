VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCajas_AplicacionMultiple 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cajas: Aplicación en Lote de Aportes y Abonos"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   11505
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5895
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   11295
      _Version        =   1441793
      _ExtentX        =   19923
      _ExtentY        =   10398
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
      Item(0).Caption =   "Selección"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "lswSel"
      Item(0).Control(1)=   "cboFiltro"
      Item(0).Control(2)=   "cboTipoMovimiento"
      Item(0).Control(3)=   "dtpCorte"
      Item(0).Control(4)=   "lblDate"
      Item(0).Control(5)=   "btnSel(0)"
      Item(0).Control(6)=   "btnSel(1)"
      Item(0).Control(7)=   "lblFechaPago"
      Item(0).Control(8)=   "dtpFechaPago"
      Item(0).Control(9)=   "chkMarcas"
      Item(1).Caption =   "Aplicación"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "gbFondos"
      Item(1).Control(1)=   "lswApl"
      Item(1).Control(2)=   "gbCreditos"
      Item(1).Control(3)=   "cboApl"
      Item(1).Control(4)=   "chkMarcasApl"
      Item(1).Control(5)=   "btnApl(0)"
      Item(1).Control(6)=   "btnApl(1)"
      Begin XtremeSuiteControls.ListView lswApl 
         Height          =   4815
         Left            =   -69880
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   8493
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
      End
      Begin XtremeSuiteControls.ListView lswSel 
         Height          =   4455
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   7858
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
      End
      Begin XtremeSuiteControls.CheckBox chkMarcas 
         Height          =   330
         Left            =   840
         TabIndex        =   33
         Top             =   960
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Marcar ?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.PushButton btnSel 
         Height          =   375
         Index           =   0
         Left            =   8280
         TabIndex        =   28
         Top             =   600
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_AplicacionMultiple.frx":0000
      End
      Begin XtremeSuiteControls.GroupBox gbFondos 
         Height          =   1695
         Left            =   -69880
         TabIndex        =   22
         Top             =   3960
         Visible         =   0   'False
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   2990
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
      End
      Begin XtremeSuiteControls.GroupBox gbCreditos 
         Height          =   1455
         Left            =   -69520
         TabIndex        =   23
         Top             =   3720
         Visible         =   0   'False
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   2566
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
      End
      Begin XtremeSuiteControls.ComboBox cboFiltro 
         Height          =   330
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.ComboBox cboTipoMovimiento 
         Height          =   330
         Left            =   2280
         TabIndex        =   25
         Top             =   600
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   315
         Left            =   4800
         TabIndex        =   27
         Top             =   600
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
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
      Begin XtremeSuiteControls.PushButton btnSel 
         Height          =   375
         Index           =   1
         Left            =   9480
         TabIndex        =   29
         Top             =   600
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Seleccionar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_AplicacionMultiple.frx":0700
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaPago 
         Height          =   315
         Left            =   6360
         TabIndex        =   31
         Top             =   600
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
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
      Begin XtremeSuiteControls.ComboBox cboApl 
         Height          =   330
         Left            =   -69880
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.PushButton btnApl 
         Height          =   375
         Index           =   0
         Left            =   -61720
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_AplicacionMultiple.frx":0FD1
      End
      Begin XtremeSuiteControls.PushButton btnApl 
         Height          =   375
         Index           =   1
         Left            =   -60520
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Eliminar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_AplicacionMultiple.frx":16D1
      End
      Begin XtremeSuiteControls.CheckBox chkMarcasApl 
         Height          =   330
         Left            =   -67480
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Marcar ?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.Label lblFechaPago 
         Height          =   255
         Left            =   6360
         TabIndex        =   32
         Top             =   360
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha Real Pago"
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
      End
      Begin XtremeSuiteControls.Label lblDate 
         Height          =   255
         Left            =   4800
         TabIndex        =   26
         Top             =   360
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha Corte"
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
      End
   End
   Begin XtremeSuiteControls.GroupBox fraFormaPago 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   7200
      Width           =   11175
      _Version        =   1441793
      _ExtentX        =   19711
      _ExtentY        =   2778
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipoDoc 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Top             =   240
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
      Begin XtremeSuiteControls.FlatEdit txtTotalCajas 
         Height          =   315
         Left            =   6000
         TabIndex        =   2
         Top             =   240
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   795
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   6735
         _Version        =   1441793
         _ExtentX        =   11880
         _ExtentY        =   1402
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
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   795
         Index           =   0
         Left            =   8280
         TabIndex        =   4
         Top             =   600
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Pago"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_AplicacionMultiple.frx":1C75
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   795
         Index           =   1
         Left            =   9240
         TabIndex        =   5
         Top             =   600
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_AplicacionMultiple.frx":20D7
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   795
         Index           =   2
         Left            =   10080
         TabIndex        =   6
         Top             =   600
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Cancelar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_AplicacionMultiple.frx":28AF
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   375
         Left            =   9840
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Apllicar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total ..:"
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
         Index           =   4
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas ..:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento ..:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1452
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtTotalAbonos 
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Top             =   6840
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotalAportes 
      Height          =   315
      Left            =   5400
      TabIndex        =   12
      Top             =   6840
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTotalPagar 
      Height          =   315
      Left            =   9240
      TabIndex        =   14
      Top             =   6840
      Width           =   2055
      _Version        =   1441793
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4080
      TabIndex        =   16
      Top             =   240
      Width           =   7095
      _Version        =   1441793
      _ExtentX        =   12515
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   2160
      TabIndex        =   17
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   240
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image imgBanner 
      Height          =   735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total a Pagar..:"
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
      Index           =   5
      Left            =   7560
      TabIndex        =   15
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Aportes Ahorros:"
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
      Index           =   2
      Left            =   3960
      TabIndex        =   13
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Abonos a Créditos:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   6840
      Width           =   1575
   End
End
Attribute VB_Name = "frmCajas_AplicacionMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean, pCharRelleno As String
Dim vCedula As String


Private Sub btnApl_Click(Index As Integer)
If vPaso Then Exit Sub

Select Case Index
    Case 0
        Call sbSelected_Load
    
    Case 1
        Call sbSeleccionar_Undo
End Select

End Sub




Private Sub sbAplicar()

On Error GoTo vError

Me.MousePointer = vbHourglass

'spCajas_AM_Registro_Control(@Caja   varchar(10), @Apertura int  , @Token varchar(100), @Usuario varchar(30)
'                        , @Cedula varchar(20), @Monto dec(18,2), @Divisa varchar(10), @TipoCambio float = 1, @Notas varchar(500)= '')
                        
                        
Dim pAM_Id As Long
                        
txtNotas.Text = fxSysCleanTxtInject(txtNotas.Text)
                        
strSQL = "exec spCajas_AM_Registro_Control '" & ModuloCajas.mCaja & "', " & ModuloCajas.mApertura & ", '" & ModuloCajas.mTiquete _
        & "', '" & glogon.Usuario & "','" & txtCedula.Text & "', " & CCur(txtTotalPagar.Text) & ", '" & ModuloCajas.mDivisa _
        & "', 1, '" & Mid(txtNotas.Text, 1, 500) & "'"
Call OpenRecordSet(rs, strSQL)
    pAM_Id = rs!CAJA_AM_ID
rs.Close
                        
'Create proc spCajas_AM_Procesa(@Cedula varchar(20), @Caja varchar(10), @Apertura int, @Token varchar(100), @Usuario varchar(30)
'        , @TipoDoc varchar(10), @Monto dec(18,2), @Divisa varchar(10), @TipoCambio float = 1, @CajasAM_Id int = 0)
                        
strSQL = "exec spCajas_AM_Procesa  '" & txtCedula.Text & "', '" & ModuloCajas.mCaja & "', " & ModuloCajas.mApertura & ", '" & ModuloCajas.mTiquete _
        & "', '" & glogon.Usuario & "','" & cboTipoDoc.ItemData(cboTipoDoc.ListIndex) & "', " & CCur(txtTotalPagar.Text) & ", '" & ModuloCajas.mDivisa _
        & "', 1, '" & Mid(txtNotas.Text, 1, 500) & "', " & pAM_Id
Call ConectionExecute(strSQL)
                                                                                                
                                                
Me.MousePointer = vbDefault

'Imprime Recibo Multiple
Call sbCaja_Recibo_Multiple(pAM_Id)
                        
                        
'Consulta Nuevo Estado
Call sbConsulta

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
                        
End Sub


Private Sub btnCajas_Click(Index As Integer)
On Error GoTo vError


'Posicion Final
'Call txtTotalPagar_LostFocus

Select Case Index
  Case 2 'Cancelar
     Call sbConsulta
     
  Case 0 'Desgloce
        If Not IsNumeric(txtTotalPagar.Text) Then txtTotalPagar.Text = 0
        If Not ModuloCajas.mConceptoValida Then
           MsgBox "Esta caja no está autorizada para registrar movimientos a este línea de crédito", vbExclamation
           Exit Sub
        End If
        
        ModuloCajas.mTotalAplicar = CCur(txtTotalPagar.Text)
        
        If ModuloCajas.mTotalAplicar = 0 Then
            MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
            Exit Sub
        End If
        
        ModuloCajas.mServicio = "Abonos a Operación de Crédito"
        
        Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)
        
        txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")
        
        
        If txtTotalCajas.Text <> txtTotalPagar.Text Then
           txtTotalCajas.BackColor = vbRed
        Else
           txtTotalCajas.BackColor = vbWhite
        End If

  Case 1   'Aplicar
        If Not IsNumeric(txtTotalPagar.Text) Then Exit Sub
        If CCur(txtTotalPagar.Text) <= 0 Then Exit Sub
        
        If CCur(txtTotalCajas.Text) <> CCur(txtTotalPagar.Text) Then
            txtTotalCajas.BackColor = vbRed
            MsgBox "No se ha indicado el monto correcto en Cajas?", vbInformation
        Else
            txtTotalCajas.BackColor = vbWhite
        
            Call sbAplicar
        End If
        
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCajaInicial()
Dim strSQL As String

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Or ModuloCajas.mUsuario <> glogon.Usuario Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

pCharRelleno = fxCajasParametros("05")

Me.Caption = "Abonos y Aportaciones en Lote   ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

txtTotalCajas.Text = 0
txtNotas.Text = ""
strSQL = "select rTrim(C.tipo_documento) as 'IdX', rtrim(D.Descripcion) as 'itmX'" _
       & " from SIF_DOCUMENTOS D inner join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
       & " Where C.cod_caja =  '" & ModuloCajas.mCaja & "' and D.Tipo_Movimiento in('A','C')" _
       & " order by C.tipo_documento"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)


ModuloCajas.mServicio = "Abonos a Operaciones de Crédito"

'If IsNumeric(ModuloCajas.mRef_01) Then
'    txtOperacion.Text = ModuloCajas.mRef_01
'    vOperacion = txtOperacion.Text
'    Call sbConsulta
'End If



End Sub



Private Sub sbSelected_Load()

On Error GoTo vError


lswApl.ListItems.Clear
lswApl.ColumnHeaders.Clear

txtTotalAbonos.Text = Format(0, "Standard")
txtTotalAportes.Text = Format(0, "Standard")
txtTotalPagar.Text = Format(0, "Standard")

If txtCedula.Text = "" Then Exit Sub

Dim curTAbonos As Currency, curTAportes As Currency

curTAbonos = 0
curTAportes = 0


With lswApl.ColumnHeaders
    .Add , , "No.Operación", 1400
    .Add , , "Linea", 1000, vbCenter
    .Add , , "Saldo", 1400, vbRightJustify
    .Add , , "Tipo", 1200, vbCenter
    
    .Add , , "Abono", 1800, vbRightJustify
    
    .Add , , "Garantía", 2500
    .Add , , "Linea Desc.", 3200
    
    .Add , , "Int.Cor.", 1300, vbRightJustify
    .Add , , "Int.Mor.", 1300, vbRightJustify
    .Add , , "Principal", 1300, vbRightJustify
    .Add , , "Cargos", 1200, vbRightJustify
    .Add , , "Pólizas", 1200, vbRightJustify
End With

Me.MousePointer = vbHourglass

strSQL = "exec spCajas_Crd_Persona_Creditos_En_Lista '" & txtCedula.Text & "', '" & ModuloCajas.mCaja & "', " & ModuloCajas.mApertura _
       & ", '" & ModuloCajas.mTiquete & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswApl.ListItems.Add(, , rs!Id_Solicitud)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = Format(rs!Saldo, "Standard")
     itmX.SubItems(3) = IIf((rs!Tipo_Abono = "C"), "Cancelación", "Pago Cuota")
     itmX.SubItems(4) = Format(rs!Abono, "Standard")
     
     itmX.SubItems(5) = rs!Garantia_Desc
     itmX.SubItems(6) = rs!Linea_Desc
     itmX.SubItems(7) = Format(rs!IntCor, "Standard")
     itmX.SubItems(8) = Format(rs!IntMor, "Standard")
     itmX.SubItems(9) = Format(rs!Amortiza, "Standard")
     itmX.SubItems(10) = Format(rs!Cargos, "Standard")
     itmX.SubItems(11) = Format(rs!Polizas, "Standard")
     itmX.Tag = rs!Creditos_ID
  
     curTAbonos = curTAbonos + rs!Abono
 
 rs.MoveNext
Loop
rs.Close

txtTotalAbonos.Text = Format(curTAbonos, "Standard")
txtTotalAportes.Text = Format(curTAportes, "Standard")
txtTotalPagar.Text = Format(curTAbonos + curTAportes, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbSel_Crd_Load()

On Error GoTo vError


lswSel.ListItems.Clear
lswSel.ColumnHeaders.Clear

With lswSel.ColumnHeaders
    .Add , , "No.Operación", 1400
    .Add , , "Linea", 1000, vbCenter
    .Add , , "Saldo", 1400, vbRightJustify
    
    .Add , , "Abono", 1800, vbRightJustify
    .Add , , "Ult.Pago", 1400, vbCenter
    
    
    .Add , , "Garantía", 2500
    .Add , , "Linea Desc.", 3200
    
    .Add , , "Int.Cor.", 1300, vbRightJustify
    .Add , , "Int.Mor.", 1300, vbRightJustify
    .Add , , "Principal", 1300, vbRightJustify
    .Add , , "Cargos", 1200, vbRightJustify
    .Add , , "Pólizas", 1200, vbRightJustify
End With


If txtCedula.Text = "" Then Exit Sub


Me.MousePointer = vbHourglass

Dim fPago As String


If dtpFechaPago.Visible Then
    fPago = "'" & Format(dtpFechaPago.Value, "yyyy-mm-dd") & "'"
Else
    fPago = "Null"
End If
strSQL = "exec spCajas_Crd_Persona_Creditos_Pendientes_Lista '" & txtCedula.Text & "', '" & ModuloCajas.mCaja & "', " & ModuloCajas.mApertura _
       & ", '" & ModuloCajas.mTiquete & "', '" & Format(dtpCorte.Value, "yyyy-mm-dd") _
       & " 23:59:59', '" & cboTipoMovimiento.ItemData(cboTipoMovimiento.ListIndex) & "', " & fPago
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswSel.ListItems.Add(, , rs!Id_Solicitud)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = Format(rs!Saldo, "Standard")
     itmX.SubItems(3) = Format(rs!Compromiso, "Standard")
     itmX.SubItems(4) = rs!CtaFechaUltCorte
     itmX.SubItems(5) = rs!GarantiaX
     itmX.SubItems(6) = rs!Descripcion
     itmX.SubItems(7) = Format(rs!IntC, "Standard")
     itmX.SubItems(8) = Format(rs!IntM, "Standard")
     itmX.SubItems(9) = Format(rs!Principal, "Standard")
     itmX.SubItems(10) = Format(rs!Cargos, "Standard")
     itmX.SubItems(11) = Format(rs!Polizas, "Standard")
 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbSel_Fnd_Load()

lswSel.ListItems.Clear

With lswSel.ColumnHeaders
    .Add , , "No.Contrato", 2000
    .Add , , "Plan Id", 900, vbCenter
    .Add , , "Mensualidad", 1800, vbRightJustify
    .Add , , "Mnt.Acumulado", 1800, vbRightJustify
    .Add , , "Aporte ?", 1800, vbRightJustify
    .Add , , "Divisa", 1000, vbCenter
    .Add , , "Plan Desc.", 3300
    .Add , , "Operadora", 3150
End With

End Sub

Private Sub sbSeleccionar()
Dim i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = ""

With lswSel.ListItems
  Select Case cboFiltro.Text
    Case "Créditos"
        For i = 1 To .Count
            If .Item(i).Checked Then
                            
                strSQL = strSQL & Space(10) & "exec spCajas_AM_Creditos_Add '" & ModuloCajas.mCaja & "', " & ModuloCajas.mApertura _
                       & ", '" & ModuloCajas.mTiquete & "', '" & txtCedula.Text & "', " & .Item(i).Text & ", '" & .Item(i).SubItems(1) _
                       & "', '" & IIf((cboTipoMovimiento.ItemData(cboTipoMovimiento.ListIndex) = "Can"), "C", "T") _
                       & "', '" & Format(dtpCorte.Value, "yyyy-mm-dd") & "', " & CCur(.Item(i).SubItems(3)) & ", " & CCur(.Item(i).SubItems(3)) _
                       & ", " & CCur(.Item(i).SubItems(2)) & ", " & CCur(.Item(i).SubItems(7)) & ", " & CCur(.Item(i).SubItems(8)) _
                       & ", " & CCur(.Item(i).SubItems(9)) & ", " & CCur(.Item(i).SubItems(10)) & ", " & CCur(.Item(i).SubItems(11))
            
                If Len(strSQL) > 20000 Then
                   Call ConectionExecute(strSQL)
                   strSQL = ""
                End If
                
            End If
        Next i
        'Lote Final
        If Len(strSQL) > 0 Then
           Call ConectionExecute(strSQL)
           strSQL = ""
        End If
        
        
    Case "Fondos"
  End Select


End With

Me.MousePointer = vbDefault

'Actualiza Lista
Call btnSel_Click(0)

'Cargar las Seleccionadas
Call sbSelected_Load

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub


Private Sub sbSeleccionar_Undo()
Dim i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = ""

With lswApl.ListItems
        
    For i = 1 To .Count
        If .Item(i).Checked Then
                        
            strSQL = strSQL & Space(10) & "exec spCajas_AM_Selected_Del '" & Mid(cboApl.Text, 1, 1) & "', " & .Item(i).Tag
            If Len(strSQL) > 20000 Then
               Call ConectionExecute(strSQL)
               strSQL = ""
            End If
            
        End If
    Next i
            
    
    'Lote Final
    If Len(strSQL) > 0 Then
       Call ConectionExecute(strSQL)
       strSQL = ""
    End If


End With

Me.MousePointer = vbDefault

'Actualiza Lista
Call btnSel_Click(0)

'Cargar las Seleccionadas
Call sbSelected_Load

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub




Private Sub btnSel_Click(Index As Integer)

Select Case Index
    Case 0
    
        Select Case cboTipoMovimiento.ItemData(cboTipoMovimiento.ListIndex)
            Case "Cor" 'Cuotas al Corte
                Call sbSel_Crd_Load
            
            Case "Can" 'Fecha de  Cancelación
                Call sbSel_Crd_Load
            
            Case "Apo", "Men"
                Call sbSel_Fnd_Load
        End Select
    
    
    Case 1
        Call sbSeleccionar
End Select

End Sub

Private Sub cboApl_Click()
If vPaso Then Exit Sub

Call sbSelected_Load

End Sub

Private Sub cboFiltro_Click()

If vPaso Then Exit Sub

cboTipoMovimiento.Clear


lswSel.ListItems.Clear
lswSel.ColumnHeaders.Clear

If cboFiltro.Text = "Créditos" Then
    cboTipoMovimiento.AddItem "Cuotas al Corte"
    cboTipoMovimiento.ItemData(cboTipoMovimiento.ListCount - 1) = "Cor"
    cboTipoMovimiento.AddItem "Cancelación"
    cboTipoMovimiento.ItemData(cboTipoMovimiento.ListCount - 1) = "Can"
    cboTipoMovimiento.Text = "Cuotas al Corte"

Else
    'Fondos
    cboTipoMovimiento.AddItem "Aporte"
    cboTipoMovimiento.ItemData(cboTipoMovimiento.ListCount - 1) = "Apo"
    cboTipoMovimiento.AddItem "Mensualidad"
    cboTipoMovimiento.ItemData(cboTipoMovimiento.ListCount - 1) = "Men"
    cboTipoMovimiento.Text = "Aporte"

End If

End Sub

Private Sub cboTipoMovimiento_Click()
If vPaso Then Exit Sub

dtpFechaPago.Visible = False

Select Case cboTipoMovimiento.ItemData(cboTipoMovimiento.ListIndex)
    Case "Cor" 'Cuotas al Corte
        lblDate.Caption = "Fec.Venc.Ctas."
        dtpCorte.MaxDate = "31/12/9999"
        dtpCorte.Value = fxCorteMesActual
        dtpFechaPago.Visible = True
        
        
        Call sbSel_Crd_Load
    
    Case "Can" 'Fecha de  Cancelación
        lblDate.Caption = "Fec.Cancelación"
        dtpCorte.Value = fxFechaServidor
        dtpCorte.MaxDate = dtpCorte.Value
    
        dtpFechaPago.Visible = False
    
        Call sbSel_Crd_Load
    
    Case "Apo", "Men"
        Call sbSel_Fnd_Load
End Select

lblFechaPago.Visible = dtpFechaPago.Visible

End Sub

Private Sub chkMarcas_Click()
Dim i As Integer


With lswSel.ListItems
 For i = 1 To .Count
    .Item(i).Checked = chkMarcas.Value
 Next i
End With

End Sub

Private Sub chkMarcasApl_Click()
Dim i As Integer


With lswApl.ListItems
 For i = 1 To .Count
    .Item(i).Checked = chkMarcasApl.Value
 Next i
End With

End Sub

Private Sub dtpCorte_Change()
If vPaso Then Exit Sub





End Sub

Private Sub Form_Activate()
 vModulo = 5
End Sub

Private Function fxCorteMesActual() As Date
Dim pFecha As Date


pFecha = fxFechaServidor

pFecha = Year(pFecha) & "/" & Format(Month(pFecha), "00") & "/01"
pFecha = DateAdd("d", -1, DateAdd("m", 1, pFecha))

fxCorteMesActual = pFecha

End Function

Private Sub Form_Load()
Dim iDias As Long, vFechaHoy As Date

 vModulo = 5
 
 vFechaHoy = fxFechaServidor
 iDias = fxCrdParametro("32")
 
 vModulo = 5
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True

cboFiltro.Clear
cboFiltro.AddItem "Créditos"
cboFiltro.AddItem "Fondos"

cboApl.Clear
cboApl.AddItem "Créditos"
cboApl.AddItem "Fondos"


vPaso = False

Call sbLimpiaDatos

dtpFechaPago.Value = vFechaHoy
dtpFechaPago.MinDate = DateAdd("d", (iDias * -1), dtpFechaPago.Value)
dtpFechaPago.MaxDate = dtpFechaPago.Value
 
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 btnCajas.Item(1).Enabled = btnAplicar.Enabled
 
 cboFiltro.Text = "Créditos"
 cboApl.Text = "Créditos"
 
End Sub

Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbCajaInicial

If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   Unload Me
   Exit Sub
End If

End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Col1Name = "Cédula Colilla"
    gBusquedas.Col2Name = "Cédula Real"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    txtCedula.Text = Trim(gBusquedas.Resultado)
    txtNombre.Text = Trim(gBusquedas.Resultado3)
    gBusquedas.Consulta = ""
    gBusquedas.Columna = ""
    gBusquedas.Orden = ""
    gBusquedas.Resultado = ""
    If Trim(txtCedula) <> "" Then
        Call sbConsulta
    End If

End If

End Sub




Private Sub sbConsulta()


Me.MousePointer = vbHourglass

Call sbLimpiaDatos
 

strSQL = "select cedula, nombre from socios where cedula = '" & txtCedula.Text & "'"
       
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    txtCedula = Trim(rs!Cedula)
    txtNombre = Trim(rs!Nombre)

    ModuloCajas.mClienteId = Trim(rs!Cedula)
    ModuloCajas.mCliente = Trim(rs!Nombre)
    ModuloCajas.mTiquete = Trim(rs!Cedula) & "." & Format(Now, "yyyymmdd") & Format(Time, "HHmmss")
    
    ModuloCajas.mConceptoValida = True  ' IIf((rs!Caja_Valida_Concepto > 0), True, False)
    
    ModuloCajas.mTotalDetallado = 0
    txtTotalCajas.Text = 0
    
    
    
    strSQL = " select rtrim(COD_DIVISA) as 'Divisa'  from vSys_Divisas where DIVISA_LOCAL = 1"
    Call OpenRecordSet(rs, strSQL)
    
    ModuloCajas.mDivisa = RTrim(rs!Divisa)
    
    
Else
 
 MsgBox "No se Encontró operación para abonos,puede que se encuentre cancelada ", vbInformation

End If
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub sbLimpiaDatos()
 
 tcMain.Item(0).Selected = True
 
 txtTotalCajas.Text = 0
 txtTotalAbonos.Text = 0
 txtTotalAportes.Text = 0
 
 lswSel.ListItems.Clear
 lswApl.ListItems.Clear
 
End Sub

