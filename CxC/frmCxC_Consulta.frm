VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCxC_Consulta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Consulta de CxC"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10995
   Icon            =   "frmCxC_Consulta.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   10995
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   7200
      Top             =   480
   End
   Begin XtremeSuiteControls.TabControl ssTab 
      Height          =   5772
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   10692
      _Version        =   1441793
      _ExtentX        =   18860
      _ExtentY        =   10181
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
      ItemCount       =   4
      Item(0).Caption =   "Operaciones"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "vgCxC"
      Item(0).Control(1)=   "txtTotalSaldo"
      Item(0).Control(2)=   "txtTotalCuota"
      Item(0).Control(3)=   "Label1(0)"
      Item(0).Control(4)=   "Label1(1)"
      Item(0).Control(5)=   "Label1(2)"
      Item(0).Control(6)=   "txtTotalMonto"
      Item(1).Caption =   "Facturas"
      Item(1).ControlCount=   14
      Item(1).Control(0)=   "feFactura"
      Item(1).Control(1)=   "Label2(4)"
      Item(1).Control(2)=   "feOperacion"
      Item(1).Control(3)=   "btnCancela"
      Item(1).Control(4)=   "btnConsulta"
      Item(1).Control(5)=   "Label2(1)"
      Item(1).Control(6)=   "Label2(2)"
      Item(1).Control(7)=   "Label2(3)"
      Item(1).Control(8)=   "dtpInicio"
      Item(1).Control(9)=   "Label2(5)"
      Item(1).Control(10)=   "vGrid_Facturas"
      Item(1).Control(11)=   "cboEstado"
      Item(1).Control(12)=   "cboFecha"
      Item(1).Control(13)=   "dtpCorte"
      Item(2).Caption =   "Mensajes"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "fraMsj"
      Item(2).Control(1)=   "vGrid"
      Item(2).Control(2)=   "imgBorraMsj"
      Item(2).Control(3)=   "imgMsjNuevo"
      Item(3).Caption =   "Desembolsos"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "lswDesembolsos"
      Item(3).Control(1)=   "lswFacturas"
      Begin XtremeSuiteControls.ListView lswFacturas 
         Height          =   2532
         Left            =   -69880
         TabIndex        =   31
         Top             =   3000
         Visible         =   0   'False
         Width           =   10452
         _Version        =   1441793
         _ExtentX        =   18436
         _ExtentY        =   4466
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswDesembolsos 
         Height          =   2532
         Left            =   -69880
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   10452
         _Version        =   1441793
         _ExtentX        =   18436
         _ExtentY        =   4466
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox fraMsj 
         Height          =   3135
         Left            =   -68560
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   7095
         _Version        =   1441793
         _ExtentX        =   12515
         _ExtentY        =   5530
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Begin XtremeSuiteControls.DateTimePicker dtpMsjVence 
            Height          =   315
            Left            =   4560
            TabIndex        =   23
            Top             =   2520
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   550
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.FlatEdit txtMsj 
            Height          =   1485
            Left            =   480
            TabIndex        =   25
            Top             =   840
            Width           =   6135
            _Version        =   1441793
            _ExtentX        =   10816
            _ExtentY        =   2625
            _StockProps     =   77
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   7095
            _Version        =   1441793
            _ExtentX        =   12515
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Mensaje"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.74
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimiento del Mensaje:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1800
            TabIndex        =   24
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Image imgGuardaMsj 
            Height          =   255
            Left            =   5880
            Picture         =   "frmCxC_Consulta.frx":6852
            Stretch         =   -1  'True
            ToolTipText     =   "Guardar Mensaje"
            Top             =   2520
            Width           =   255
         End
         Begin VB.Image imgMsjCierraFrame 
            Height          =   255
            Left            =   6240
            Picture         =   "frmCxC_Consulta.frx":7031
            Stretch         =   -1  'True
            ToolTipText     =   "Guardar Mensaje"
            Top             =   2520
            Width           =   255
         End
      End
      Begin FPSpreadADO.fpSpread vgCxC 
         Height          =   4572
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   10212
         _Version        =   524288
         _ExtentX        =   18013
         _ExtentY        =   8065
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   14
         SpreadDesigner  =   "frmCxC_Consulta.frx":77EE
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4692
         Left            =   -69880
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   10452
         _Version        =   524288
         _ExtentX        =   18436
         _ExtentY        =   8276
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
         SpreadDesigner  =   "frmCxC_Consulta.frx":9867
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid_Facturas 
         Height          =   4332
         Left            =   -69760
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   10212
         _Version        =   524288
         _ExtentX        =   18013
         _ExtentY        =   7641
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
         MaxCols         =   482
         ScrollBars      =   2
         SpreadDesigner  =   "frmCxC_Consulta.frx":9ED7
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit feFactura 
         Height          =   330
         Left            =   -68320
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit feOperacion 
         Height          =   330
         Left            =   -69760
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   -66520
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.PushButton btnCancela 
         Height          =   315
         Left            =   -60280
         TabIndex        =   13
         ToolTipText     =   "Detalle de Facturas"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Cancela"
         BackColor       =   -2147483633
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   315
         Left            =   -60280
         TabIndex        =   14
         ToolTipText     =   "Detalle de Facturas"
         Top             =   720
         Visible         =   0   'False
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Consulta"
         BackColor       =   -2147483633
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ComboBox cboFecha 
         Height          =   330
         Left            =   -64600
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   330
         Left            =   -63160
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
         Left            =   -61840
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtTotalMonto 
         Height          =   330
         Left            =   1080
         TabIndex        =   32
         Top             =   5280
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalSaldo 
         Height          =   330
         Left            =   3720
         TabIndex        =   33
         Top             =   5280
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalCuota 
         Height          =   330
         Left            =   6360
         TabIndex        =   34
         Top             =   5280
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Rango de Consulta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   -63160
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Base"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   -64600
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   -69760
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "No. Factura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   -68320
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   -66520
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Image imgMsjNuevo 
         Height          =   240
         Left            =   -69640
         Picture         =   "frmCxC_Consulta.frx":A589
         ToolTipText     =   "Crear Nuevo Mensaje"
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBorraMsj 
         Height          =   240
         Left            =   -69280
         Picture         =   "frmCxC_Consulta.frx":AC99
         ToolTipText     =   "Eliminar Mensaje Marcados"
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   5640
         TabIndex        =   6
         Top             =   5280
         Width           =   732
      End
      Begin VB.Label Label1 
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
         Height          =   312
         Index           =   1
         Left            =   3120
         TabIndex        =   5
         Top             =   5280
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   5280
         Width           =   732
      End
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   7560
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":B22D
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":B34B
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":B471
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":B59B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":B6AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":B7C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":B8C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":B9FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":BB11
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":BC35
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Consulta.frx":BD5E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   0
      Top             =   6972
      Width           =   10992
      _ExtentX        =   19394
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   0
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "15/4/2024"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   0
            Object.Width           =   1658
            MinWidth        =   1658
            TextSave        =   "04:38:p. m."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2892
            MinWidth        =   2892
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2541
            MinWidth        =   2541
            Object.ToolTipText     =   "Intereses Pendientes"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3070
            MinWidth        =   3070
            Object.ToolTipText     =   "Cancelación"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Cargo x Anticipo"
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   2280
      TabIndex        =   27
      Top             =   120
      Width           =   5892
      _Version        =   1441793
      _ExtentX        =   10393
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   240
      TabIndex        =   26
      Top             =   120
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
   Begin XtremeSuiteControls.FlatEdit txtRate 
      Height          =   330
      Left            =   8160
      TabIndex        =   28
      Top             =   120
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
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
   Begin VB.Label lblFacturas 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas Descontadas PF?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   5880
      TabIndex        =   29
      Top             =   480
      Width           =   5052
   End
   Begin XtremeSuiteControls.FormExtender FormExtender1 
      Left            =   120
      Top             =   1080
      _Version        =   1441793
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
      Transparency    =   130
   End
   Begin VB.Label lblClasificacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Clasificación ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   540
      Width           =   5055
   End
   Begin VB.Image imgClasificacion 
      Height          =   252
      Left            =   240
      Picture         =   "frmCxC_Consulta.frx":C146
      Stretch         =   -1  'True
      ToolTipText     =   "Clasificacion ABCD de la persona"
      Top             =   540
      Width           =   252
   End
   Begin VB.Image imgNo 
      Height          =   252
      Left            =   7080
      Picture         =   "frmCxC_Consulta.frx":C25A
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image imgSi 
      Height          =   252
      Left            =   6720
      Picture         =   "frmCxC_Consulta.frx":C374
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11652
   End
End
Attribute VB_Name = "frmCxC_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Sub EstadoInicial()
On Error Resume Next

Call Limpia

    txtCedula.Enabled = True
    txtCedula.SetFocus

End Sub

Sub Limpia()
 txtCedula = ""
 
 lblClasificacion.Caption = "Clasificación ?"
 lblFacturas.Caption = "Facturas Descontadas Ultimo año?"
 ssTab.Item(0).Selected = True
 
End Sub


Private Sub sbConsultaCuentas(Optional pSheet As Integer = 1)
Dim strSQL As String, rs As New ADODB.Recordset

Dim curCuota As Currency, curMonto As Currency
Dim curSaldo As Currency, vMora As Boolean
Dim i As Integer

'On Error Resume Next

curCuota = 0
curMonto = 0
curSaldo = 0
vMora = False


txtTotalMonto.Text = ""
txtTotalSaldo.Text = ""
txtTotalCuota.Text = ""

StatusBar.Panels(5).Text = "0.00"
StatusBar.Panels(6).Text = "0.00"

Me.MousePointer = vbHourglass


vMora = False

With vgCxC
 .Sheet = pSheet
 .ActiveSheet = pSheet
 
 
 .MaxRows = 0
 strSQL = "exec spCxC_PersonasCuentas '" & txtCedula.Text & "','" & Mid(.SheetName, 1, 1) & "'"
 Call OpenRecordSet(rs, strSQL)

  Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    
    For i = 1 To .MaxCols
      .col = i
      Select Case i
        Case 1 'Status

              .TypePictPicture = imgSemaforos.ListImages.Item(1).Picture
        
       
            If rs!Warning = 1 Then
               .TypePictPicture = imgSemaforos.ListImages.Item(2).Picture
              .TextTip = TextTipFixed
              .TextTipDelay = 1000
              .CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
              .CellNoteIndicatorColor = vbRed
              .CellNote = "Dias para el vencimiento : " & DateDiff("d", rs!Fecha_Server, rs!Fecha_Pago)
            End If
        
             
             If Mid(rs!Estado, 1, 1) = "C" Then
                .TypePictPicture = imgSemaforos.ListImages.Item(6).Picture
             End If

            'Indicador de Morosidad
            If rs!MoraMonto > 0 Then
              
              .TypePictPicture = imgSemaforos.ListImages.Item(3).Picture
              vMora = True
            
              .TextTip = TextTipFixed
              .TextTipDelay = 1000
            
              .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
              .CellNoteIndicatorColor = vbBlue
              
              .CellNote = "Morosidad:" & vbCrLf _
                        & "   Intereses : " & Format(rs!MoraInt, "Standard") & vbCrLf _
                        & "   Cargos    : " & Format(rs!MoraCargos, "Standard") & vbCrLf _
                        & "   Principal : " & Format(rs!MoraPrincipal, "Standard") & vbCrLf _
                        & "   Días Mora : " & Format(rs!MoraDias, "###,##0") & vbCrLf _
                        & "   Cta. Ultima : " & Format(rs!MoraFecha, "dd-mm-yyyy") & vbCrLf & vbCrLf _
                        & "   Total Mora  : " & Format(rs!MoraMonto, "Standard") & vbCrLf
            
            End If
        
        Case 2 'Operacion
           .CellTag = CStr(rs!Operacion)
           If pSheet = 1 Then
                .TypeCheckText = CStr(rs!Operacion)
           Else
                .Text = CStr(rs!Operacion)
           End If
        
        Case 3 'Concepto
            .Text = rs!cod_Concepto
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
            .CellNoteIndicatorColor = vbBlue
  
            .CellNote = Trim(rs!ConceptoDesc) & vbCrLf & vbCrLf & "Activación: " & Format(rs!Activa_Fecha, "dd/mm/yyyy") & vbCrLf & "Usuario: " & Trim(rs!Activa_Usuario) & vbCrLf & "Oficina:" & rs!OficinaDesc & ""
        
        Case 4 'Documento
            .Text = Trim(rs!Num_Documento & "")
        Case 5 'Fecha Pago
            .Text = Format(rs!Fecha_Pago, "dd/mm/yyyy")
        Case 6 'Monto
            .Text = Format(rs!Monto, "Standard")
        Case 7 'Saldo
            .Text = Format(rs!Saldo, "Standard")
        Case 8 'Cuota
            .Text = Format(rs!Cuota, "Standard")
        Case 9 'Ultimo Movimiento
            .Text = Format(rs!Fecha_UltMov & "", "dd/mm/yyyy")
        
        Case 10 'Estado
            .Text = CStr(rs!Estado)
            
        Case 11 'Pagador
            .Text = rs!Nombre_Pagador & ""
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNote = rs!cedula_pagador & ""
            
        Case 12 'Contrato
            .Text = rs!COD_CONTRATO & ""
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNote = rs!ContratoDesc & ""
            
            
        Case 13 'Oficina
            .Text = CStr(rs!OficinaDesc)
            
        
        Case 14 'Mora Dias
            .Text = CStr(rs!MoraDias)
      
      End Select
    Next i
    
     curMonto = curMonto + rs!Monto
     curSaldo = curSaldo + rs!Saldo
     curCuota = curCuota + rs!Cuota

    rs.MoveNext
  Loop
  rs.Close
  
End With

  
'Totales
txtTotalMonto.Text = Format(curMonto, "Standard")
txtTotalCuota.Text = Format(curCuota, "Standard")
txtTotalSaldo.Text = Format(curSaldo, "Standard")

Me.MousePointer = vbDefault

End Sub


Private Sub sbSolicitudes(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spSIFEstadoSolicitud '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)

With vgCxC
    .ActiveSheet = 3
    .Sheet = 3
    .MaxRows = 0
    
    Do While Not rs.EOF
    
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    For i = 1 To .MaxCols
      .col = i
      Select Case i
        Case 1 'Status
            .TypePictPicture = imgSemaforos.ListImages.Item(5).Picture
             
        Case 2 'Operacion
            .Text = CStr(rs!Operacion)
        Case 3 'Linea
            .Text = CStr(rs!cod_Concepto)
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
            .CellNoteIndicatorColor = vbBlue
  
            .CellNote = Trim(rs!LineaX) & vbCrLf & vbCrLf & "Solicitado: " & Format(rs!FechaSol, "dd/mm/yyyy") & vbCrLf & "Usuario: " & Trim(rs!userRec)
        
        Case 4 'Cédula
            .Text = CStr(rs!Cedula)
        Case 5 'Solicitud
            .Text = Format(rs!FechaSol, "dd/mm/yyyy")
        Case 6 'Monto
            .Text = Format(rs!montosol, "Standard")
        Case 7 'Estado
            Select Case rs!estadosol
             Case "R"
              .Text = "Recibida"
             Case "P"
              .Text = "Pendiente"
             Case "A"
              .Text = "Aprobada"
             Case "D"
              .Text = "Denegada"
             Case "F"
              .Text = "Formalizada"
             Case "N"
              .Text = "Anulada"
            End Select
      End Select
     
     
    Next i

     rs.MoveNext
    Loop
    rs.Close

End With
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbPreAnalisis(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spSIFEstadoPreAnalisis '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)


With vgCxC
    .ActiveSheet = 4
    .Sheet = 4
    .MaxRows = 0
    
    Do While Not rs.EOF
    
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    For i = 1 To .MaxCols
      .col = i
      Select Case i
        Case 1 'Status
            .TypePictPicture = imgSemaforos.ListImages.Item(4).Picture
             
        Case 2 'Expediente
            .Text = CStr(rs!cod_PreAnalisis)
        
        Case 3 'Tipo
            .Text = CStr(rs!Tipo)
            
        Case 4 'Linea
            .Text = CStr(rs!cod_linea)
        
        Case 5 'Monto
            .Text = Format(rs!Monto, "Standard")
        
        Case 6 'Estado
            .Text = CStr(rs!Estado)
        
        Case 7 'Operacion
            .Text = CStr(rs!Operacion & "")
        
        Case 8 'Fecha
            .Text = CStr(rs!fecha_creacion & "")
        
        Case 9 'Usuario
            .Text = CStr(rs!Usuario & "")
        
      End Select
     
     
    Next i

     rs.MoveNext
    Loop
    rs.Close

End With
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbIncobrable(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer


On Error GoTo vError

'
'lswDetalle.ColumnHeaders.Add , , "Operación", 1200
'lswDetalle.ColumnHeaders.Add , , "ID", 400
'lswDetalle.ColumnHeaders.Add , , "Línea", 800
'lswDetalle.ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
'lswDetalle.ColumnHeaders.Add , , "Estado", 1200
'lswDetalle.ColumnHeaders.Add , , "Usuario", 1400
'lswDetalle.ColumnHeaders.Add , , "Fecha", 1200
'lswDetalle.ColumnHeaders.Add , , "Documento", 1400
'lswDetalle.ColumnHeaders.Add , , "Notas", 2400
'

Me.MousePointer = vbHourglass

strSQL = "exec spSIFEstadoIncobrable '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)


With vgCxC
    .ActiveSheet = 5
    .Sheet = 5
    .MaxRows = 0
    
    Do While Not rs.EOF
    
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    For i = 1 To .MaxCols
      .col = i
      Select Case i
        Case 1 'Status
            .TypePictPicture = imgSemaforos.ListImages.Item(7).Picture
             
        Case 2 'Operacion
            .Text = CStr(rs!Operacion)
        
        Case 3 'Linea
            .Text = CStr(rs!cod_Concepto)
        
        Case 4 'Monto
            .Text = Format(rs!Saldo + rs!IntCor + rs!IntMor + rs!Cargos + rs!Poliza, "Standard")
        
        Case 5 'Estado
            .Text = CStr(rs!EstadoX)
        
        Case 6 'Usuario
            .Text = CStr(rs!Registro_Usuario & "")
        
        Case 7 'Fecha
            .Text = Format(rs!Registro_Fecha, "dd/mm/yyyy")
        
        Case 8 'Documento
            .Text = "NC." & rs!genera_documento
        
        Case 9 'Notas
            .Text = "[Id.Incobrable: " & rs!cod_incobrable & " ] " & rs!Notas_Registro
        
      End Select
     
     
    Next i

     rs.MoveNext
    Loop
    rs.Close

End With
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub btnCancela_Click()
On Error GoTo vError

GLOBALES.gTag = "0"
GLOBALES.gTag2 = ""
GLOBALES.gTag3 = txtCedula.Text

Call sbFormsCall("frmCxC_Facturas_Cancela", , , , False, Me)

'Call sbFacturas_Buscar

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnConsulta_Click()
Call sbFacturas_Buscar
End Sub


Private Sub sbFacturas_Buscar()
Dim strSQL As String, i As Integer
Dim vCadena As String, iCantidad As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
iCantidad = 0

strSQL = "select Operacion,cod_Factura,Fecha_Pago,  Monto, Factura_Estado_Desc" _
       & ", Adelanto_Monto,Liberado, Pagador_Nombre" _
       & "  From vCxC_Facturas_Control" _
       & " where cedula = '" & txtCedula.Text & "'" _
       & " and cod_Factura like '%" & feFactura.Text & "%'"
    
If IsNumeric(feOperacion.Text) Then
  strSQL = strSQL & " and operacion = " & feOperacion.Text

End If

Select Case cboFecha.Text
  Case "Registro"
    strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Emisión"
    strSQL = strSQL & " and fecha_Emision between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Pago"
    strSQL = strSQL & " and fecha_Pago between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Libera"
    strSQL = strSQL & " and Liberado_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

  Case "Cancela"
    strSQL = strSQL & " and Cancela_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

  Case "Activación"
    strSQL = strSQL & " and Activa_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

  Case "Desembolso 1"
    strSQL = strSQL & " and Pago_Principal_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

  Case "Desembolso 2"
    strSQL = strSQL & " and Pago_Secundario_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

End Select

If cboEstado.Text <> "TODOS" Then
     strSQL = strSQL & " and factura_estado = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
End If

Call sbCargaGridLocal(vGrid_Facturas, 8, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim curMonto As Currency

On Error GoTo vError

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.col = i
 vGrid.Text = ""
Next i

curMonto = 0

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i

'    If rs.Fields(i - 1).Type = 135 Then
'        If Year(rs.Fields(i - 1).Value) > 1900 Then
'           vGrid.Text = Format((rs.Fields(i - 1).Value & ""), "dd/mm/yyyy")
'        End If
'    Else
'        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
'    End If
    vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  curMonto = curMonto + rs!Monto
  rs.MoveNext
Loop

StatusBar.Panels(5).Text = "Casos ..: " & Format(rs.RecordCount, "###,###,##0")
StatusBar.Panels(6).Text = "Monto ..: " & Format(curMonto, "Standard")

rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub






Private Sub cboEstado_Click()
If vPaso Then Exit Sub
Call sbFacturas_Buscar
End Sub

Private Sub cboFecha_Click()
If vPaso Then Exit Sub
Call sbFacturas_Buscar
End Sub

Private Sub dtpCorte_Change()
If vPaso Then Exit Sub
Call sbFacturas_Buscar

End Sub

Private Sub dtpInicio_Change()
If vPaso Then Exit Sub
Call sbFacturas_Buscar

End Sub

Private Sub feFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   cboEstado.SetFocus
   Call sbFacturas_Buscar
End If
End Sub

Private Sub feOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   feFactura.SetFocus
   Call sbFacturas_Buscar
End If
End Sub

Private Sub FlatEdit1_Change()

End Sub

Private Sub Form_Load()
'Tiene Menú Interno


vModulo = 31

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


vGrid.AppearanceStyle = fxGridStyle
vGrid_Facturas.AppearanceStyle = fxGridStyle

vgCxC.AppearanceStyle = fxGridStyle
vgCxC.ActiveSheet = 1
vgCxC.Sheet = 1
vgCxC.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)

Call EstadoInicial
 dtpMsjVence.Value = fxFechaServidor


ssTab.Item(0).Selected = True

StatusBar.Panels(4).Text = glogon.Usuario
StatusBar.Panels(5).Text = "0.00"
StatusBar.Panels(6).Text = "0.00"
StatusBar.Panels(7).Text = "0.00"


With lswFacturas.ColumnHeaders
   .Add , , "Operación", 1200
   .Add , , "Factura", 1400
   .Add , , "Monto", 1400, vbRightJustify
   .Add , , "Adelanto", 1400, vbRightJustify
   .Add , , "Liberado", 1400, vbRightJustify
   .Add , , "Divisa", 1100
   .Add , , "T.C.", 1400, vbRightJustify
   .Add , , "Op.Origen", 1200
   
End With

With lswDesembolsos
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Operación", 1800
    .ColumnHeaders.Add , , "Monto", 2200, vbRightJustify
    .ColumnHeaders.Add , , "Tesoreria [Id]", 2200, vbCenter
    .ColumnHeaders.Add , , "Estado", 1200, vbCenter
    .ColumnHeaders.Add , , "Fecha", 1800, vbCenter
    .ColumnHeaders.Add , , "Tipo", 1200, vbCenter
    .ColumnHeaders.Add , , "Banco", 3200
    .ColumnHeaders.Add , , "Beneficiario", 3200
    .ColumnHeaders.Add , , "No.Documento", 1400, vbCenter
    .ColumnHeaders.Add , , "No.Remesa", 1400, vbCenter
    .ColumnHeaders.Add , , "No.Giro", 1400, vbCenter
End With


Me.Width = 11235
Me.Height = 6930


End Sub

Private Sub sbBusqueda()

On Error GoTo vError

gBusquedas.Convertir = "N"

    Call Limpia
    gBusquedas.Consulta = "Select cedula,nombre from CxC_Personas"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    txtCedula = Trim(gBusquedas.Resultado)
    gBusquedas.Consulta = ""
    gBusquedas.Columna = ""
    gBusquedas.Orden = ""
    gBusquedas.Resultado = ""
    If Trim(txtCedula) <> "" Then
        Call sbConsulta(txtCedula)
    End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Resize()

On Error Resume Next

ssTab.Width = Me.Width - 400
ssTab.Height = Me.Height - 2100

vgCxC.Width = ssTab.Width - 350
vgCxC.Height = ssTab.Height - 1160

txtTotalMonto.top = vgCxC.top + vgCxC.Height + 80
Label1.Item(0).top = txtTotalMonto.top
Label1.Item(1).top = txtTotalMonto.top
Label1.Item(2).top = txtTotalMonto.top

txtTotalCuota.top = txtTotalMonto.top
txtTotalSaldo.top = txtTotalMonto.top

vGrid.Width = vgCxC.Width
vGrid.Height = ssTab.Height - 850

vGrid_Facturas.Width = vGrid.Width
vGrid_Facturas.Height = ssTab.Height - 1300

lswDesembolsos.Width = vGrid.Width
lswDesembolsos.Height = (ssTab.Height - 700) / 2

lswFacturas.Width = vGrid.Width
lswFacturas.top = lswDesembolsos.top + lswDesembolsos.Height + 120

lswFacturas.Height = lswDesembolsos.Height

imgBanner.Width = Me.Width


End Sub

Private Sub imgBorraMsj_Click()
Dim strSQL As String, i As Integer
Dim msj(2) As String

On Error GoTo vError

With vGrid
    For i = 1 To vGrid.MaxRows
      .Row = i
      .col = 3
        If .Value = 1 Then
           .col = 1
           msj(0) = Format(.Text, "yyyy/mm/dd")
           msj(1) = .CellTag
           .col = 2
           msj(2) = Mid(.Text, 1, 15)
           
           strSQL = "delete CxC_Personas_Mensajes where cedula = '" & txtCedula _
                  & "' and usuario = '" & msj(1) & "' and vencimiento = '" _
                  & msj(0) & "' and substring(mensaje,1,15) = '" _
                  & msj(2) & "'"
           Call ConectionExecute(strSQL)
        End If
    Next i
End With

Call sbCargaMsj(txtCedula)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub imgGuardaMsj_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "insert CxC_Personas_Mensajes(fecha,cedula,usuario,vencimiento,mensaje) values(dbo.MyGetdate(),'" _
       & txtCedula & "','" & glogon.Usuario & "','" & Format(dtpMsjVence.Value, "yyyy/mm/dd") & "','" _
       & txtMsj & "')"
Call ConectionExecute(strSQL)

txtMsj = ""
fraMsj.Visible = False
MsgBox "Mensaje Registrado...", vbInformation

Call sbCargaMsj(txtCedula)

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub imgMsjCierraFrame_Click()
fraMsj.Visible = False
End Sub

Private Sub imgMsjNuevo_Click()
fraMsj.Visible = True
End Sub

Private Sub sbCargaMsj(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

'Inicializa Datos y Encabezados
dtpMsjVence.Value = fxFechaServidor
vGrid.MaxRows = 0
vGrid.MaxCols = 5

txtMsj = ""
fraMsj.Visible = False

strSQL = "select * from CxC_Personas_Mensajes where cedula = '" _
       & vCedula & "' and datediff(d,dbo.MyGetdate(),vencimiento) >= 0"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  vGrid.col = 1
  vGrid.Text = Format(rs!Vencimiento, "dd/mm/yyyy")
  vGrid.TextTip = TextTipFixed
  vGrid.TextTipDelay = 1000

  vGrid.CellNote = "Fecha : " & rs!fecha & vbCrLf & "Usuario : " & rs!Usuario
  vGrid.CellTag = rs!Usuario
   
    
  vGrid.col = 2
  vGrid.Text = rs!Mensaje
      
  vGrid.col = 4
  vGrid.Text = rs!fecha
      
  vGrid.col = 5
  vGrid.Text = rs!Usuario
      
  
  vGrid.RowHeight(vGrid.Row) = vGrid.MaxTextRowHeight(vGrid.Row)
  
 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub






Private Sub sbConsulta(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

Dim vFechaIng As Date, vFianzas As Boolean
Dim rsTmp As New ADODB.Recordset
     
     
If Not fxSIFValidaCadena(txtCedula.Text) Then
   Exit Sub
End If
  
 
strSQL = "select P.cedula,P.nombre,C.descripcion as 'CategoriaDesc',dbo.fxCxC_PersonasMsjNum(P.cedula) as IndMensajes" _
       & ", isnull(Fa.Facturas,0) as 'Facturas', isnull(Fa.Total,0) as 'Facturas_Total'" _
       & " from CxC_Personas P inner join CxC_Categoria_Clientes C on P.cod_categoria = C.cod_Categoria" _
       & " left join vCxC_C_Persona_Facturas_Anio Fa on P.cedula = Fa.Cedula" _
       & " Where P.Cedula = '" & Trim(pCedula) & "'"

Call OpenRecordSet(rs, strSQL)
 
If Not rs.EOF And Not rs.BOF Then
   
  
   txtCedula.Text = Trim(rs!Cedula & "")
   txtNombre.Text = rs!Nombre & ""
   
     'Clasificación de la Persona
     lblClasificacion.Caption = "Categoría de Cliente : [" & rs!CategoriaDesc & "]"
     lblFacturas.Caption = "Facturas Ultimo Año [ " & Format(rs!facturas, "###,###,###") & " ]" _
                         & vbCrLf & "[ " & Format(rs!Facturas_Total, "Standard") & " ]"
    
     'Indica los Mensajes
     If rs!IndMensajes = 0 Then
       txtRate.Text = "Mensajes"
     Else
       txtRate.Text = "Mensajes (" & rs!IndMensajes & ")"
     End If
     
    
     ssTab.Item(0).Selected = True
     
     'Actualiza el Detalle de Cuentas
     Call sbConsultaCuentas
     
 Else
   MsgBox "No Se encontró registro de la persona solicitada", vbInformation
   Exit Sub
 End If
   
 rs.Close
   
 
End Sub

Public Sub sbXConsultaAsistida(vCedula As String)
  txtCedula = vCedula
  Call txtCedula_KeyDown(vbKeyReturn, 0)
End Sub


Private Sub sbCargaDesembolsos(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

On Error GoTo vError

strSQL = "select Ct.OPERACION, Cg.MONTO, Td.ESTADO, Td.FECHA_EMISION , Td.TIPO, Cg.Id_Giro" _
        & ", Tb.DESCRIPCION as 'BancoDesc', Td.BENEFICIARIO , Cg.TESORERIA_SOLICITUD, Cg.TESORERIA_REMESA , Td.NDOCUMENTO" _
        & "    from CXC_CUENTAS Ct inner join CXC_CUENTAS_GIROS Cg on Ct.OPERACION = Cg.OPERACION" _
        & "           inner join TES_TRANSACCIONES Td on Cg.TESORERIA_SOLICITUD = Td.NSOLICITUD" _
        & "           inner join TES_BANCOS Tb on Td.ID_BANCO = Tb.ID_BANCO" _
        & "           inner join TES_BANCOS_GRUPOS Bg on Tb.COD_GRUPO = Bg.COD_GRUPO" _
        & "    where Ct.CEDULA = '" & txtCedula.Text & "'"

vPaso = True

With lswDesembolsos
    .ListItems.Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
        Set itmX = .ListItems.Add(, , rs!Operacion)
            itmX.SubItems(1) = Format(rs!Monto, "Standard")
            itmX.SubItems(2) = rs!Tesoreria_Solicitud
            itmX.SubItems(3) = rs!Estado
            itmX.SubItems(4) = rs!Fecha_Emision & ""
            itmX.SubItems(5) = rs!Tipo
            itmX.SubItems(6) = rs!BancoDesc
            itmX.SubItems(7) = rs!Beneficiario
            itmX.SubItems(8) = rs!nDocumento & ""
            itmX.SubItems(9) = rs!TESORERIA_REMESA
            itmX.SubItems(10) = rs!ID_GIRO
            
        rs.MoveNext
    Loop
    rs.Close

End With

vPaso = False

Exit Sub

vError:
    vPaso = False
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswDesembolsos_DblClick()

Dim frm As Form

If Not IsNumeric(lswDesembolsos.SelectedItem.SubItems(2)) Then Exit Sub
If CCur(lswDesembolsos.SelectedItem.SubItems(2)) <= 0 Then Exit Sub
 

 Call sbFormsCall("frmTES_Transacciones")
 For Each frm In Forms
   If UCase(frm.Name) = UCase("frmTES_Transacciones") Then
     Call frm.sbTESDocConsulta(lswDesembolsos.SelectedItem.SubItems(2))
     Exit For
   End If
 Next frm
End Sub

Private Sub lswDesembolsos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, pOperacion As Long, pGiro As Integer

On Error GoTo vError

lswFacturas.ListItems.Clear


pOperacion = Item.Text
pGiro = Item.SubItems(10)

strSQL = "select OPERACION, COD_FACTURA, MONTO, ADELANTO_MONTO , LIBERADO, COD_DIVISA , TIPO_CAMBIO, OPERACION_ORIGEN" _
       & " From CXC_CUENTAS_FACTURAS" _
       & " where OPERACION = " & pOperacion & " and " & pGiro & " in(ID_GIRO , ID_GIRO_PENDIENTE)"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswFacturas.ListItems.Add(, , rs!Operacion)
      itmX.SubItems(1) = rs!cod_Factura
      itmX.SubItems(2) = Format(rs!Monto, "Standard")
      itmX.SubItems(3) = Format(rs!Adelanto_Monto, "Standard")
      itmX.SubItems(4) = Format(rs!LIBERADO, "Standard")
      itmX.SubItems(5) = rs!cod_Divisa
      itmX.SubItems(6) = rs!TIPO_CAMBIO
      itmX.SubItems(7) = rs!Operacion_Origen
 
  rs.MoveNext
Loop
rs.Close

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub ssTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
  Case 1 'Facturas
    Call sbFacturas_Buscar
  Case 2 'Mensajes
    Call sbCargaMsj(txtCedula)
  Case 3 'Desembolsos
    Call sbCargaDesembolsos(txtCedula.Text)
  Case Else
End Select
End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

TimerX.Interval = 0
TimerX.Enabled = False

vPaso = True

strSQL = "select FACTURA_ESTADO as 'IdX', DESCRIPCION as 'ItmX' from CXC_FACTURAS_ESTADOS"
Call sbCbo_Llena_New(cboEstado, strSQL, True, True)


cboFecha.Clear
cboFecha.AddItem "Registro"
cboFecha.AddItem "Emisión"
cboFecha.AddItem "Pago"
cboFecha.AddItem "Libera"
cboFecha.AddItem "Cancela"
cboFecha.AddItem "Activación"
cboFecha.AddItem "Desembolso 1"
cboFecha.AddItem "Desembolso 2"
cboFecha.AddItem "[TODAS]"

cboFecha.Text = "Pago"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)
dtpCorte.Value = DateAdd("d", 7, dtpCorte.Value)

vPaso = False

txtCedula.SetFocus

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCedTemp As String

On Error GoTo vError

vCedTemp = ""

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 strSQL = "select isnull(count(*),0) as Existe from CxC_Personas where cedula = '" & txtCedula & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
   rs.Close
   strSQL = "select cedula from CxC_Cuentas where Operacion = " & txtCedula
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
      vCedTemp = Trim(rs!Cedula)
   End If
 End If
 rs.Close
 
    If vCedTemp = "" Then
     Call sbConsulta(txtCedula.Text)
    Else
     Call sbConsulta(vCedTemp)
    End If
End If

If KeyCode = vbKeyF4 Then Call sbBusqueda

Exit Sub

vError:

End Sub






Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub vgCxC_DblClick(ByVal col As Long, ByVal Row As Long)
Dim frm As Form

On Error GoTo vError

With vgCxC
    .Sheet = .ActiveSheet
    .Row = .ActiveRow
    

    Select Case .ActiveSheet
      Case 1 'Creditos Activos
         If .MaxRows = 0 Then Exit Sub
      
'         .Col = 2
'         Operacion.OperacionConsulta = .CellTag
'         frmCR_ConsultaDetalle.Show vbModal
      
      
      Case 2 'Creditos Cancelados
         If .MaxRows = 0 Then Exit Sub
'         .Col = 2
'         Operacion.OperacionConsulta = .Text
'         frmCR_ConsultaDetalle.Show vbModal
      
     
      Case 3 'Tramite
'         .Col = 7
'            Set X = New clsPreaAnalisis
'            Set X.vCon = glogon.Conection
'            X.xOperacion = .Text
'            X.xkey = glogon.ConectRPT
'
'
'         .Col = 2
'            If .MaxRows = 0 Then
'                X.vSolicitudPreanalisis = 0
'            Else
'                X.vSolicitudPreanalisis = .Text
'            End If
'            Call X.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
'                        , App.Path, glogon.ConectRPT, 11)
'
'            Set X = Nothing
      
    End Select
    

End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub vgCxC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
    Call PopupMenu(MDIPrincipal.mnuCxC, , x, y * 2)
End If
End Sub

Private Sub vgCxC_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim i As Integer


With MDIPrincipal.mnuAccionesSub
  For i = 0 To .Count - 1
      .Item(i).Visible = True
  Next i
End With



Select Case NewSheet
    Case 1 'Activos
       Call sbConsultaCuentas(NewSheet)
     
     Case 2 'Cancelados
       Call sbConsultaCuentas(NewSheet)
    
    Case 3  'Tramites
       Call sbConsultaCuentas(NewSheet)
End Select

End Sub







