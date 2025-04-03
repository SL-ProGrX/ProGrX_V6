VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_LiquidacionW 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "frmAF_LiquidacionW.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   10275
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   255
      Left            =   7920
      TabIndex        =   86
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Asiento"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox fraPrg 
      Height          =   1092
      Left            =   1920
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   5172
      _Version        =   1441793
      _ExtentX        =   9123
      _ExtentY        =   1926
      _StockProps     =   79
      Caption         =   "Procesando Liquidación"
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
      Appearance      =   17
      BorderStyle     =   1
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   252
         Left            =   240
         TabIndex        =   76
         Top             =   840
         Width           =   4692
         _ExtentX        =   8281
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblX 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   240
         TabIndex        =   77
         Top             =   240
         Width           =   4692
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5775
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   10095
      _Version        =   1441793
      _ExtentX        =   17806
      _ExtentY        =   10186
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
      ItemCount       =   5
      Item(0).Caption =   "Renuncia"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "GroupBox1(2)"
      Item(0).Control(1)=   "txtCedula"
      Item(0).Control(2)=   "txtNombre"
      Item(0).Control(3)=   "Label1(0)"
      Item(0).Control(4)=   "GroupBox1(4)"
      Item(0).Control(5)=   "GroupBox1(1)"
      Item(0).Control(6)=   "GroupBox1(0)"
      Item(1).Caption =   "Patrimonio"
      Item(1).ControlCount=   29
      Item(1).Control(0)=   "txtRetenerMonto"
      Item(1).Control(1)=   "Label4(4)"
      Item(1).Control(2)=   "lblTotalNeto(0)"
      Item(1).Control(3)=   "Label4(2)"
      Item(1).Control(4)=   "Label4(1)"
      Item(1).Control(5)=   "lblTotalBruto"
      Item(1).Control(6)=   "Label4(0)"
      Item(1).Control(7)=   "Label3(0)"
      Item(1).Control(8)=   "lblAporteExtra"
      Item(1).Control(9)=   "lblCapitalizacion"
      Item(1).Control(10)=   "lblFCI"
      Item(1).Control(11)=   "lblAportePatronal"
      Item(1).Control(12)=   "lblAporteObrero"
      Item(1).Control(13)=   "chkAplObrero"
      Item(1).Control(14)=   "chkAplPatronal"
      Item(1).Control(15)=   "chkAplCapGen"
      Item(1).Control(16)=   "chkAplCapExtra"
      Item(1).Control(17)=   "lblRenta"
      Item(1).Control(18)=   "Label3(1)"
      Item(1).Control(19)=   "chkAplExcedente"
      Item(1).Control(20)=   "Label3(2)"
      Item(1).Control(21)=   "lblExcedenteRenta"
      Item(1).Control(22)=   "lblExcedente"
      Item(1).Control(23)=   "lblCustodia"
      Item(1).Control(24)=   "txtDivisa"
      Item(1).Control(25)=   "scTitulos(0)"
      Item(1).Control(26)=   "txtTipoCambio"
      Item(1).Control(27)=   "Label4(8)"
      Item(1).Control(28)=   "txtDivisaLocal"
      Item(2).Caption =   "Planes de Ahorros"
      Item(2).ControlCount=   10
      Item(2).Control(0)=   "lswPlanes"
      Item(2).Control(1)=   "Label4(5)"
      Item(2).Control(2)=   "lblTotalNeto(1)"
      Item(2).Control(3)=   "lblTotalNeto(2)"
      Item(2).Control(4)=   "Label4(3)"
      Item(2).Control(5)=   "Label4(6)"
      Item(2).Control(6)=   "Label4(7)"
      Item(2).Control(7)=   "txtFndRendGravado"
      Item(2).Control(8)=   "txtFndRendLiquidar"
      Item(2).Control(9)=   "scTitulos(1)"
      Item(3).Caption =   "Abonos"
      Item(3).ControlCount=   5
      Item(3).Control(0)=   "cmdDistribucionAuto"
      Item(3).Control(1)=   "fraAbono"
      Item(3).Control(2)=   "lblLsw"
      Item(3).Control(3)=   "Label2"
      Item(3).Control(4)=   "lswAbonos"
      Item(4).Caption =   "Resumen"
      Item(4).ControlCount=   2
      Item(4).Control(0)=   "fraObservacion"
      Item(4).Control(1)=   "fraResumen"
      Begin XtremeSuiteControls.ListView lswAbonos 
         Height          =   4695
         Left            =   -69880
         TabIndex        =   46
         Top             =   960
         Visible         =   0   'False
         Width           =   9855
         _Version        =   1441793
         _ExtentX        =   17383
         _ExtentY        =   8281
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
         View            =   3
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswPlanes 
         Height          =   3735
         Left            =   -69880
         TabIndex        =   43
         Top             =   840
         Visible         =   0   'False
         Width           =   9855
         _Version        =   1441793
         _ExtentX        =   17383
         _ExtentY        =   6588
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
         Checkboxes      =   -1  'True
         View            =   3
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.GroupBox fraResumen 
         Height          =   5175
         Left            =   -70000
         TabIndex        =   73
         Top             =   360
         Visible         =   0   'False
         Width           =   10095
         _Version        =   1441793
         _ExtentX        =   17806
         _ExtentY        =   9128
         _StockProps     =   79
         Caption         =   "Resumen"
         ForeColor       =   0
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton cmdAplicar 
            Height          =   615
            Left            =   8160
            TabIndex        =   74
            Top             =   4440
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Aplicar"
            ForeColor       =   0
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmAF_LiquidacionW.frx":08CA
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.FlatEdit txtSumario 
            Height          =   3855
            Left            =   0
            TabIndex        =   111
            Top             =   480
            Width           =   9975
            _Version        =   1441793
            _ExtentX        =   17595
            _ExtentY        =   6800
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   10.5
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
         End
         Begin XtremeShortcutBar.ShortcutCaption lblSumario 
            Height          =   375
            Left            =   0
            TabIndex        =   112
            Top             =   0
            Width           =   10095
            _Version        =   1441793
            _ExtentX        =   17806
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "..:: Sumario ::.."
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
         Begin VB.Image imgObservacion 
            Height          =   480
            Left            =   240
            Picture         =   "frmAF_LiquidacionW.frx":10A8
            ToolTipText     =   "Digitar Observaciones"
            Top             =   4560
            Width           =   480
         End
      End
      Begin VB.TextBox txtRetenerMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -66520
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "frmAF_LiquidacionW.frx":1A21
         Top             =   4080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1215
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   9495
         _Version        =   1441793
         _ExtentX        =   16748
         _ExtentY        =   2143
         _StockProps     =   79
         ForeColor       =   0
         BackColor       =   -2147483633
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboCausa 
            Height          =   330
            Left            =   3720
            TabIndex        =   5
            Top             =   360
            Width           =   5295
            _Version        =   1441793
            _ExtentX        =   9340
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
         Begin XtremeSuiteControls.ComboBox cboTipo 
            Height          =   315
            Left            =   840
            TabIndex        =   6
            Top             =   360
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
         Begin XtremeSuiteControls.CheckBox chkMortalidad 
            Height          =   255
            Left            =   4680
            TabIndex        =   7
            Top             =   720
            Width           =   4335
            _Version        =   1441793
            _ExtentX        =   7646
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Aplicar procedimiento por Causa de Muerte"
            ForeColor       =   0
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkAltPlanilla 
            Height          =   255
            Left            =   4080
            TabIndex        =   8
            Top             =   960
            Width           =   4935
            _Version        =   1441793
            _ExtentX        =   8705
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cancela Créditos Pendientes por Medio de Planilla ?"
            ForeColor       =   0
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
            Alignment       =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Causa"
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
            Left            =   2640
            TabIndex        =   10
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo"
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
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1095
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16960
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   "Desembolso"
         ForeColor       =   0
         BackColor       =   -2147483633
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboBanco 
            Height          =   330
            Left            =   1080
            TabIndex        =   13
            Top             =   360
            Width           =   5175
            _Version        =   1441793
            _ExtentX        =   9128
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
         Begin XtremeSuiteControls.ComboBox cboCuenta 
            Height          =   330
            Left            =   1080
            TabIndex        =   14
            Top             =   720
            Width           =   5175
            _Version        =   1441793
            _ExtentX        =   9128
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
         Begin XtremeSuiteControls.ComboBox cboTipoDoc 
            Height          =   315
            Left            =   7440
            TabIndex        =   15
            Top             =   360
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
         Begin XtremeSuiteControls.DateTimePicker dtpPago 
            Height          =   315
            Left            =   7440
            TabIndex        =   115
            Top             =   720
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2773
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
         Begin VB.Label Label16 
            Caption         =   "F.Pago"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   6720
            TabIndex        =   116
            Top             =   750
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta"
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
            Left            =   0
            TabIndex        =   18
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Banco"
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
            Index           =   6
            Left            =   0
            TabIndex        =   17
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "T.Doc."
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
            Index           =   5
            Left            =   6720
            TabIndex        =   16
            Top             =   360
            Width           =   615
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   735
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   3840
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16960
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Datos Actuales "
         ForeColor       =   0
         BackColor       =   -2147483633
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.Label lblBoleta 
            Height          =   315
            Left            =   7320
            TabIndex        =   125
            Top             =   360
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   79
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
            Alignment       =   2
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblIngreso 
            Height          =   315
            Left            =   3840
            TabIndex        =   124
            Top             =   360
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   79
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
            Alignment       =   2
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblEstadoActual 
            Height          =   315
            Left            =   960
            TabIndex        =   123
            Top             =   360
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   79
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
            Alignment       =   2
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "Boleta"
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
            Index           =   11
            Left            =   6480
            TabIndex        =   22
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   10
            Left            =   2760
            TabIndex        =   21
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Estado"
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
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   612
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   4920
         Width           =   9495
         _Version        =   1441793
         _ExtentX        =   16748
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Datos de la Acción de Personal "
         ForeColor       =   0
         BackColor       =   -2147483633
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.DateTimePicker dtpAc_fecha 
            Height          =   315
            Left            =   2760
            TabIndex        =   89
            Top             =   360
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2773
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
         Begin XtremeSuiteControls.FlatEdit txtAc_Boleta 
            Height          =   330
            Left            =   7200
            TabIndex        =   117
            Top             =   360
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboAc_Tipo 
            Height          =   330
            Left            =   5280
            TabIndex        =   122
            Top             =   360
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
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
         Begin XtremeSuiteControls.Label lblAcFecha 
            Height          =   315
            Left            =   2760
            TabIndex        =   126
            Top             =   360
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   79
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
            Alignment       =   2
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblAC_Tipo 
            Caption         =   "Tipo Acción"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5280
            TabIndex        =   121
            Top             =   180
            Width           =   1260
         End
         Begin VB.Image imgFechaAccion 
            Height          =   240
            Left            =   4560
            Picture         =   "frmAF_LiquidacionW.frx":1A23
            Top             =   375
            Width           =   240
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Rige a partir de"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   25
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label lblAc_Boleta 
            Caption         =   "Boleta"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7215
            TabIndex        =   24
            Top             =   180
            Width           =   660
         End
      End
      Begin XtremeSuiteControls.CheckBox chkAplObrero 
         Height          =   252
         Left            =   -69520
         TabIndex        =   39
         Top             =   1080
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1441793
         _ExtentX        =   4466
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica Aporte Obrero"
         ForeColor       =   0
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkAplPatronal 
         Height          =   252
         Left            =   -69520
         TabIndex        =   40
         Top             =   1440
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1441793
         _ExtentX        =   4466
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica Aporte Patronal"
         ForeColor       =   0
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
      Begin XtremeSuiteControls.CheckBox chkAplCapGen 
         Height          =   252
         Left            =   -69520
         TabIndex        =   41
         Top             =   2280
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1441793
         _ExtentX        =   4466
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica Capitalización"
         ForeColor       =   0
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
      Begin XtremeSuiteControls.CheckBox chkAplCapExtra 
         Height          =   252
         Left            =   -69520
         TabIndex        =   42
         Top             =   3120
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1441793
         _ExtentX        =   4466
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica Ahorro Extraordinario"
         ForeColor       =   0
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
      Begin VB.Frame fraAbono 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4332
         Left            =   -69880
         TabIndex        =   47
         Top             =   840
         Visible         =   0   'False
         Width           =   8532
         Begin VB.TextBox txtMAbono 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5880
            MultiLine       =   -1  'True
            TabIndex        =   53
            Top             =   3000
            Width           =   1812
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   1212
            Left            =   600
            TabIndex        =   48
            Top             =   3240
            Width           =   7092
            _Version        =   1441793
            _ExtentX        =   12509
            _ExtentY        =   2138
            _StockProps     =   79
            BackColor       =   -2147483633
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.PushButton cmdMAceptar 
               Height          =   372
               Left            =   4440
               TabIndex        =   49
               Top             =   360
               Width           =   1332
               _Version        =   1441793
               _ExtentX        =   2350
               _ExtentY        =   656
               _StockProps     =   79
               Caption         =   "Aceptar"
               ForeColor       =   0
               BackColor       =   -2147483633
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
            Begin XtremeSuiteControls.PushButton cmdMCancelar 
               Height          =   372
               Left            =   5760
               TabIndex        =   50
               Top             =   360
               Width           =   1332
               _Version        =   1441793
               _ExtentX        =   2350
               _ExtentY        =   656
               _StockProps     =   79
               Caption         =   "Cancelar"
               ForeColor       =   0
               BackColor       =   -2147483633
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
            Begin VB.Label lblDisponible 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
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
               Height          =   312
               Left            =   1920
               TabIndex        =   52
               Top             =   360
               Width           =   1812
            End
            Begin VB.Label Label10 
               Caption         =   "Disponible"
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
               Left            =   240
               TabIndex        =   51
               Top             =   360
               Width           =   1332
            End
         End
         Begin VB.Label Label19 
            Caption         =   "Mora/Vencido Total:"
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
            Left            =   240
            TabIndex        =   104
            Top             =   1440
            Width           =   1212
         End
         Begin VB.Label lblMMoraTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   1560
            TabIndex        =   103
            Top             =   1560
            Width           =   1932
         End
         Begin VB.Label Label18 
            Caption         =   "Cargos"
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
            Left            =   4440
            TabIndex        =   84
            Top             =   960
            Width           =   972
         End
         Begin VB.Label lblMCargos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   5880
            TabIndex        =   83
            Top             =   960
            Width           =   1812
         End
         Begin VB.Label Label18 
            Caption         =   "Pólizas"
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
            Left            =   4440
            TabIndex        =   82
            Top             =   600
            Width           =   972
         End
         Begin VB.Label lblMPolizas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   5880
            TabIndex        =   81
            Top             =   600
            Width           =   1812
         End
         Begin VB.Label lblMLineaDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   240
            TabIndex        =   80
            Top             =   2880
            Width           =   3492
         End
         Begin VB.Label Label14 
            Caption         =   "Abono"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   252
            Index           =   1
            Left            =   4440
            TabIndex        =   72
            Top             =   3000
            Width           =   1332
         End
         Begin VB.Label lblMMoraIntMor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   5880
            TabIndex        =   71
            Top             =   1680
            Width           =   1812
         End
         Begin VB.Label Label14 
            Caption         =   "Int.Mor."
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
            Left            =   4440
            TabIndex        =   70
            Top             =   1680
            Width           =   972
         End
         Begin VB.Label lblMMorIntCor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   5880
            TabIndex        =   69
            Top             =   1320
            Width           =   1812
         End
         Begin VB.Label Label13 
            Caption         =   "Int.Cor."
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
            Left            =   4440
            TabIndex        =   68
            Top             =   1320
            Width           =   972
         End
         Begin VB.Label lblMMorPrincipal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   5880
            TabIndex        =   67
            Top             =   2040
            Width           =   1812
         End
         Begin VB.Label Label12 
            Caption         =   "Principal"
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
            Left            =   4440
            TabIndex        =   66
            Top             =   2040
            Width           =   1092
         End
         Begin VB.Label lblMSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   5880
            TabIndex        =   65
            Top             =   240
            Width           =   1812
         End
         Begin VB.Label Label11 
            Caption         =   "Saldo"
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
            Left            =   4440
            TabIndex        =   64
            Top             =   240
            Width           =   972
         End
         Begin VB.Label lblMTotalDeuda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   5880
            TabIndex        =   63
            Top             =   2640
            Width           =   1812
         End
         Begin VB.Label Label10 
            Caption         =   "Total Deuda"
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
            Index           =   0
            Left            =   4440
            TabIndex        =   62
            Top             =   2640
            Width           =   1332
         End
         Begin VB.Label lblMGarantia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   240
            TabIndex        =   61
            Top             =   2520
            Width           =   3492
         End
         Begin VB.Label Label9 
            Caption         =   "Garantía / Línea"
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
            TabIndex        =   60
            Top             =   2280
            Width           =   1452
         End
         Begin VB.Label lblMTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   1560
            TabIndex        =   59
            Top             =   960
            Width           =   1932
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo"
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
            TabIndex        =   58
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblMCodigo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   1560
            TabIndex        =   57
            Top             =   600
            Width           =   1932
         End
         Begin VB.Label Label7 
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
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblMOperacion 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   312
            Left            =   1560
            TabIndex        =   55
            Top             =   240
            Width           =   1932
         End
         Begin VB.Label Label6 
            Caption         =   "Operación"
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
            TabIndex        =   54
            Top             =   240
            Width           =   975
         End
      End
      Begin XtremeSuiteControls.GroupBox fraObservacion 
         Height          =   5295
         Left            =   -70000
         TabIndex        =   78
         Top             =   360
         Visible         =   0   'False
         Width           =   10095
         _Version        =   1441793
         _ExtentX        =   17806
         _ExtentY        =   9340
         _StockProps     =   79
         Caption         =   "Digite la Observación de la Liquidación: "
         ForeColor       =   0
         BackColor       =   -2147483633
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnNotasCierra 
            Height          =   615
            Left            =   8040
            TabIndex        =   79
            Top             =   4440
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Cerrar Notas"
            ForeColor       =   0
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
            Picture         =   "frmAF_LiquidacionW.frx":1B40
            TextImageRelation=   4
         End
         Begin XtremeSuiteControls.FlatEdit txtObservacion 
            Height          =   3975
            Left            =   0
            TabIndex        =   114
            Top             =   360
            Width           =   10095
            _Version        =   1441793
            _ExtentX        =   17806
            _ExtentY        =   7011
            _StockProps     =   77
            ForeColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   2
         End
         Begin XtremeShortcutBar.ShortcutCaption lblObservacion 
            Height          =   375
            Left            =   0
            TabIndex        =   113
            Top             =   0
            Width           =   10095
            _Version        =   1441793
            _ExtentX        =   17806
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Anotaciones para esta Liquidación"
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Left            =   3960
         TabIndex        =   87
         Top             =   600
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9340
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
         TabIndex        =   88
         Top             =   600
         Width           =   1815
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFndRendLiquidar 
         Height          =   315
         Left            =   -67480
         TabIndex        =   96
         Top             =   4800
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFndRendGravado 
         Height          =   315
         Left            =   -67480
         TabIndex        =   97
         Top             =   5160
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkAplExcedente 
         Height          =   252
         Left            =   -69520
         TabIndex        =   98
         Top             =   2640
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica Excedente del Periodo"
         ForeColor       =   0
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
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   312
         Left            =   -64840
         TabIndex        =   105
         ToolTipText     =   "Divisa Origen"
         Top             =   5040
         Visible         =   0   'False
         Width           =   612
         _Version        =   1441793
         _ExtentX        =   1080
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
         Text            =   "COL"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTipoCambio 
         Height          =   312
         Left            =   -66520
         TabIndex        =   108
         ToolTipText     =   "Tipo de Cambio"
         Top             =   5040
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441793
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisaLocal 
         Height          =   312
         Left            =   -64840
         TabIndex        =   110
         ToolTipText     =   "Divisa Convertida"
         Top             =   4680
         Visible         =   0   'False
         Width           =   612
         _Version        =   1441793
         _ExtentX        =   1080
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
         Text            =   "COL"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdDistribucionAuto 
         Height          =   495
         Left            =   -66160
         TabIndex        =   120
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Distribución Automática"
         ForeColor       =   0
         BackColor       =   -2147483633
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
         Picture         =   "frmAF_LiquidacionW.frx":2321
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   510
         Left            =   -69880
         TabIndex        =   119
         Top             =   360
         Visible         =   0   'False
         Width           =   3975
         _Version        =   1441793
         _ExtentX        =   7011
         _ExtentY        =   900
         _StockProps     =   79
         Caption         =   "Indique o Modifique con Doble Click el Abono a Cada Operación"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLsw 
         Height          =   495
         Left            =   -34720
         TabIndex        =   118
         Top             =   -14280
         Visible         =   0   'False
         Width           =   3495
         _Version        =   1441793
         _ExtentX        =   6165
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin VB.Label Label4 
         Caption         =   "Tipo de Cambio/ Divisa Origen"
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
         Left            =   -69040
         TabIndex        =   109
         Top             =   5040
         Visible         =   0   'False
         Width           =   2532
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulos 
         Height          =   375
         Index           =   1
         Left            =   -70000
         TabIndex        =   107
         Top             =   360
         Visible         =   0   'False
         Width           =   10215
         _Version        =   1441793
         _ExtentX        =   18018
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Marque los Planes de Ahorros a Liquidar"
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
      Begin XtremeShortcutBar.ShortcutCaption scTitulos 
         Height          =   375
         Index           =   0
         Left            =   -70000
         TabIndex        =   106
         Top             =   360
         Visible         =   0   'False
         Width           =   10095
         _Version        =   1441793
         _ExtentX        =   17806
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Indique los Aportes a Utilizar en Abonos a Deudas"
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
      Begin VB.Label lblCustodia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -66520
         TabIndex        =   102
         Top             =   1800
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label lblExcedente 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -66520
         TabIndex        =   101
         Top             =   2640
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label lblExcedenteRenta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -64840
         TabIndex        =   100
         Top             =   2640
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I.R."
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
         Height          =   312
         Index           =   2
         Left            =   -63160
         TabIndex        =   99
         ToolTipText     =   "Impuesto de Renta"
         Top             =   2640
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Rendimiento Gravado:"
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
         Left            =   -70000
         TabIndex        =   95
         Top             =   5160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Rendimiento a Liquidar:"
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
         Index           =   6
         Left            =   -70000
         TabIndex        =   94
         Top             =   4800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Impuesto Renta s/Rend:"
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
         Index           =   3
         Left            =   -65080
         TabIndex        =   93
         Top             =   4800
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblTotalNeto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Index           =   2
         Left            =   -62200
         TabIndex        =   92
         Top             =   4800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I.R."
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
         Height          =   312
         Index           =   1
         Left            =   -63160
         TabIndex        =   91
         ToolTipText     =   "Impuesto de Renta"
         Top             =   2280
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.Label lblRenta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -64840
         TabIndex        =   90
         Top             =   2280
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label lblTotalNeto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Index           =   1
         Left            =   -62200
         TabIndex        =   45
         Top             =   5160
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Neto Disponible:"
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
         Index           =   5
         Left            =   -65080
         TabIndex        =   44
         Top             =   5160
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblAporteObrero 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -66520
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label lblAportePatronal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -66520
         TabIndex        =   37
         Top             =   1440
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label lblFCI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -64840
         TabIndex        =   36
         Top             =   1440
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label lblCapitalizacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -66520
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label lblAporteExtra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -66520
         TabIndex        =   34
         Top             =   3120
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.C.I."
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
         Height          =   312
         Index           =   0
         Left            =   -63160
         TabIndex        =   33
         ToolTipText     =   "Fondo de Capitalización Individual"
         Top             =   1440
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.Label Label4 
         Caption         =   "Total Bruto Disponible"
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
         Left            =   -69040
         TabIndex        =   32
         Top             =   3720
         Visible         =   0   'False
         Width           =   2412
      End
      Begin VB.Label lblTotalBruto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   -66520
         TabIndex        =   31
         Top             =   3720
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label4 
         Caption         =   "Retener Monto por la Suma de"
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
         Left            =   -69040
         TabIndex        =   30
         Top             =   4080
         Visible         =   0   'False
         Width           =   2412
      End
      Begin VB.Label Label4 
         Caption         =   "Total Neto Disponible"
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
         Left            =   -69040
         TabIndex        =   29
         Top             =   4680
         Visible         =   0   'False
         Width           =   2412
      End
      Begin VB.Label lblTotalNeto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   0
         Left            =   -66520
         TabIndex        =   28
         Top             =   4680
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label4 
         Caption         =   "Impuesto Renta [Pendiente] sobre Capitalización + Adelanto Excedentes"
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
         Left            =   -64720
         TabIndex        =   27
         Top             =   4080
         Visible         =   0   'False
         Width           =   3252
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Identificacion"
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
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionW.frx":2A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionW.frx":929C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionW.frx":FAFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionW.frx":16360
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionW.frx":1CBC2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnSiguiente 
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Siguiente"
      ForeColor       =   0
      BackColor       =   16777215
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
      Picture         =   "frmAF_LiquidacionW.frx":23424
      TextImageRelation=   4
   End
   Begin XtremeSuiteControls.PushButton btnAnterior 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Anterior"
      ForeColor       =   0
      BackColor       =   16777215
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
      Picture         =   "frmAF_LiquidacionW.frx":23CDF
   End
   Begin XtremeSuiteControls.FlatEdit txtAsientoNo 
      Height          =   252
      Left            =   7200
      TabIndex        =   85
      Top             =   1320
      Visible         =   0   'False
      Width           =   732
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   444
      _StockProps     =   77
      ForeColor       =   0
      Text            =   "1"
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Liquidación de la Persona"
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
      Height          =   495
      Index           =   12
      Left            =   1880
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmAF_LiquidacionW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mFechaSistema As Date, vControlRenuncias As Boolean
Dim vPaso As Boolean, vTipoDoc As String
Dim vModoSIF As Boolean, vConcepto As String
Dim nDocumento As String
Dim iPromotor As Integer, iAplicaReIngreso As Integer

Private Function fxVerificaDatos() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""

strSQL = "select isnull(count(*),0) as Existe from socios where cedula = '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then vMensaje = vMensaje & " - La persona especificada no existe Registrada..." & vbCrLf
rs.Close

If (cboTipoDoc.ItemData(cboTipoDoc.ListIndex) = "TE" Or cboTipoDoc.ItemData(cboTipoDoc.ListIndex) = "TS") And (cboCuenta.ListCount = 0 Or cboCuenta.Text = "") Then
    vMensaje = vMensaje & " - No se ha indicado una cuenta bancaria para realizar la transferencia a la persona..." & vbCrLf

End If

If cboCausa.ListCount = 0 Then
    vMensaje = vMensaje & " - No se ha indicado una causa de reuncia..." & vbCrLf
End If

If Mid(cboTipo, 1, 2) = "03" Or Mid(cboTipo, 1, 2) = "" Then vMensaje = vMensaje & " - El proceso siguiente no aplica..." & vbCrLf


If lblAc_Boleta.Visible Then
  If Len(txtAc_Boleta.Text) = 0 Then
    vMensaje = vMensaje & " - Especifique el Número de Boleta de Accion de Personal..." & vbCrLf
  End If
End If

If Len(vMensaje) > 0 Then
 MsgBox vMensaje, vbCritical
 fxVerificaDatos = False
Else
 fxVerificaDatos = True
End If

End Function


Private Sub sbAportesTotales()
Dim curTotalBruto As Currency, curRenta As Currency

curTotalBruto = 0

If chkAplObrero.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblAporteObrero.Caption)
If chkAplPatronal.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblAportePatronal.Caption)
If chkAplPatronal.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblCustodia.Caption)

If chkAplPatronal.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblFCI.Caption)
If chkAplCapGen.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblCapitalizacion.Caption)
If chkAplCapExtra.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblAporteExtra.Caption)
If chkAplExcedente.Value = xtpChecked Then curTotalBruto = curTotalBruto + CCur(lblExcedente.Caption)


lblTotalBruto.Caption = Format(curTotalBruto, "Standard")

curRenta = CCur(lblRenta.Caption)

If chkAplExcedente.Value = xtpChecked Then
    curRenta = curRenta + CCur(lblExcedenteRenta.Caption)
End If

txtRetenerMonto = Format(curRenta, "Standard")

If IsNumeric(txtRetenerMonto) Then
  lblTotalNeto(0).Caption = Format(curTotalBruto - CCur(txtRetenerMonto), "Standard")
Else
  lblTotalNeto(0).Caption = Format(curTotalBruto, "Standard")
End If

End Sub



Private Sub btnAnterior_Click()
Dim i As Integer

If tcMain.SelectedItem > 0 Then
 tcMain.Item(tcMain.SelectedItem - 1).Selected = True
 
 For i = 0 To tcMain.ItemCount - 1
   tcMain.Item(i).Enabled = False
 Next i
 'Preguntar si desea limpiar los datos
 i = MsgBox("Desea Limpiar Los Datos Anteriores...", vbYesNo)
 If i = vbYes Then
   Call sbLimpiaDatos
   Call sbCargaDatos
 End If
End If

tcMain.Item(tcMain.SelectedItem).Enabled = True
        
End Sub

Private Sub btnNotasCierra_Click()
    fraPrg.Visible = False
    fraObservacion.Visible = False
    
    fraResumen.Visible = True
End Sub

Private Sub btnSiguiente_Click()
Dim i As Integer
           

If tcMain.SelectedItem = 0 Then
    If fxVerificaDatos Then
       tcMain.Item(tcMain.SelectedItem + 1).Selected = True
       Call sbLimpiaDatos
       Call sbCargaDatos
    End If
Else
    If tcMain.SelectedItem < 4 Then
       tcMain.Item(tcMain.SelectedItem + 1).Selected = True
      Call sbLimpiaDatos
      Call sbCargaDatos
    End If
End If

End Sub

Private Sub cboBanco_Click()

If vPaso Or cboBanco.ListCount = 0 Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & txtCedula.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

strSQL = "exec spAFI_Renuncia_Emite_TDoc " & cboBanco.ItemData(cboBanco.ListIndex) & ""
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)

vError:

End Sub

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCuenta.SetFocus
End Sub

Private Sub cboCausa_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim x As Integer

If vPaso Then Exit Sub
If cboCausa.ListCount = 0 Then Exit Sub

chkMortalidad.Value = vbUnchecked
chkAltPlanilla.Value = vbChecked

chkMortalidad.Enabled = False
chkAltPlanilla.Enabled = False

x = cboCausa.ItemData(cboCausa.ListIndex)

strSQL = "select mortalidad,liq_alterna from causas_renuncias where id_causa = " & x
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  If rs!Mortalidad = 1 Then
     chkMortalidad.Enabled = True
     chkMortalidad.Value = vbChecked
  End If
  
  If rs!liq_Alterna = 1 Then
     chkAltPlanilla.Value = vbChecked
     chkAltPlanilla.Enabled = True
  End If
End If
rs.Close

End Sub

Private Sub cboCausa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoDoc.SetFocus
End Sub

Private Sub cboTipo_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If Mid(cboTipo.Text, 1, 2) = "01" Then
    dtpAc_fecha.Visible = False
    txtAc_Boleta.Visible = False
    cboAc_Tipo.Visible = False
Else
    dtpAc_fecha.Visible = True
    txtAc_Boleta.Visible = True
    cboAc_Tipo.Visible = True
End If

lblAc_Boleta.Visible = txtAc_Boleta.Visible
lblAC_Tipo.Visible = cboAc_Tipo.Visible


vPaso = True
    'Carga Causas
    strSQL = "select id_Causa as 'IdX', Descripcion as 'ItmX'" _
           & " from causas_renuncias WHERE ACTIVO = 1" _
           & " and Tipo_Apl in('A', '" & IIf((Mid(cboTipo.Text, 1, 2) = "01"), "I", "P") & "')"
    Call sbCbo_Llena_New(cboCausa, strSQL, False, True)
vPaso = False
Call cboCausa_Click

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCausa.SetFocus
End Sub

Private Sub cboTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboBanco.SetFocus
End Sub

Private Sub chkAplCapExtra_Click()
Call sbAportesTotales
End Sub

Private Sub chkAplCapGen_Click()
Call sbAportesTotales
End Sub

Private Sub chkAplObrero_Click()
Call sbAportesTotales
End Sub

Private Sub chkAplPatronal_Click()
Call sbAportesTotales
End Sub

Private Sub sbLiqP1Registro()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curTotalLiq As Currency, curTotalPrestamos As Currency
Dim curAhorroLiq As Currency, curAporteLiq As Currency
Dim curExtraLiq As Currency, curCapitalLiq As Currency
Dim curFCILiq As Currency, vUbicacion As String
Dim curExcedente As Currency, curExcedenteIR As Currency
Dim curExcLiq As Currency, curExcIRLiq As Currency, curCustodiaLiq As Currency

Dim vOficina As String, i As Integer


'Actualización del Código de Oficina.

strSQL = "select cod_oficina from socios where cedula = '" & Trim(txtCedula) & "'"
Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF Then
    vOficina = IIf(IsNull(rs!COD_OFICINA), "", rs!COD_OFICINA)
  End If
rs.Close


'Ingresa Registro de Renuncia
If vModoSIF Then
    ' Codigo SIF
    strSQL = "Insert into Renuncias(ID_Causa,Cedula,fecha,tipo)" _
           & " values(" & cboCausa.ItemData(cboCausa.ListIndex) _
           & ",'" & Trim(txtCedula) & "','"
    
    Select Case Mid(cboTipo, 1, 2)
      Case "01" 'Ren.Interna
         strSQL = strSQL & Format(mFechaSistema, "yyyy/mm/dd  hh:mm:ss") & "','A')"
      Case "02" 'Ren.Total
         strSQL = strSQL & Format(mFechaSistema, "yyyy/mm/dd  hh:mm:ss") & "','P')"
    End Select

Else
    'Codigo ASE
    strSQL = "Select Cedula,Nacta,id_Promotor,Id_Boleta_Af" _
           & " from Socios Where Cedula='" & Trim(txtCedula) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    strSQL = "Insert into Renuncias(ID_Causa,ID_Promotor,Cedula,Id_Boleta,"
    
    Select Case Mid(cboTipo, 1, 2)
      Case "01" 'Ren.Interna
         strSQL = strSQL & "FechaRenA,"
      Case "02" 'Ren.Total
       strSQL = strSQL & "FechaRenP,"
    End Select
    
    strSQL = strSQL & "TipoRen,Nacta,NCausaRen,RenMor) values(" & cboCausa.ItemData(cboCausa.ListIndex) _
           & "," & rs!ID_PROMOTOR & ",'" & Trim(txtCedula) & "'," & rs!id_Boleta_AF & ",'"
    
    Select Case Mid(cboTipo, 1, 2)
      Case "01" 'Ren.Interna
         strSQL = strSQL & Format(mFechaSistema, "yyyy/mm/dd  hh:mm:ss") & "','A',"
      Case "02" 'Ren.Total
         strSQL = strSQL & Format(mFechaSistema, "yyyy/mm/dd  hh:mm:ss") & "','P',"
    End Select
    
    strSQL = strSQL & CLng(IIf(IsNull(rs!Nacta), 0, rs!Nacta)) & ",0," & chkMortalidad.Value & ")"

End If

Call ConectionExecute(strSQL)
'--- Fin Registro de Renuncia


'Ingresa Registro Maestro de la Liquidacion
'        EXCEDENTE_PERIODO DEC(14,2) DEFAULT 0
'       ,EXCEDENTE_IR      DEC(12,2) DEFAULT 0
'       ,EXCEDENTE_LIQ     DEC(12,2) DEFAULT 0
'       ,EXCEDENTE_IR_LIQ  DEC(12,2) DEFAULT 0
'       ,APL_EXCEDENTE     SMALLINT

strSQL = "insert liquidacion(cedula, ahorro, aporte, custodia, capitaliza, extra, montofci, retenido" _
       & ", fechaingreso, fecliq, estadoActLiq, estadoactual, aplAhorro, aplAporte, aplCapitalizado" _
       & ", aplExtra, TotalBruto, TNeto, Ahorro_Liq, Aporte_liq, custodia_liq, Capitalizado_liq, extra_liq" _
       & ", tdocumento, cod_banco, CTA_AHORROS, ubicacion, liq_tcon, estadoAsiento, estado, mortalidad" _
       & ", id_causa, observacion, usuario, cod_oficina, ac_boleta, ac_fecha, fecha_pago" _
       & ", EXCEDENTE_PERIODO, EXCEDENTE_IR, EXCEDENTE_LIQ, EXCEDENTE_IR_LIQ, APL_EXCEDENTE, COD_DIVISA, TIPO_CAMBIO)" _
       & " values('" & Trim(txtCedula.Text) & "'," & CCur(lblAporteObrero.Caption) & "," & CCur(lblAportePatronal.Caption) & "," & CCur(lblCustodia.Caption) _
       & "," & CCur(lblCapitalizacion.Caption) & "," & CCur(lblAporteExtra.Caption) & "," & CCur(lblFCI.Caption) _
       & "," & CCur(txtRetenerMonto.Text) & ",'" & Format(lblIngreso.Caption, "yyyy/mm/dd") & "',dbo.MyGetdate()" _
       & ",'" & IIf(Mid(cboTipo, 1, 2) = "01", "A", "P") & "','" & lblEstadoActual.Tag _
       & "'," & chkAplObrero.Value & "," & chkAplPatronal.Value & "," & chkAplCapGen.Value _
       & "," & chkAplCapExtra.Value & "," & CCur(lblTotalNeto(1).Caption) & ","

Select Case Mid(cboTipo.Text, 1, 2)
  Case "01" 'Liq.Interna
    curTotalLiq = CCur(lblAporteExtra.Caption) + CCur(lblAporteObrero.Caption) _
                + CCur(lblCapitalizacion.Caption)
  Case "02" 'Liq.Total
    curTotalLiq = CCur(lblAporteExtra.Caption) + CCur(lblAporteObrero.Caption) _
                + CCur(lblCapitalizacion.Caption) + CCur(lblAportePatronal.Caption) _
                + CCur(lblFCI.Caption) + CCur(lblCustodia.Caption)
    
    If lblExcedente.Tag = "1" And chkAplExcedente.Value = xtpChecked Then
        curTotalLiq = curTotalLiq + CCur(lblExcedente.Caption)
        curExcedente = CCur(lblExcedente.Caption)
        curExcedenteIR = CCur(lblExcedenteRenta.Caption)
    End If
End Select

'Suma Planes Liquidados
With lswPlanes.ListItems
    For i = 1 To .Count
      If .Item(i).Checked Then
          curTotalLiq = curTotalLiq + .Item(i).SubItems(2)
      End If
    Next i
End With


curTotalLiq = curTotalLiq - CCur(txtRetenerMonto)
curTotalPrestamos = CCur(lblTotalNeto(1).Caption) - CCur(lblDisponible.Caption)

'Monto A Girar
strSQL = strSQL & (curTotalLiq - curTotalPrestamos) & ","

If (curTotalLiq - curTotalPrestamos) > 0 Then
 vUbicacion = "T"
Else
 vUbicacion = "C"
End If

'Aportes Liquidados
curAhorroLiq = CCur(lblAporteObrero.Caption)
curAporteLiq = CCur(lblAportePatronal.Caption)
curCustodiaLiq = CCur(lblCustodia.Caption)
curCapitalLiq = CCur(lblCapitalizacion.Caption)


curExtraLiq = CCur(lblAporteExtra.Caption)
curFCILiq = CCur(lblFCI.Caption)

If lblExcedente.Tag = "1" And chkAplExcedente.Value = xtpChecked Then
    curExcLiq = curExcedente
    curExcIRLiq = curExcedenteIR
Else
    curExcLiq = 0
    curExcIRLiq = 0
End If


'Ahorro Obrero Remanente
If curTotalPrestamos > 0 Then
    If chkAplObrero.Value = vbChecked Then
     If curAhorroLiq >= curTotalPrestamos Then
        curAhorroLiq = curAhorroLiq - curTotalPrestamos
        curTotalPrestamos = 0
     Else
        curTotalPrestamos = curTotalPrestamos - curAhorroLiq
        curAhorroLiq = 0
     End If
    End If
End If

'Aporte Patronal Remanente
If curTotalPrestamos > 0 Then
    If chkAplPatronal.Value = vbChecked Then
     If curAporteLiq >= curTotalPrestamos Then
        curAporteLiq = curAporteLiq - curTotalPrestamos
        curTotalPrestamos = 0
     Else
        curTotalPrestamos = curTotalPrestamos - curAporteLiq
        curAporteLiq = 0
     End If
    End If
End If

'Aporte Patronal Custodia Remanente
If curTotalPrestamos > 0 Then
    If chkAplPatronal.Value = vbChecked Then
     If curCustodiaLiq >= curTotalPrestamos Then
        curCustodiaLiq = curCustodiaLiq - curTotalPrestamos
        curTotalPrestamos = 0
     Else
        curTotalPrestamos = curTotalPrestamos - curCustodiaLiq
        curCustodiaLiq = 0
     End If
    End If
End If



'FCI Remanente
If curTotalPrestamos > 0 Then
    If chkAplPatronal.Value = vbChecked Then
     If curFCILiq >= curTotalPrestamos Then
        curFCILiq = curFCILiq - curTotalPrestamos
        curTotalPrestamos = 0
     Else
        curTotalPrestamos = curTotalPrestamos - curFCILiq
        curFCILiq = 0
     End If
    End If
End If


'Cap.General Remanente
If curTotalPrestamos > 0 Then
    If chkAplCapGen.Value = vbChecked Then
     If curCapitalLiq >= curTotalPrestamos Then
        curCapitalLiq = curCapitalLiq - curTotalPrestamos
        curTotalPrestamos = 0
     Else
        curTotalPrestamos = curTotalPrestamos - curCapitalLiq
        curCapitalLiq = 0
     End If
    End If
End If

'Ahorro Extra Remanente
If curTotalPrestamos > 0 Then
    If chkAplCapExtra.Value = vbChecked Then
     If curExtraLiq >= curTotalPrestamos Then
        curExtraLiq = curExtraLiq - curTotalPrestamos
        curTotalPrestamos = 0
     Else
        curTotalPrestamos = curTotalPrestamos - curExtraLiq
        curExtraLiq = 0
     End If
    End If
End If

Dim pTipoPago As String

pTipoPago = cboTipoDoc.ItemData(cboTipoDoc.ListIndex)

'Inserta Remanentes y Finaliza la Clausula Insert del Maestro de Liquidacion
strSQL = strSQL & curAhorroLiq & "," & curAporteLiq & "," & curCustodiaLiq & "," & curCapitalLiq _
       & "," & curExtraLiq & ",'" & pTipoPago & "'," _
       & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & cboCuenta.ItemData(cboCuenta.ListIndex) & "','" & vUbicacion & "',5,'" _
       & IIf((chkMortalidad.Value = vbChecked), "G", "P") & "','P'," & chkMortalidad.Value _
       & "," & cboCausa.ItemData(cboCausa.ListIndex) & ",'" & Mid(Trim(txtObservacion), 1, 250) _
       & "','" & glogon.Usuario & "','" & vOficina & "'," & IIf(txtAc_Boleta.Visible, ("'" & txtAc_Boleta.Text & "'"), "Null") _
       & "," & IIf(dtpAc_fecha.Visible, ("'" & Format(dtpAc_fecha.Value, "yyyy/mm/dd") & "'"), "Null") _
       & ",'" & Format(dtpPago.Value, "yyyy/mm/dd") & "', " & curExcedente & ", " _
       & curExcedenteIR & ", " & curExcLiq & ", " & curExcIRLiq & ", " & chkAplExcedente.Value _
       & ", '" & txtDivisa.Text & "', " & CCur(txtTipoCambio.Text) & ")"
Call ConectionExecute(strSQL)

End Sub




Private Sub sbLiqP2Planes(vLiq As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

strSQL = ""

'Inserta los registros de Planes de Ahorros a Liquidar
With lswPlanes.ListItems
  For i = 1 To .Count
    If .Item(i).Checked Then
        strSQL = strSQL & Space(10) & "insert into LIQUIDA_FONDOS(CONSEC, COD_CONTRATO, COD_OPERADORA, COD_PLAN, DISPONIBLE, MULTA" _
               & ", REND_PENDIENTE, LIQ_FND, APORTES, RENDIMIENTOS, COD_DIVISA, TIPO_CAMBIO)" _
               & " values(" & vLiq & "," & CLng(.Item(i).Text) & "," & CLng(.Item(i).Tag) & ",'" & .Item(i).SubItems(1) _
               & "'," & CCur(.Item(i).SubItems(2)) & "," & CCur(.Item(i).SubItems(6)) & "," & CCur(.Item(i).SubItems(5)) _
               & ",0," & CCur(.Item(i).SubItems(3)) & "," & CCur(.Item(i).SubItems(4)) & ",'" & Trim(.Item(i).SubItems(12)) _
               & "'," & CCur(.Item(i).SubItems(13)) & ")"
    End If
  Next i
End With

If Len(strSQL) > 0 Then
    'Registra Lista a Procesar
    Call ConectionExecute(strSQL)
    
    'Procesa Liquidaciones de Planes
    strSQL = "exec spAfiLiquidaPlanes " & vLiq & ",'" & glogon.Usuario & "','" & GLOBALES.gOficinaTitular & "'"
    Call ConectionExecute(strSQL)
End If

End Sub

Private Sub sbLiqP2Aportes(vLiq As Long)
Dim strSQL As String

strSQL = "exec spAfi_Liquidacion_Patrimonio " & vLiq & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

End Sub

Private Sub sbLiqP3Creditos(vLiq As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, curAbono As Currency, curPrincipal As Currency
Dim curAbIntc As Currency, curAbIntm As Currency, curAbAmortiza As Currency
Dim curInteresCor As Currency, curInteresMor As Currency

With lswAbonos.ListItems
  For i = 1 To .Count
   curInteresCor = 0
   curInteresMor = 0
   curPrincipal = 0
   curAbono = CCur(.Item(i).SubItems(11))
   
   
     'Aplica Sin Plan de Pagos
     'Procesa Morosidad
     If (CCur(.Item(i).SubItems(6)) + CCur(.Item(i).SubItems(7)) _
        + CCur(.Item(i).SubItems(8))) <= curAbono Then  'Esta Moroso
        
        curAbono = curAbono - (CCur(.Item(i).SubItems(6)) + CCur(.Item(i).SubItems(7)) _
        + CCur(.Item(i).SubItems(8)))
        
        curInteresCor = curInteresCor + CCur(.Item(i).SubItems(6))
        curInteresMor = curInteresMor + CCur(.Item(i).SubItems(7))
        curPrincipal = curPrincipal + CCur(.Item(i).SubItems(8))
        
        strSQL = "update morosidad set abintc = intc,abintm = intm,abamortiza = amortiza" _
               & ",estado = 'C',Tcon = '" & vTipoDoc & "',Ncon = '" & vLiq _
               & "',fecult = dbo.MyGetdate(),cod_concepto = 'CRD001',cod_caja = '', usuario = '" & glogon.Usuario & "'" _
               & " where estado = 'A' and id_solicitud = " & .Item(i).Text
        Call ConectionExecute(strSQL)
        
      Else 'el Abono no es Mayor hay que procesar uno por uno
      
        strSQL = "select * from morosidad where estado = 'A' and " _
               & " id_solicitud = " & .Item(i).Text
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
         If curAbono >= (rs!IntC + rs!IntM + rs!Amortiza) Then
            curAbono = curAbono - (rs!IntC + rs!IntM + rs!Amortiza)
            
            curPrincipal = curPrincipal + rs!Amortiza
            curInteresCor = curInteresCor + rs!IntC
            curInteresMor = curInteresMor + rs!IntM
            
            strSQL = "update morosidad set abintc = intc,abintm = intm,abamortiza = amortiza" _
                   & ",estado = 'C',Tcon = '" & vTipoDoc & "',Ncon = '" & vLiq _
                   & "',fecult = dbo.MyGetdate()" _
                   & ",usuario = '" & glogon.Usuario & "', Cod_Caja = '', cod_Concepto = 'CRD001'" _
                   & " where id_moro = " & rs!id_moro
            Call ConectionExecute(strSQL)
         
         Else
           'Distribuir Abono
           curAbAmortiza = 0
           curAbIntc = 0
           curAbIntm = 0
           
           If curAbono >= rs!IntC Then
              curAbono = curAbono - rs!IntC
              curAbIntc = rs!IntC
           Else
              curAbIntc = curAbono
              curAbono = 0
           End If
           
           If curAbono >= rs!IntM Then
              curAbono = curAbono - rs!IntM
              curAbIntm = rs!IntM
           Else
              curAbIntm = curAbono
              curAbono = 0
           End If
           
           If curAbono >= rs!Amortiza Then
              curAbono = curAbono - rs!Amortiza
              curAbAmortiza = rs!Amortiza
           Else
              curAbAmortiza = curAbono
              curAbono = 0
           End If
           
           'Totales
           curPrincipal = curPrincipal + curAbAmortiza
           curInteresCor = curInteresCor + curAbIntc
           curInteresMor = curInteresMor + curAbIntm
           
           
           If (curAbAmortiza + curAbIntc + curAbIntm) > 0 Then
              strSQL = "update morosidad set abintc = " & curAbIntc & ",abintm = " _
                     & curAbIntm & ",abamortiza = " & curAbAmortiza _
                     & ",estado = 'C',Tcon = '" & vTipoDoc & "',Ncon = '" & vLiq _
                     & "',fecult = dbo.MyGetdate()" _
                     & ",usuario = '" & glogon.Usuario & "', Cod_Caja = '', cod_Concepto = 'CRD001'" _
                     & " where id_moro = " & rs!id_moro
              Call ConectionExecute(strSQL)
              
              'Registrar La diferencia
              strSQL = "insert morosidad(id_solicitud,codigo,estado,estadoi,intc,intm,amortiza" _
                     & ",abintc,abintm,abamortiza,tcon,ncon,fechap,fecap,fecult,cuota_morosa,usuario,cod_concepto,cod_caja) values(" _
                     & rs!Id_Solicitud & ",'" & Trim(rs!Codigo) & "','A','A'," & rs!IntC - curAbIntc _
                     & "," & rs!IntM - curAbIntm & "," & rs!Amortiza - curAbAmortiza & ",0,0,0,'" & vTipoDoc & "','" _
                     & vLiq & "'," & rs!fechap & "," & rs!fecap & ",dbo.MyGetdate()," & (rs!IntM + rs!IntC + rs!Amortiza) - (curAbAmortiza + curAbIntc + curAbIntm) _
                     & ",'" & glogon.Usuario & "','CRD001','')"
              Call ConectionExecute(strSQL)
           
           End If 'Fin Distribucion
           
         End If 'Mora Entera
         rs.MoveNext
        Loop
        rs.Close
        
     End If ' Fin de la Morosidad
     
     
     'Procesa Cuota ExtraOrdinaria
     If curAbono > 0 Then
        curPrincipal = curPrincipal + curAbono
        strSQL = "insert creditos_dt(id_solicitud,codigo,cuota,abono,intcp,amortiza,fechap,fechas,tcon,ncon" _
               & ",estado,usuario,cod_concepto,cod_caja) values(" & .Item(i).Text & ",'" & .Item(i).SubItems(1) & "',0," & curAbono _
               & ",0," & curAbono & "," & GLOBALES.glngFechaCR & ",dbo.MyGetdate(),'" & vTipoDoc & "','" & vLiq & "','A','" & glogon.Usuario & "','CRD002','')"
        Call ConectionExecute(strSQL)
        curAbono = 0
     End If
     
     'Actualiza Registro Maestro
     If (curPrincipal + curInteresCor + curInteresMor) > 0 Then
        Select Case Trim(.Item(i).SubItems(3))
          Case "Crédito Cartera"
              strSQL = "update reg_creditos set saldo = saldo - " & curPrincipal _
                     & ",amortiza = amortiza + " & curPrincipal & ",saldo_mes = saldo_mes - " _
                     & curPrincipal & ",interesc = interesc + " & curInteresCor + curInteresMor _
                     & ",estado = '" & IIf((CCur(.Item(i).SubItems(5)) - curPrincipal) <= 0, "C", "A") _
                     & "' where id_solicitud = " & .Item(i).Text
          Case "Reten. A Plazo"
              strSQL = "update reg_creditos set amortiza = amortiza + " & curPrincipal _
                     & ",interesc = interesc + " & curInteresCor + curInteresMor _
                     & ",estado = '" & IIf((CCur(.Item(i).SubItems(5)) - curPrincipal) <= 0, "C", "A") _
                     & "' where id_solicitud = " & .Item(i).Text
          Case "Reten.Indefinida"
              strSQL = "update reg_creditos set amortiza = amortiza + " & curPrincipal _
                     & ",interesc = interesc + " & curInteresCor + curInteresMor _
                     & " where id_solicitud = " & .Item(i).Text
          
        End Select
        Call ConectionExecute(strSQL)
     
     End If
    
     'Inserta Detalle del Abono en la Liquidación
      strSQL = "insert liquida_detalle(CONSEC,ID_SOLICITUD,CODIGO,LIQ_ABONO,LIQ_FECHA," _
             & "LIQ_SALDO,LIQ_INTCOR,LIQ_INTMOR,LIQ_AMORTIZA) values(" & vLiq _
             & "," & .Item(i).Text & ",'" & .Item(i).SubItems(1) & "'," & CCur(.Item(i).SubItems(11)) _
             & ",dbo.mygetdate()," & CCur(.Item(i).SubItems(5)) _
             & "," & curInteresCor & "," & curInteresMor & "," & curPrincipal & ")"
      Call ConectionExecute(strSQL)

  Next i

End With

End Sub


Private Sub sbLiqP3Creditos_PlanPagos(vLiq As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError_P3

'Paso 1: Inserta la Lista de Abonos a Procesar
strSQL = ""
With lswAbonos.ListItems
  For i = 1 To .Count
  
    strSQL = strSQL & Space(10) & "insert liquida_detalle(CONSEC,ID_SOLICITUD, CODIGO, LIQ_ABONO, LIQ_FECHA" _
           & ", LIQ_SALDO, LIQ_INTCOR, LIQ_INTMOR, LIQ_AMORTIZA, COD_DIVISA, TIPO_CAMBIO)" _
           & " values(" & vLiq _
           & "," & .Item(i).Text & ",'" & .Item(i).SubItems(1) & "'," & CCur(.Item(i).SubItems(11)) _
           & ",dbo.MyGetDate()," & CCur(.Item(i).SubItems(5)) _
           & "," & 0 & "," & 0 & "," & 0 & ",'" & Trim(.Item(i).SubItems(12)) & "'," & CCur(.Item(i).SubItems(13)) & ")"
  Next i

If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)

    'Paso 2: Aplica Proc de Abonos
    strSQL = "exec spAfi_Liquidacion_Abonos_PlanPagos " & vLiq & ", '" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
End If
End With



Exit Sub

vError_P3:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbLiqP4Asiento(vLiq As Long)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError_Asiento

strSQL = "exec spAFI_Liquidacion_Asiento " & vLiq
Call ConectionExecute(strSQL)

Exit Sub

vError_Asiento:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbLiqP5Traslados(vLiq As Long)
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spAFI_Liquidacion_Traslado_OpEx " & vLiq
Call ConectionExecute(strSQL)


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbLiqP6FondoDevoluciones(vLiq As Long)
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spAFI_Liquidacion_Fondos_Devolucion " & vLiq & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, lngLiq As Long, lngRenuncia As Long

On Error GoTo vError

'Verificar Congelamiento
If fxgCongelamiento(txtCedula, "per_liquidacion") Then
  MsgBox "Esta Persona se encuentra CONGELADA, verifique...", vbExclamation
  Exit Sub
End If

'Verificar si hay control de renuncias que se encuentre perdida y sin codigo de liquidacion

lngRenuncia = 0

If vControlRenuncias Then
   strSQL = "select Top 1 cod_renuncia from afi_cr_renuncias" _
          & " where liq is null and estado in('P','V') and cedula = '" _
          & txtCedula & "' order by cod_renuncia desc"
   Call OpenRecordSet(rs, strSQL)
   If rs.EOF And rs.BOF Then
      MsgBox "Esta persona no se encuentra registrada en control de Renuncias, o ya se liquidó...", vbExclamation
      rs.Close
      Exit Sub
   Else
     lngRenuncia = rs!Cod_Renuncia
   End If
   rs.Close
End If

' 0.preguntar si está seguro de aplicar la liquidacion
i = MsgBox("Esta seguro que desea aplicar esta liquidación ?", vbYesNo)
If i = vbNo Then Exit Sub

Call sbEstadoCuenta(txtCedula)

Me.MousePointer = vbHourglass

cmdAplicar.Enabled = False

fraPrg.Visible = True
PrgBar.Max = 7
PrgBar.Value = 1

lblX.Caption = "Procesando Registro Maestro Personas , Renuncia y Liquidación..."
lblX.Refresh

mFechaSistema = fxFechaServidor

'Pasos
'1. Generar Registro de Liquidacion
Call sbLiqP1Registro

'Recupera ID de liquidacion para los otros procesos
strSQL = "select isnull(max(consec),0) as Liq from liquidacion where estado = 'P' and cedula = '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
   lngLiq = rs!liq
rs.Close

'Actualiza el control de renuncias
If vControlRenuncias Then
   strSQL = "update afi_cr_renuncias set liq = " & lngLiq _
          & " where cod_renuncia = " & lngRenuncia
   Call ConectionExecute(strSQL)
End If

PrgBar.Value = 2
lblX.Caption = "Procesando Movimientos a las Cuentas de Patrimonio..."
lblX.Refresh

'2. Aplicar Movimientos Aportes
Call sbLiqP2Aportes(lngLiq)


PrgBar.Value = 3
lblX.Caption = "Liquidando Planes de Ahorros..."
lblX.Refresh

'2. Aplicar Liq. de Planes de Ahorros
Call sbLiqP2Planes(lngLiq)

PrgBar.Value = 4
lblX.Caption = "Procesando Movimientos de Abonos a Créditos..."
lblX.Refresh

'3. Aplicar Movimientos Creditos
If GLOBALES.SysPlanPagos = 1 Then
    Call sbLiqP3Creditos_PlanPagos(lngLiq)
Else
    Call sbLiqP3Creditos(lngLiq)
End If

PrgBar.Value = 5
lblX.Caption = "Procesando Asiento (1) Liquidación..."
lblX.Refresh

'4. Generar Asiento de Liq
Call sbLiqP4Asiento(lngLiq)

PrgBar.Value = 6
lblX.Caption = "Procesando Estados de los Préstamos y Traslados de Cuentas..."
lblX.Refresh

'5. Si era Socio Activo, Establecer Estado Opex, Subir la Tasa y Readecuar
'   la cuota de los prestamos, Finalizar Asiento con el traslado de Saldos a cuentas
'   y Cierra Asiento de Liquidacion
Call sbLiqP5Traslados(lngLiq)

lblX.Caption = "Procesando Fondos de Devoluciones..."
lblX.Refresh

Call sbLiqP6FondoDevoluciones(lngLiq)


PrgBar.Value = 7
lblX.Caption = "Procesando Reporte de la Liquidación..."
lblX.Refresh


Call Bitacora("Aplica", "Liquidación # " & lngLiq & " - Ced:" & txtCedula)

If vParametros.BitacoraEspecial Then
   Call sbgAFIBitacora("07", "Aplica Liquidación # " & lngLiq & " - Ced: " & txtCedula.Text, Trim(txtCedula.Text))
End If
          



'6. Generar reporte de liquidacion, Ojo con los traslados y Custodias.
Call sbgAFIBoletaLiquidacion(lngLiq)

'codigo para reactivacion
If lngRenuncia > 0 Then
    Call sbDatosRenuncia(lngRenuncia)
    
    If iAplicaReIngreso > 0 Then
       Call sbReIngreso(txtCedula.Text, fxInstitucion(txtCedula), iPromotor, 0, fxFechaServidor)
       strSQL = "Update liquidacion set aplica_reingreso = 1 where consec = " & lngLiq
       Call ConectionExecute(strSQL)
    End If
End If
fraPrg.Visible = False

Me.MousePointer = vbDefault

MsgBox "Liquidación Procesada Satisfactorimente...", vbInformation

'Inicializa Ventana
tcMain.Item(0).Selected = True

Call sbLimpiaDatos
cmdAplicar.Enabled = True

Exit Sub

vError:
  cmdAplicar.Enabled = True
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub cmdDistribucionAuto_Click()
Dim curDisponible As Currency, i As Integer
Dim curAbono As Currency
 
Me.MousePointer = vbHourglass
 
'Inicializa
curDisponible = lblTotalNeto(1).Caption

For i = 1 To lswAbonos.ListItems.Count
  lswAbonos.ListItems(i).SubItems(11) = 0
Next i

'Distribuye
With lswAbonos.ListItems
 For i = 1 To .Count
   
   curAbono = 0

   'Polizas
   If curDisponible > CCur(.Item(i).SubItems(10)) Then
     curAbono = curAbono + CCur(.Item(i).SubItems(10))
     curDisponible = curDisponible - CCur(.Item(i).SubItems(10))
   Else
     curAbono = curAbono + curDisponible
     curDisponible = 0
   End If

   'Cargos
   If curDisponible > CCur(.Item(i).SubItems(9)) Then
     curAbono = curAbono + CCur(.Item(i).SubItems(9))
     curDisponible = curDisponible - CCur(.Item(i).SubItems(9))
   Else
     curAbono = curAbono + curDisponible
     curDisponible = 0
   End If


   'MoraIntCor
   If curDisponible > CCur(.Item(i).SubItems(6)) Then
     curAbono = curAbono + CCur(.Item(i).SubItems(6))
     curDisponible = curDisponible - CCur(.Item(i).SubItems(6))
   Else
     curAbono = curAbono + curDisponible
     curDisponible = 0
   End If
  
   'MoraIntMor
   If curDisponible > CCur(.Item(i).SubItems(7)) Then
     curAbono = curAbono + CCur(.Item(i).SubItems(7))
     curDisponible = curDisponible - CCur(.Item(i).SubItems(7))
   Else
     curAbono = curAbono + curDisponible
     curDisponible = 0
   End If
  
  
   'Saldo o principal atrasado
    If curDisponible > CCur(.Item(i).SubItems(5)) Then
      curAbono = curAbono + CCur(.Item(i).SubItems(5))
      curDisponible = curDisponible - CCur(.Item(i).SubItems(5))
    Else
      curAbono = curAbono + curDisponible
      curDisponible = 0
    End If
  
   .Item(i).SubItems(11) = Format(curAbono, "Standard")
   lblDisponible.Caption = Format(curDisponible, "Standard")
  
 Next i
End With

Me.MousePointer = vbDefault
MsgBox "Distribución Automática Aplicada...", vbInformation

End Sub

Private Sub cmdMAceptar_Click()
Dim i As Integer

If CCur(txtMAbono) > CCur(lblMTotalDeuda.Caption) Then
  MsgBox "El monto del Abono es Mayor que el Total Adeudado...", vbExclamation
  Exit Sub
End If

If CCur(txtMAbono) > CCur(lblDisponible.Caption) Then
  MsgBox "El monto del Abono es Mayor que el Disponible de Aplicación...", vbExclamation
  Exit Sub
End If

'Pasar Dato del Abono
lswAbonos.SelectedItem.SubItems(11) = txtMAbono
lblDisponible.Caption = Format(CCur(lblDisponible.Caption) - CCur(txtMAbono), "Standard")

fraAbono.Visible = False
lswAbonos.Visible = True

End Sub

Private Sub cmdMCancelar_Click()
lblDisponible.Caption = Format(CCur(lblDisponible.Caption) - CCur(lswAbonos.SelectedItem.SubItems(11)), "Standard")

fraAbono.Visible = False
lswAbonos.Visible = True

End Sub

Private Sub sbLimpiaDatos()

tcMain.Item(0).Enabled = False
tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False
tcMain.Item(3).Enabled = False
tcMain.Item(4).Enabled = False

fraObservacion.Visible = False

Select Case tcMain.SelectedItem
  Case 0 'Renuncia
    tcMain.Item(0).Enabled = True
    txtCedula = ""
    txtNombre = ""
    cboTipo.Clear
    cboTipoDoc.Text = "Cheque"
    cboCuenta.Clear
    
    lblEstadoActual.Caption = ""
    lblIngreso.Caption = ""
    lblBoleta.Caption = ""
    
    txtAc_Boleta.Text = ""
    dtpAc_fecha.Value = fxFechaServidor
    dtpAc_fecha.Visible = False
    lblAcFecha.Caption = ""
    
     
  
  Case 1 'Aportes
    tcMain.Item(1).Enabled = True
    
    lblAporteObrero.Caption = 0
    lblAportePatronal.Caption = 0
    lblCustodia.Caption = 0
    lblFCI.Caption = 0
    lblCapitalizacion.Caption = 0
    lblAporteExtra.Caption = 0
    lblRenta.Caption = 0
    lblExcedenteRenta.Caption = 0
    lblExcedente.Caption = 0
    lblExcedente.Tag = 0
     
    lblTotalBruto.Caption = 0
    txtRetenerMonto = 0
    lblTotalNeto(0).Caption = 0
     
    chkAplObrero.Value = xtpUnchecked
    chkAplPatronal.Value = xtpUnchecked
    chkAplCapGen.Value = xtpUnchecked
    chkAplCapExtra.Value = xtpUnchecked
    chkAplExcedente.Value = xtpUnchecked
     
    chkAplObrero.Enabled = False
    chkAplPatronal.Enabled = False
    chkAplCapGen.Enabled = False
    chkAplCapExtra.Enabled = False
    chkAplExcedente.Enabled = False
     
  Case 2 'Planes de Ahorro
    tcMain.Item(2).Enabled = True
    
    lswPlanes.ListItems.Clear
    lblTotalNeto.Item(1).Caption = lblTotalNeto.Item(0).Caption
         
     
  Case 3 'Abonos
    tcMain.Item(4).Enabled = True
    
    lswAbonos.ListItems.Clear
'    lblLsw.Caption = ""
     
    lblDisponible.Caption = 0
     
    lblMOperacion.Caption = ""
    lblMCodigo.Caption = ""
    lblMTipo.Caption = ""
    lblMGarantia.Caption = ""
    lblMLineaDesc.Caption = ""
    lblMTotalDeuda.Caption = ""
     
    lblMSaldo.Caption = ""
    lblMMorPrincipal.Caption = ""
    lblMMorIntCor.Caption = ""
    lblMMoraIntMor.Caption = ""
    lblMCargos.Caption = ""
    lblMPolizas.Caption = ""
    txtMAbono = 0
    
  
  Case 4 'Sumario
    tcMain.Item(4).Enabled = True
    
    txtSumario = ""
    cmdAplicar.Enabled = True
    
    fraPrg.Visible = False
    fraResumen.Visible = False
    fraObservacion.Visible = True


End Select

End Sub

'Mantiene por Compatibilidad
Private Function fxFCI(vCedula As String) As Currency
 fxFCI = 0
End Function

Private Sub sbCargaDatos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, rsTmp As New ADODB.Recordset
Dim curTotalLiq As Currency, curTotalPrestamos As Currency
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case tcMain.SelectedItem
  Case 0 'Renuncia
    strSQL = "select S.cedula,S.nombre,S.fechaingreso,S.estadoactual,0 as Boleta,isnull(E.descripcion,'') as 'EstadoPersona'" _
           & " from socios S inner join AFI_ESTADOS_PERSONA E on S.estadoActual = E.cod_estado" _
           & " where S.cedula = '" & txtCedula & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
       txtCedula.Text = rs!Cedula
       txtNombre.Text = rs!Nombre & ""
       cboTipo.Clear
       
       lblEstadoActual.Caption = rs!EstadoPersona
       lblEstadoActual.Tag = rs!EstadoActual
       Select Case UCase(rs!EstadoActual)
          Case "S"
            cboTipo.AddItem "01 - Ren.Asociación"
            cboTipo.ItemData(cboTipo.ListCount - 1) = "A"
            
            cboTipo.AddItem "02 - Ren.Patronal"
            cboTipo.ItemData(cboTipo.ListCount - 1) = "P"
            cboTipo.Text = "01 - Ren.Asociación"
            cboTipo.ItemData(cboTipo.ListCount - 1) = "A"
          
          Case "A"
            cboTipo.AddItem "02 - Ren.Patronal"
            cboTipo.ItemData(cboTipo.ListCount - 1) = "P"
            cboTipo.Text = "02 - Ren.Patronal"
          Case "P"
            cboTipo.AddItem "03 - No Aplica"
            cboTipo.ItemData(cboTipo.ListCount - 1) = "N"
            cboTipo.Text = "03 - No Aplica"
          Case "N"
            cboTipo.AddItem "03 - No Aplica"
            cboTipo.ItemData(cboTipo.ListCount - 1) = "N"
            cboTipo.Text = "03 - No Aplica"
       End Select
       lblIngreso.Caption = Format(IIf(IsNull(rs!FechaIngreso), Date, rs!FechaIngreso), "yyyy/mm/dd")
       lblBoleta.Caption = IIf(IsNull(rs!Boleta), 0, rs!Boleta)
    
    Else
       MsgBox "No Se encontró ningun registro de la Persona, verifique...", vbInformation
    End If
    rs.Close
    
    
  Case 1 'Aportes
    
    strSQL = "exec spAFI_Liq_Consulta_Patrimonio '" & txtCedula & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF And Not rs.BOF Then
        lblAporteObrero.Caption = Format(rs!ahorro, "Standard")
        lblAportePatronal.Caption = Format(rs!Aporte, "Standard")
        
        lblCustodia.Caption = Format(rs!Custodia, "Standard")
        
        lblCapitalizacion.Caption = Format(rs!capitaliza, "Standard")
        lblAporteExtra.Caption = Format(rs!Extra, "Standard")
        
        
        lblRenta.Caption = Format(rs!Renta, "Standard")
        
        lblExcedente.Caption = Format(rs!Excedente, "Standard")
        lblExcedenteRenta.Caption = Format(rs!EXC_RENTA, "Standard")
        lblExcedente.Tag = rs!EXC_APLICA
        
        txtRetenerMonto = Format(rs!Renta, "Standard")
        
        txtDivisa.Text = rs!COD_DIVISA
        
        txtDivisaLocal.Text = rs!divisa_local
        txtTipoCambio.Text = Format(rs!TIPO_CAMBIO, "########0.0000")
        
    End If
    rs.Close
    lblFCI.Caption = "0"

    chkAplPatronal.Value = xtpUnchecked
    chkAplExcedente.Value = xtpChecked
    
    Select Case Mid(cboTipo, 1, 2)
      Case "01" 'Renuncia a la Asociación (Liquidacion Parcial)
        chkAplObrero.Enabled = True
        chkAplCapGen.Enabled = True
        chkAplCapExtra.Enabled = True
        
        chkAplPatronal.Enabled = False
        chkAplExcedente.Enabled = False
        
      Case "02" 'Renuncia Patronal (Liquidacion Total)
        chkAplObrero.Enabled = True
        chkAplPatronal.Enabled = True
        chkAplCapGen.Enabled = True
        chkAplCapExtra.Enabled = True
        
        chkAplExcedente.Enabled = False
        If chkAplExcedente.Tag = "1" Then
            chkAplExcedente.Enabled = True
        End If
        
    End Select
     
    'Aplica Marcas por Default
        chkAplObrero.Value = vbChecked
        chkAplCapGen.Value = vbChecked
        chkAplCapExtra.Value = vbChecked
        
        
     
    Call sbAportesTotales
     '********** SE DESBLOQUEA ESTA OPCION PORQUE HAY UN COMITE SI DECIDE LA CAUSA
     'DE MUERTE, POR TANTO NO SE SABE Y SE DEJA A CRITERIO DEL USUARIO
     '**********
'    'Si la Causa es por Muerte no se le aplica nada a las deudas
'    If chkMortalidad.Value = vbChecked Then
'        chkAplObrero.Enabled = False
'        chkAplPatronal.Enabled = False
'        chkAplCapGen.Enabled = False
'        chkAplCapExtra.Enabled = False
'    End If
     
  Case 2 'Planes de Ahorros
  
     vPaso = True
     
    mFechaSistema = fxFechaServidor
    
    Select Case Mid(cboTipo, 1, 2)
      Case "01" 'Renuncia a la Asociación (Liquidacion Parcial)
         strSQL = "exec spAfiLiquidaListaPlanes '" & txtCedula.Text & "','A'"
      
      Case "02" 'Renuncia Patronal (Liquidacion Total)
         strSQL = "exec spAfiLiquidaListaPlanes '" & txtCedula.Text & "','P'"
    End Select
     
     
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       Set itmX = lswPlanes.ListItems.Add(, , rs!COD_Contrato)
           itmX.Tag = rs!COD_OPERADORA
           itmX.SubItems(1) = rs!Cod_Plan
           itmX.SubItems(2) = Format(rs!APORTES + rs!Rendimiento + rs!RendPendiente - rs!Multa, "Standard")
           itmX.SubItems(3) = Format(rs!APORTES, "Standard")
           itmX.SubItems(4) = Format(rs!Rendimiento, "Standard")
           itmX.SubItems(5) = Format(rs!RendPendiente, "Standard")
           itmX.SubItems(6) = Format(rs!Multa, "Standard")
           itmX.SubItems(7) = rs!operadoraX
           itmX.SubItems(8) = rs!PlanX
           itmX.SubItems(9) = 0
           itmX.SubItems(10) = 0
           itmX.SubItems(11) = IIf((rs!RENTA_GLOBAL = 1), "Sí", "No")
           itmX.SubItems(12) = rs!COD_DIVISA
           itmX.SubItems(13) = rs!TIPO_CAMBIO
           
           
       rs.MoveNext
     Loop
     rs.Close
     vPaso = False
     
     txtFndRendGravado.Text = "0"
     txtFndRendLiquidar.Text = "0"
     lblTotalNeto.Item(2).Caption = "0" 'ISR Monto
     lblTotalNeto.Item(1).Caption = Format(CCur(lblTotalNeto.Item(0).Caption), "Standard")


    'Marcada por Default
    For i = 1 To lswPlanes.ListItems.Count
        lswPlanes.ListItems.Item(i).Checked = True
    Next i

  
  Case 3 'Abonos
    'op,cod,tipo,garantia,saldo,moraintc,moraintm,moraprin,abono
    lblDisponible.Caption = lblTotalNeto(1).Caption

    strSQL = "exec spAfi_Liquidacion_CreditosPersona '" & txtCedula.Text & "'"
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      
         Set itmX = lswAbonos.ListItems.Add(, , rs!Id_Solicitud)
             itmX.SubItems(1) = rs!Codigo
             itmX.SubItems(2) = rs!Descripcion
             itmX.SubItems(3) = rs!Detalle
             itmX.SubItems(4) = rs!GarantiaX
             itmX.SubItems(5) = Format(rs!Saldo, "Standard")
             itmX.SubItems(6) = Format(rs!IntC, "Standard")
             itmX.SubItems(7) = Format(rs!IntM, "Standard")
             itmX.SubItems(8) = Format(rs!Amortiza, "Standard")
             itmX.SubItems(9) = Format(rs!Cargos, "Standard")
             itmX.SubItems(10) = Format(rs!Polizas, "Standard")
             itmX.SubItems(11) = 0
             itmX.SubItems(12) = rs!COD_DIVISA
             itmX.SubItems(13) = rs!TIPO_CAMBIO
      
      rs.MoveNext
    Loop
    rs.Close
    
    'Aplica por Default Abono Auto
    Call cmdDistribucionAuto_Click
    
  Case 4 'Sumario
    
    fraObservacion.Visible = True
    fraResumen.Visible = False
    fraPrg.Visible = False
    
    If chkMortalidad.Value = vbChecked Then
      txtObservacion = "RENUNCIA POR MORTALIDAD, SE DEBE GIRAR MONTO ESPECIFICADO A BENEFICIARIOS" _
                     & ", SE ADJUNTA BOLETA CON EL ASIENTO DE LA LIQUIDACION"
    Else
      txtObservacion = ""
    End If
    
    curTotalLiq = 0
    curTotalPrestamos = 0
    
    txtSumario = "-----------------------------------------------------" & vbCrLf
    txtSumario = txtSumario & "PROCESAR LA LIQUIDACION PARA: " & txtNombre & vbCrLf
    txtSumario = txtSumario & "-----------------------------------------------------" & vbCrLf
    txtSumario = txtSumario & "RENUNCIA >>>>>" & vbCrLf
    txtSumario = txtSumario & "TIPO: " & UCase(cboTipo.Text) & vbCrLf
    txtSumario = txtSumario & "CAUSA: " & cboCausa.Text & vbCrLf
    txtSumario = txtSumario & "MORTALIDAD: " & IIf(chkMortalidad.Value = vbChecked, "SI", "NO") & vbCrLf & vbCrLf
    txtSumario = txtSumario & "DEPOSITOS >>>>>" & vbCrLf
    txtSumario = txtSumario & "TIPO: " & UCase(cboTipoDoc.Text) & vbCrLf
    txtSumario = txtSumario & "BANCO: " & cboBanco.Text & vbCrLf
    txtSumario = txtSumario & "CUENTA: " & cboCuenta.Text & vbCrLf & vbCrLf
    
    txtSumario = txtSumario & "APLICACION DE APORTES >>>>>" & vbCrLf
    
    If chkAplObrero.Value = vbChecked Then
      txtSumario = txtSumario & "[x] APORTE OBRERO : " & lblAporteObrero.Caption & vbCrLf
    Else
      txtSumario = txtSumario & "[ ] APORTE OBRERO : " & lblAporteObrero.Caption & vbCrLf
    End If

    If chkAplPatronal.Value = vbChecked Then
      txtSumario = txtSumario & "[x] APORTE PATRONAL : " & lblAportePatronal.Caption & vbCrLf
      txtSumario = txtSumario & "[x] APORTE CUSTODIA : " & lblCustodia.Caption & vbCrLf
    Else
      txtSumario = txtSumario & "[ ] APORTE PATRONAL : " & lblAportePatronal.Caption & vbCrLf
      txtSumario = txtSumario & "[ ] APORTE CUSTODIA : " & lblCustodia.Caption & vbCrLf
    End If

    If chkAplCapGen.Value = vbChecked Then
      txtSumario = txtSumario & "[x] CAPITALIZACION : " & lblCapitalizacion.Caption & vbCrLf
    Else
      txtSumario = txtSumario & "[ ] CAPITALIZACION : " & lblCapitalizacion.Caption & vbCrLf
    End If


    If chkAplExcedente.Value = vbChecked Then
      txtSumario = txtSumario & "[x] EXCEDENTE PERIODO : " & lblExcedente.Caption & vbCrLf
    Else
      txtSumario = txtSumario & "[ ] EXCEDENTE PERIODO : " & lblExcedente.Caption & vbCrLf
    End If


    If chkAplCapExtra.Value = vbChecked Then
      txtSumario = txtSumario & "[x] AHORRO EXTRAORDINARIO : " & lblAporteExtra.Caption & vbCrLf
    Else
      txtSumario = txtSumario & "[ ] AHORRO EXTRAORDINARIO : " & lblAporteExtra.Caption & vbCrLf & vbCrLf
    End If
    
    Select Case Mid(cboTipo.Text, 1, 2)
      Case "01" 'Liq.Interna
        curTotalLiq = CCur(lblAporteExtra.Caption) + CCur(lblAporteObrero.Caption) _
                    + CCur(lblCapitalizacion.Caption)
      Case "02" 'Liq.Total
        curTotalLiq = CCur(lblAporteExtra.Caption) + CCur(lblAporteObrero.Caption) _
                    + CCur(lblCapitalizacion.Caption) _
                    + CCur(lblAportePatronal.Caption) + CCur(lblCustodia.Caption)
        
        If lblExcedente.Tag = "1" Then
          curTotalLiq = curTotalLiq + CCur(lblExcedente.Caption)
        End If
        
    End Select
    
    txtSumario = txtSumario & vbCrLf & "PLANES DE AHORRO A LIQUIDAR: " & vbCrLf

    With lswPlanes.ListItems
        For i = 1 To .Count
          If .Item(i).Checked Then
              curTotalLiq = curTotalLiq + .Item(i).SubItems(2)
              txtSumario = txtSumario & "CONTRATO: " & .Item(i) & " >> PLAN: " & .Item(i).SubItems(1) & " >> MONTO: " & .Item(i).SubItems(2) & vbCrLf
          End If
        Next i
    End With
    
    txtSumario = txtSumario & vbCrLf & "APLICACIONES A CREDITOS: " & vbCrLf
    
    Dim pSaldoRes As Currency, pAbono As Currency
    
    With lswAbonos.ListItems
        For i = 1 To .Count
          If .Item(i).SubItems(11) > 0 Then
              pAbono = CCur(.Item(i).SubItems(11))
              pSaldoRes = CCur(.Item(i).SubItems(5))
              
              pAbono = pAbono - (CCur(.Item(i).SubItems(6)) + CCur(.Item(i).SubItems(7)) + CCur(.Item(i).SubItems(9)) _
                    + CCur(.Item(i).SubItems(10)))
              
              If pAbono > 0 Then
                pSaldoRes = pSaldoRes - pAbono
              End If
              
              txtSumario = txtSumario & "OPERACION: " & .Item(i).Text & " >> CODIGO: " & .Item(i).SubItems(1) & " >> GARANTIA: " & .Item(i).SubItems(4) & vbCrLf _
                    & " >> TOTAL ABONO    : " & .Item(i).SubItems(11) & vbCrLf _
                    & " >> ABONO PRINCIPAL: " & Format(pAbono, "Standard") _
                    & " >> NUEVO SALDO: " & Format(pSaldoRes, "Standard") & vbCrLf & vbCrLf
          End If
        Next i
    End With
    

    
    curTotalLiq = curTotalLiq - CCur(txtRetenerMonto)
    curTotalPrestamos = CCur(lblTotalNeto(1).Caption) - CCur(lblDisponible.Caption)

    
    txtSumario = txtSumario & vbCrLf & vbCrLf & "TOTALES >>>>>" & vbCrLf
    txtSumario = txtSumario & "TOTAL A LIQUIDAR : " & Format(curTotalLiq, "Standard") & vbCrLf
    txtSumario = txtSumario & "TOTAL APLICADO A PRESTAMOS : " & Format(curTotalPrestamos, "Standard") & vbCrLf
    txtSumario = txtSumario & "TOTAL RETENIDO : " & txtRetenerMonto & vbCrLf
    txtSumario = txtSumario & "TOTAL A GIRAR : " & Format(curTotalLiq - curTotalPrestamos, "Standard") & vbCrLf
    
    
End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Function fxModoSIF() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim Resultado As Boolean

On Error GoTo vError:

Resultado = True

'strSQL = "select Top 1 Fecha from renuncias"
'Call OpenRecordSet(rs, strSQL)
'    Resultado = False
'rs.Close

fxModoSIF = Resultado

Exit Function

vError:
 fxModoSIF = Resultado
 
End Function


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 1
Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture


Call Formularios(Me)


vConcepto = "AFL002"
vTipoDoc = "LIQ"

tcMain.Item(0).Selected = True


With lswPlanes.ColumnHeaders
   .Clear
   .Add , , "No. Contrato", 1440
   .Add , , "Plan", 1140, vbCenter
   .Add , , "Disponible", 1440, vbRightJustify
   .Add , , "Aportes", 1440, vbRightJustify
   .Add , , "Rendimientos", 1440, vbRightJustify
   .Add , , "Rend.Pend.", 1440, vbRightJustify
   .Add , , "(-) Multas", 1440, vbRightJustify
   .Add , , "Operadora", 1100, vbCenter
   .Add , , "Plan Desc.", 3440
   .Add , , "ISR Monto", 2440, vbRightJustify
   .Add , , "ISR Porc", 1440, vbRightJustify
   .Add , , "ISR Apl?", 1440, vbCenter
   .Add , , "Divisa", 1440, vbCenter
   .Add , , "T.C.", 1440, vbRightJustify
End With

With lswAbonos.ColumnHeaders
   .Add , , "Operación", 1200
   .Add , , "Código", 1000, vbCenter
   .Add , , "Descripción", 2000
   .Add , , "Tipo", 1000, vbCenter
   .Add , , "Garantia", 1040, vbCenter
   .Add , , "Saldo", 1440, vbRightJustify
   .Add , , "Int.Cor.", 1140, vbRightJustify
   .Add , , "Int.Mor.", 1140, vbRightJustify
   .Add , , "Principal", 1240, vbRightJustify
   .Add , , "Cargos", 1140, vbRightJustify
   .Add , , "Pólizas", 1140, vbRightJustify
   .Add , , "Abono", 1440, vbRightJustify
   .Add , , "Divisa", 1440, vbCenter
   .Add , , "T.C.", 1440, vbRightJustify
End With

cboTipoDoc.AddItem "Cheque"
cboTipoDoc.ItemData(cboTipoDoc.ListCount - 1) = "CK"
cboTipoDoc.AddItem "Transferencia"
cboTipoDoc.ItemData(cboTipoDoc.ListCount - 1) = "TE"

With glogon
    .strSQL = "select dbo.fxAFI_Liquidacion_FP_Fondos() as 'Fondo'"
    Call OpenRecordSet(.Recordset, .strSQL)
    If .Recordset!Fondo = 1 Then
            cboTipoDoc.AddItem "Fondos"
            cboTipoDoc.ItemData(cboTipoDoc.ListCount - 1) = "FD"
    End If
    .Recordset.Close
End With

cboTipoDoc.Text = "Cheque"


vPaso = True
    'Carga Cuentas Bancarias Autorizadas
    strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
    Call sbCbo_Llena_New(cboBanco, strSQL, False, True)
    
'    'Carga Causas
'    strSQL = "select id_Causa as 'IdX', Descripcion as 'ItmX'" _
'           & " from causas_renuncias WHERE ACTIVO = 1"
'    Call sbCbo_Llena_New(cboCausa, strSQL, False, True)


vPaso = False
Call cboCausa_Click

'Carga Tipos de Acciones (Documentos)
strSQL = "select Id_Documento as 'IdX', Descripcion as 'ItmX' from AFI_CR_RENUNCIAS_TIPO_DOCUMENTO"
Call sbCbo_Llena_New(cboAc_Tipo, strSQL, False, True)




strSQL = "select Activar_Control from afi_cr_parametros"
Call OpenRecordSet(rs, strSQL)
    vControlRenuncias = IIf((rs!activar_control = 1), True, False)
rs.Close

Call sbLimpiaDatos
Call RefrescaTags(Me)


'Modo de Sistema (ASE o SIF)
vModoSIF = fxModoSIF

strSQL = "update afi_cr_renuncias set estado = 'V'" _
       & " where vencimiento < dbo.MyGetdate() and estado = 'T'"
If vControlRenuncias Then Call ConectionExecute(strSQL)

dtpAc_fecha.Value = Format(fxFechaServidor, "dd/mm/yyyy")
dtpPago.Value = dtpAc_fecha.Value

End Sub


Private Sub imgFechaAccion_Click()
If dtpAc_fecha.Visible Then
  dtpAc_fecha.Visible = False
Else
  dtpAc_fecha.Visible = True
End If
End Sub

Private Sub imgObservacion_Click()
    
    fraPrg.Visible = False
    fraResumen.Visible = False
    fraObservacion.Visible = True

End Sub



Private Sub lswAbonos_DblClick()

fraAbono.Left = lswAbonos.Left
fraAbono.top = lswAbonos.top
fraAbono.Visible = True
lswAbonos.Visible = False

With lswAbonos.SelectedItem
    lblMOperacion.Caption = .Text
    lblMCodigo.Caption = .SubItems(1)
    lblMLineaDesc.Caption = .SubItems(2)
    lblMTipo.Caption = .SubItems(3)
    lblMGarantia.Caption = .SubItems(4)
    
    lblMSaldo.Caption = .SubItems(5)
    lblMMorIntCor.Caption = .SubItems(6)
    lblMMoraIntMor.Caption = .SubItems(7)
    lblMMorPrincipal.Caption = .SubItems(8)
    lblMCargos.Caption = .SubItems(9)
    lblMPolizas.Caption = .SubItems(10)
    
    lblMTotalDeuda.Caption = Format(CCur(.SubItems(5)) + CCur(.SubItems(6)) + CCur(.SubItems(7)) + CCur(.SubItems(9)) + CCur(.SubItems(10)), "Standard")
    
    lblMMoraTotal.Caption = Format(CCur(.SubItems(8)) + CCur(.SubItems(6)) + CCur(.SubItems(7)) + CCur(.SubItems(9)) + CCur(.SubItems(10)), "Standard")
    
    txtMAbono = .SubItems(11)
    
    lblDisponible.Caption = Format(CCur(lblDisponible.Caption) + CCur(.SubItems(11)), "Standard")
End With


End Sub


Private Sub tlbNotas_ButtonClick(ByVal Button As MSComctlLib.Button)
fraObservacion.Visible = False
End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer


Select Case Button.Key

  Case "Anterior"
        If tcMain.SelectedItem > 0 Then
         tcMain.Item(tcMain.Selected - 1).Selected = True
         
         For i = 0 To tcMain.ItemCount - 1
           tcMain.Item(i).Enabled = False
         Next i
         'Preguntar si desea limpiar los datos
         i = MsgBox("Desea Limpiar Los Datos Anteriores...", vbYesNo)
         If i = vbYes Then
           Call sbLimpiaDatos
           Call sbCargaDatos
         End If
        End If
        
        tcMain.Item(tcMain.SelectedItem).Enabled = True
        
  Case "Siguiente"
        If tcMain.SelectedItem = 0 Then
            If fxVerificaDatos Then
               tcMain.Item(tcMain.Selected + 1).Selected = True
               Call sbLimpiaDatos
               Call sbCargaDatos
            End If
        Else
            If tcMain.SelectedItem < 4 Then
              tcMain.Item(tcMain.Selected + 1).Selected = True
              Call sbLimpiaDatos
              Call sbCargaDatos
            End If
        End If

End Select


End Sub

Private Sub lswAbonos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curAbono As Currency, curDeuda As Currency
Dim curMora As Currency

On Error GoTo vError

curAbono = CCur(lswAbonos.SelectedItem.SubItems(11))
curDeuda = CCur(lswAbonos.SelectedItem.SubItems(5)) + CCur(lswAbonos.SelectedItem.SubItems(6)) + CCur(lswAbonos.SelectedItem.SubItems(7) + CCur(lswAbonos.SelectedItem.SubItems(9)) + CCur(lswAbonos.SelectedItem.SubItems(10)))
curMora = CCur(lswAbonos.SelectedItem.SubItems(8)) + CCur(lswAbonos.SelectedItem.SubItems(6)) + CCur(lswAbonos.SelectedItem.SubItems(7) + CCur(lswAbonos.SelectedItem.SubItems(9)) + CCur(lswAbonos.SelectedItem.SubItems(10)))

lblLsw.Caption = "La Operación : " & lswAbonos.SelectedItem & vbCrLf
If curDeuda > curAbono Then
    If curMora > curAbono Then
      lblLsw.Caption = lblLsw.Caption & " -- Queda con Morosidad"
    Else
      lblLsw.Caption = lblLsw.Caption & " -- Queda al día"
    End If
Else
  lblLsw.Caption = lblLsw.Caption & " -- Queda Cancelada"
End If

Exit Sub

vError:
  lblLsw.Caption = "..."

End Sub



Private Sub lswPlanes_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass


If vPaso Then Exit Sub

lblTotalNeto.Item(1).Caption = CCur(lblTotalNeto.Item(0).Caption)
txtFndRendLiquidar.Text = "0"

With lswPlanes.ListItems
    For i = 1 To .Count
     
         If .Item(i).Checked Then
           lblTotalNeto.Item(1).Caption = Format(CCur(lblTotalNeto.Item(1).Caption) + CCur(.Item(i).SubItems(2)), "Standard")
           If Mid(.Item(i).SubItems(11), 1, 1) = "S" Then
               txtFndRendLiquidar.Text = CCur(txtFndRendLiquidar.Text) + CCur(.Item(i).SubItems(4)) + CCur(.Item(i).SubItems(5))
           End If
        End If
    
    
    Next i

End With

'Consulta Renta Global
   strSQL = "exec spFnd_Renta_Global '" & txtCedula & "', '" & Format(mFechaSistema, "yyyy/mm/dd hh:mm") _
       & "'," & CCur(txtFndRendLiquidar.Text)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   Item.SubItems(10) = Format(rs!RG_Porcentaje, "Standard")
  
   txtFndRendGravado.Text = Format(rs!Retiro_Gravable, "Standard")
   lblTotalNeto.Item(2).Caption = Format(rs!ISR_MONTO, "Standard")
End If
rs.Close
       

lblTotalNeto.Item(1).Caption = Format(CCur(lblTotalNeto.Item(1).Caption) - CCur(lblTotalNeto.Item(2).Caption), "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 
End Sub


Private Sub PushButton1_Click()
Dim strSQL As String

strSQL = "update R set OPEX = 0" _
       & " from REG_CREDITOS R inner join LIQUIDA_DETALLE L on R.ID_SOLICITUD = L.ID_SOLICITUD" _
       & " Where L.CONSEC = " & txtAsientoNo.Text
Call ConectionExecute(strSQL)


strSQL = "delete SIF_TRANSACCIONES_ASIENTO where Tipo_Documento = 'LIQ' and cod_transaccion = '" & txtAsientoNo.Text & "'"
Call ConectionExecute(strSQL)

strSQL = "delete SIF_TRANSACCIONES where Tipo_Documento = 'LIQ' and cod_transaccion = '" & txtAsientoNo.Text & "'"
Call ConectionExecute(strSQL)

Call sbLiqP4Asiento(txtAsientoNo.Text)

strSQL = "update R set OPEX = 1" _
       & " from REG_CREDITOS R inner join LIQUIDA_DETALLE L on R.ID_SOLICITUD = L.ID_SOLICITUD" _
       & " Where L.CONSEC = " & txtAsientoNo.Text & " and R.saldo > 0"
Call ConectionExecute(strSQL)

MsgBox "Comprobante Reconstruido Satisfactoriamente!", vbInformation

End Sub





Private Sub txtAc_Boleta_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 Call sbCargaDatos
 txtNombre.SetFocus
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   
   If vControlRenuncias Then
       gBusquedas.Consulta = "select Cedula, Id_Alterno, Nombre from vAFI_Renuncias_SinLiquidar" _
                           & ""
   Else
       gBusquedas.Consulta = "select Cedula, Nombre from socios"
       gBusquedas.Filtro = " and EstadoActual in('S', 'A')"
   End If
   gBusquedas.Convertir = "N"
   
   frmBusquedas.Show vbModal
   
   txtCedula.Text = RTrim(gBusquedas.Resultado)
   txtNombre.Text = RTrim(gBusquedas.Resultado2)
   Call sbCargaDatos
End If

End Sub


Private Sub txtCuentaAhorros_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call btnSiguiente_Click

End Sub

Private Sub txtMAbono_GotFocus()
On Error GoTo vError
txtMAbono = CCur(txtMAbono)
vError:
End Sub

Private Sub txtMAbono_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtMAbono = Format(CCur(txtMAbono), "Standard")
  cmdMAceptar.SetFocus
End If

vError:
End Sub

Private Sub txtMAbono_LostFocus()
On Error GoTo vError
txtMAbono = Format(CCur(txtMAbono), "Standard")
vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
If KeyCode = vbKeyF4 Then
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "S.Nombre"
   gBusquedas.Orden = "S.Nombre"
   If vControlRenuncias Then
       gBusquedas.Consulta = "select S.Cedula,S.Nombre from socios S inner join afi_cr_renuncias R" _
                           & " on S.cedula = R.cedula and R.liq is null and R.estado in('P','V')"
   Else
       gBusquedas.Consulta = "select S.Cedula,S.Nombre from socios S "
   End If
   
   frmBusquedas.Show vbModal
   
   txtCedula = gBusquedas.Resultado
   txtNombre = gBusquedas.Resultado2
   
   Call sbCargaDatos

End If

End Sub

Private Sub txtRetenerMonto_Change()
Call sbAportesTotales
End Sub

Private Sub txtRetenerMonto_GotFocus()
On Error GoTo vError
txtRetenerMonto = CCur(txtRetenerMonto)
vError:
End Sub

Private Sub txtRetenerMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call btnSiguiente_Click
End Sub

Private Sub txtRetenerMonto_LostFocus()
On Error GoTo vError
txtRetenerMonto = Format(CCur(txtRetenerMonto), "Standard")
vError:
End Sub

Private Sub sbDocumento(pConcepto As String, pNumDoc As Long, pAporteObrero As Currency _
                                , pAportePatronal As Currency, pExtraordinario As Currency, pCapitaliza As Currency _
                                , pPlanes As Currency, pRetenido As Currency, pFecha As Date)


Dim strSQL As String, strLinea(10) As String
Dim vAseDocDetalle As String, vAseDocDeposito As String



vAseDocDetalle = ""
vAseDocDeposito = ""

strLinea(1) = "# LIQ.          " & pNumDoc
strLinea(2) = "Aporte Obrero   " & pAporteObrero
strLinea(3) = "Aporte Patronal " & pAportePatronal
strLinea(4) = "Extra Ordinario " & pExtraordinario
strLinea(5) = "Capitalización  " & pCapitaliza
strLinea(6) = "Planes Ahorros  " & pPlanes
strLinea(7) = ""
strLinea(8) = "Imp.Renta       " & pRetenido
strLinea(9) = ""
strLinea(10) = "Desembolso    " & cboTipoDoc.Text


'strSQL = "delete SIF_TRANSACCIONES_asiento where tipo_documento = 'LIQ' and cod_transaccion = '" & pNumDoc & "'"
'strSQL = strSQL & Space(10) & "delete SIF_TRANSACCIONES where tipo_documento = 'LIQ' and cod_transaccion = '" & pNumDoc & "'"
'Call ConectionExecute(strSQL)
             
             
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
        & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,detalle,documento)" _
        & " values('" & pNumDoc & "','LIQ','" & Format(pFecha, "yyyy/mm/dd hh:mm:ss") & "','" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
        & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "',0,'P','" & pNumDoc _
        & "','" & Mid(cboTipo.Text, 1, 30) & "','" & cboTipoDoc.Text & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & Mid(txtObservacion.Text, 1, 254) & "','" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)

End Sub


Private Function fxInstitucion(strCedula As String) As Integer
Dim strSQL As String, rs As New ADODB.Recordset

'Codigo para extraer la institucion del asociado

strSQL = "select cod_institucion from socios where cedula  = '" & strCedula & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Or Not rs.BOF Then
   fxInstitucion = rs!cod_institucion
Else
  fxInstitucion = 0
End If
rs.Close
End Function

Private Sub sbDatosRenuncia(iRenuncia As Long)
Dim strSQL As String, rs As New ADODB.Recordset

'Codigo de verificación para aplicacion de reingreso

strSQL = "select id_promotor,aplica_reingreso  from afi_cr_renuncias where cod_renuncia = " & iRenuncia & ""
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Or Not rs.BOF Then
   iAplicaReIngreso = rs!Aplica_Reingreso
   iPromotor = rs!ID_PROMOTOR
End If
rs.Close



End Sub

