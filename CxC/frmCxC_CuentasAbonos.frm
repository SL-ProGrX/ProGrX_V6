VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCxC_CuentasAbonos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CxC: Movimientos>Abonos"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCuotas 
      Caption         =   "Cuotas Activas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   2292
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   9732
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   1572
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   9492
         _Version        =   1441793
         _ExtentX        =   16743
         _ExtentY        =   2773
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
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnAjustes 
         Height          =   288
         Left            =   8040
         TabIndex        =   31
         Top             =   240
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "Ajustes"
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
      End
      Begin XtremeSuiteControls.CheckBox chkMarcaTodas 
         Height          =   252
         Left            =   360
         TabIndex        =   32
         Top             =   240
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Marcar Todas"
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
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8760
      Top             =   120
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7635
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Linea"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Recurso"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   372
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8488
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7429
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   312
      Left            =   7680
      TabIndex        =   8
      Top             =   1200
      Width           =   612
      _Version        =   1441793
      _ExtentX        =   1080
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1092
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   9492
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   1926
      _StockProps     =   79
      Caption         =   "Estado"
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
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1212
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   1320
         TabIndex        =   26
         Top             =   240
         Width           =   1932
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amortización"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   3240
         TabIndex        =   25
         Top             =   240
         Width           =   1572
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intereses"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4800
         TabIndex        =   24
         Top             =   240
         Width           =   1572
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   6360
         TabIndex        =   23
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ult.Mov."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   7800
         TabIndex        =   22
         ToolTipText     =   "Si es menor a la fecha de proceso se Utiliza la Fecha de Proceso"
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label lblSaldo 
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
         Height          =   252
         Left            =   1320
         TabIndex        =   21
         Top             =   480
         Width           =   1932
      End
      Begin VB.Label lblAmortiza 
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
         Height          =   252
         Left            =   3240
         TabIndex        =   20
         Top             =   480
         Width           =   1572
      End
      Begin VB.Label lblInteres 
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
         Height          =   252
         Left            =   4800
         TabIndex        =   19
         Top             =   480
         Width           =   1572
      End
      Begin VB.Label lblCuota 
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
         Height          =   252
         Left            =   6360
         TabIndex        =   18
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label lblFecUltMov 
         Alignment       =   2  'Center
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
         Height          =   252
         Left            =   7800
         TabIndex        =   17
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label lblSaldoR 
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
         Height          =   252
         Left            =   1320
         TabIndex        =   16
         Top             =   720
         Width           =   1932
      End
      Begin VB.Label lblAmortizaR 
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
         Height          =   252
         Left            =   3240
         TabIndex        =   15
         Top             =   720
         Width           =   1572
      End
      Begin VB.Label lblInteresR 
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
         Height          =   252
         Left            =   4800
         TabIndex        =   14
         Top             =   720
         Width           =   1572
      End
      Begin VB.Label lblCuotaR 
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
         Height          =   252
         Left            =   6360
         TabIndex        =   13
         Top             =   720
         Width           =   1452
      End
      Begin VB.Label lblFecUltMovR 
         Alignment       =   2  'Center
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
         Height          =   252
         Left            =   7800
         TabIndex        =   12
         Top             =   720
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.GroupBox fraAbono 
      Height          =   975
      Left            =   120
      TabIndex        =   33
      Top             =   3000
      Width           =   9615
      _Version        =   1441793
      _ExtentX        =   16960
      _ExtentY        =   1720
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.DateTimePicker dtpFechaCancelacion 
         Height          =   312
         Left            =   6360
         TabIndex        =   34
         Top             =   600
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.PushButton optAbono 
         Height          =   372
         Index           =   0
         Left            =   1560
         TabIndex        =   35
         Top             =   120
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Ordinario"
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
      End
      Begin XtremeSuiteControls.PushButton optAbono 
         Height          =   372
         Index           =   1
         Left            =   3120
         TabIndex        =   36
         Top             =   120
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Extraordinario"
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
      End
      Begin XtremeSuiteControls.PushButton optAbono 
         Height          =   372
         Index           =   2
         Left            =   4800
         TabIndex        =   37
         Top             =   120
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cancelación"
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
      End
      Begin XtremeSuiteControls.PushButton optAbono 
         Height          =   372
         Index           =   3
         Left            =   6480
         TabIndex        =   38
         Top             =   120
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Adelanto"
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
      End
      Begin VB.Label lblFechaCancelacion 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Abono (Real) por parte del cliente para cancelación...:"
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
         Left            =   1560
         TabIndex        =   40
         Top             =   600
         Width           =   4692
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   0
         TabIndex        =   39
         Top             =   120
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Tipo de Abono:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XtremeSuiteControls.GroupBox fraDatosAbono 
      Height          =   1815
      Left            =   120
      TabIndex        =   41
      Top             =   4080
      Width           =   9615
      _Version        =   1441793
      _ExtentX        =   16960
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Abono:"
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
      Begin XtremeSuiteControls.CheckBox chkFacturas 
         Height          =   252
         Left            =   4440
         TabIndex        =   70
         Top             =   240
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Facturas Pendientes?"
         ForeColor       =   4210752
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
         Alignment       =   1
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   252
         Left            =   2760
         TabIndex        =   42
         Top             =   240
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.CheckBox chkRecalculaCuota 
         Height          =   252
         Left            =   7080
         TabIndex        =   43
         Top             =   1320
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Recalcular Cuota?"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboDiferenciaApl 
         Height          =   312
         Left            =   6840
         TabIndex        =   44
         Top             =   840
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtCuotas 
         Height          =   312
         Left            =   1680
         TabIndex        =   45
         Top             =   240
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
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
         Text            =   "1"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosAmortiza 
         Height          =   312
         Left            =   1680
         TabIndex        =   46
         Top             =   600
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosInteres 
         Height          =   312
         Left            =   1680
         TabIndex        =   47
         Top             =   960
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosCargos 
         Height          =   312
         Left            =   1680
         TabIndex        =   48
         Top             =   1320
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDiferencia 
         Height          =   312
         Left            =   4920
         TabIndex        =   49
         Top             =   1320
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalCancela 
         Height          =   312
         Left            =   4920
         TabIndex        =   50
         Top             =   600
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalPagar 
         Height          =   312
         Left            =   4920
         TabIndex        =   51
         Top             =   960
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   14737632
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
         BackColor       =   14737632
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label23 
         Caption         =   "Aplicar diferencias como..:"
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
         Left            =   6840
         TabIndex        =   59
         Top             =   600
         Width           =   2052
      End
      Begin VB.Label Label27 
         Caption         =   "No. Cuotas"
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
         TabIndex        =   58
         Top             =   240
         Width           =   1572
      End
      Begin VB.Label Label26 
         Caption         =   "Intereses ..:"
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
         TabIndex        =   57
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label25 
         Caption         =   "Amortización ..:"
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
         Top             =   600
         Width           =   1572
      End
      Begin VB.Label Label24 
         Caption         =   "Total a Pagar"
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
         Left            =   3600
         TabIndex        =   55
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label24 
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
         Index           =   1
         Left            =   3600
         TabIndex        =   54
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label Label26 
         Caption         =   "Cargos ..:"
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
         TabIndex        =   53
         Top             =   1320
         Width           =   1092
      End
      Begin VB.Label Label27 
         Caption         =   "Diferencia ...:"
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
         Left            =   3600
         TabIndex        =   52
         Top             =   1320
         Width           =   972
      End
   End
   Begin XtremeSuiteControls.GroupBox fraFormaPago 
      Height          =   1575
      Left            =   120
      TabIndex        =   60
      Top             =   6000
      Width           =   9615
      _Version        =   1441793
      _ExtentX        =   16960
      _ExtentY        =   2778
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipoDoc 
         Height          =   312
         Left            =   1200
         TabIndex        =   61
         Top             =   240
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4895
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
         Height          =   312
         Left            =   4920
         TabIndex        =   62
         Top             =   240
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1200
         TabIndex        =   63
         Top             =   600
         Width           =   5412
         _Version        =   1441793
         _ExtentX        =   9546
         _ExtentY        =   1397
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
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   855
         Index           =   0
         Left            =   6720
         TabIndex        =   64
         Top             =   480
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Pago"
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
         Picture         =   "frmCxC_CuentasAbonos.frx":0000
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   855
         Index           =   1
         Left            =   7680
         TabIndex        =   65
         Top             =   480
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Picture         =   "frmCxC_CuentasAbonos.frx":04AD
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   855
         Index           =   2
         Left            =   8520
         TabIndex        =   66
         Top             =   480
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Cancelar"
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
         Picture         =   "frmCxC_CuentasAbonos.frx":0C85
         TextImageRelation=   1
      End
      Begin VB.Label Label3 
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
         Height          =   252
         Index           =   4
         Left            =   4080
         TabIndex        =   69
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label3 
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
         TabIndex        =   68
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label3 
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
         TabIndex        =   67
         Top             =   240
         Width           =   1452
      End
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
      Height          =   252
      Index           =   1
      Left            =   600
      TabIndex        =   10
      Top             =   840
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Index           =   3
      Left            =   600
      TabIndex        =   9
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12012
   End
End
Attribute VB_Name = "frmCxC_CuentasAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOperacion As Long, vInteres As Currency
Dim vDiasActivo As Long, vFechaHoy As Date, vScroll As Boolean
Dim pCharRelleno As String

Private Sub btnAjustes_Click()
  GLOBALES.gTag = txtOperacion.Text
  frmCxC_CuentasAjustes.Show vbModal
  
  'Verifica si recibio modificaciones, en cuyo caso procede a actualizar datos en pantalla
  If GLOBALES.gTag2 = 1 Then
    Call sbConsultaOperacion
  End If
  
End Sub

Private Sub btnCajas_Click(Index As Integer)
On Error GoTo vError


'Posicion Final
Call txtTotalPagar_LostFocus

Select Case Index
  Case 2 'Cancelar
     Call sbConsultaOperacion
     
  Case 0 'Desgloce
        If Not IsNumeric(txtTotalCancela.Text) Then txtTotalCancela.Text = 0
        If Not ModuloCajas.mConceptoValida Then
           MsgBox "Esta caja no está autorizada para registrar movimientos a este Concepto de Cuentas por Cobrar", vbExclamation
           Exit Sub
        End If
                
        ModuloCajas.mTotalAplicar = CCur(txtTotalPagar.Text)
        
        If ModuloCajas.mTotalAplicar = 0 Then
            MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
            Exit Sub
        End If
        
        ModuloCajas.mServicio = "Abonos a Cuentas por Cobrar"
        
        Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)
        
        txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")
        
        
        If txtTotalCajas.Text <> txtTotalPagar.Text Then
           txtTotalCajas.BackColor = vbRed
        Else
           txtTotalCajas.BackColor = vbWhite
        End If

  Case 1   'Aplicar
    Call CmdAbono_Click
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cboDiferenciaApl_Click()

If Not cboDiferenciaApl.Enabled Then Exit Sub

If cboDiferenciaApl.Text = "Abono Extraordinario" Then
   chkRecalculaCuota.Enabled = True
Else
   chkRecalculaCuota.Enabled = False
End If

End Sub



Private Sub chkMarcaTodas_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkMarcaTodas.Value
Next

If lsw.ListItems.Count > 0 Then
    Call lsw_ItemCheck(lsw.ListItems.Item(1))
End If

End Sub

Private Sub chkRecalculaCuota_Click()

'Call txtTotalCajas_Change

End Sub

Private Sub sbAbono()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNumDoc As String, vCuenta As String
Dim vTipoDoc As String, vFecha As Date
Dim i As Integer, vExtraOrdinario As Boolean


Me.MousePointer = vbHourglass

On Error GoTo vError


vFecha = fxFechaServidor
vExtraOrdinario = False


vTipoDoc = cboTipoDoc.ItemData(cboTipoDoc.ListIndex)

vNumDoc = fxDocumentoConsecutivo(vTipoDoc)


'Inicia Transaccion
glogon.Conection.BeginTrans


Select Case True
  Case optAbono(0).Checked  'Abono Ordinario
  
        If Not cboDiferenciaApl.Enabled Then
            strSQL = "exec spCxC_AbonoOrdinario " & vOperacion & ",'CRD001','" & glogon.Usuario & "','" & vTipoDoc _
                   & "','" & vNumDoc & "'," & CCur(txtTotalCajas.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
        Else
            strSQL = "exec spCxC_AbonoOrdinario " & vOperacion & ",'CRD001','" & glogon.Usuario & "','" & vTipoDoc _
                   & "','" & vNumDoc & "'," & CCur(txtTotalCancela.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
           
           Select Case cboDiferenciaApl.Text
             Case "Adelanto de Cuota"
                strSQL = "exec spCxC_AbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & vTipoDoc _
                       & "','" & vNumDoc & "'," & Abs(CCur(txtDiferencia.Text)) & ",'" _
                       & Format(vFecha, "yyyy/mm/dd") & "',''"
                Call ConectionExecute(strSQL)

             Case "Abono Extraordinario"
                'Calcula Datos del Abono Extraordinario (Dias, Intereses, Cargos, Principal)
'                strSQL = "exec spCxC_CuentaPlanPagosInfoExtraordinario " & vOperacion & "," & Abs(CCur(txtDiferencia.Text)) & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
'                Call OpenRecordSet(rs, strSQL)
'                    'Aplica Cargo por Anticipo
'                    If rs!Cargos > 0 Then
'                       strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & rs!Cargos & ",'" & GLOBALES.gOficinaUnidad _
'                              & "','" & GLOBALES.gOficinaCentroCosto & "','Pago Anticipado','" & glogon.Usuario & "','CA','','',0"
'                       Call ConectionExecute(strSQL)
'                    End If
'                    'Aplica Abono Extraordinario
'                    strSQL = "exec spCxC_AbonoExtraOrdinario " & vOperacion & ",'CRD002','" & glogon.Usuario & "','" & fxTipoASENumero(vTipoDoc) _
'                           & "','" & vNumDoc & "'," & rs!Dias & "," & rs!Intereses & "," & rs!Principal _
'                           & "," & rs!Cargos & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''," & chkRecalculaCuota.Value
'                    Call ConectionExecute(strSQL)
'                rs.Close
                
                vExtraOrdinario = True
           End Select
        
        End If
  
  Case optAbono(1).Checked 'Abono Extraordinario
        'Elimina Cuotas Activas, Registra Abono y Recalcula Plan de Pagos
        'Se Supone que solo queda una cuota activa para poder realizar un ab. extraordinario
        
'        If CCur(lblDatosAnticipo.Caption) > 0 Then
'           strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & CCur(lblDatosAnticipo.Caption) & ",'" & GLOBALES.gOficinaUnidad _
'                  & "','" & GLOBALES.gOficinaCentroCosto & "','Pago Anticipado','" & glogon.Usuario & "','CA','','',0"
'           Call ConectionExecute(strSQL)
'        End If
        
        strSQL = "exec spCxC_AbonoExtraOrdinario " & vOperacion & ",'CRD002','" & glogon.Usuario & "','" & vTipoDoc _
               & "','" & vNumDoc & "'," & vDiasActivo & "," & CCur(txtDatosInteres.Text) & "," & CCur(txtDatosAmortiza.Text) _
               & ",0,'" & Format(vFecha, "yyyy/mm/dd") & "',''," & chkRecalculaCuota.Value
        Call ConectionExecute(strSQL)

        vExtraOrdinario = True
        
  Case optAbono(2).Checked  'Cancelacion
        'Actualiza el estado de la morosidad
        
'        strSQL = "exec spCxC_CuentaPlanPagosMoraActualizaOp " & vOperacion & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
'        Call ConectionExecute(strSQL)
'
'        If CCur(lblDatosAnticipo.Caption) > 0 Then
'           strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & CCur(lblDatosAnticipo.Caption) & ",'" & GLOBALES.gOficinaUnidad _
'                  & "','" & GLOBALES.gOficinaCentroCosto & "','Cancelacion Anticipada','" & glogon.Usuario & "','CA','','',0"
'           Call ConectionExecute(strSQL)
'        End If
        
        strSQL = "exec spCxC_AbonoCancelacion " & vOperacion & ",'CRD003','" & glogon.Usuario & "','" & vTipoDoc _
               & "','" & vNumDoc & "'," & CCur(txtTotalCajas.Text) & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "',''"
        Call ConectionExecute(strSQL)
  
  
  Case optAbono(3).Checked 'Adelanto de Cuotas
       'Activa Nuevas Cuotas y luego las abona

        If Not cboDiferenciaApl.Enabled Then
            strSQL = "exec spCxC_AbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & vTipoDoc _
                   & "','" & vNumDoc & "'," & CCur(txtTotalCajas.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
        Else
            strSQL = "exec spCxC_AbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & vTipoDoc _
                   & "','" & vNumDoc & "'," & CCur(txtTotalCancela.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
           
           Select Case cboDiferenciaApl.Text
             Case "Adelanto de Cuota"
                strSQL = "exec spCxC_AbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & vTipoDoc _
                       & "','" & vNumDoc & "'," & Abs(CCur(txtDiferencia.Text)) & ",'" _
                       & Format(vFecha, "yyyy/mm/dd") & "',''"
                Call ConectionExecute(strSQL)

             Case "Abono Extraordinario"
'                'Calcula Datos del Abono Extraordinario (Dias, Intereses, Cargos, Principal)
'                strSQL = "exec spCxC_CuentaPlanPagosInfoExtraordinario " & vOperacion & "," & Abs(CCur(txtDiferencia.Text)) _
'                       & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
'                Call OpenRecordSet(rs, strSQL)
'                    'Aplica Cargo por Anticipo
'                    If rs!Cargos > 0 Then
'                       strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & rs!Cargos & ",'" & GLOBALES.gOficinaUnidad _
'                              & "','" & GLOBALES.gOficinaCentroCosto & "','Pago Anticipado','" & glogon.Usuario & "','CA','','',0"
'                       Call ConectionExecute(strSQL)
'                    End If
'                    'Aplica Abono Extraordinario
'                    strSQL = "exec spCxC_AbonoExtraOrdinario " & vOperacion & ",'CRD002','" & glogon.Usuario & "','" & fxTipoASENumero(vTipoDoc) _
'                           & "','" & vNumDoc & "'," & rs!Dias & "," & rs!Intereses & "," & rs!Principal _
'                           & "," & rs!Cargos & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''," & chkRecalculaCuota.Value
'                    Call ConectionExecute(strSQL)
'                rs.Close
                
                vExtraOrdinario = True
           End Select
        
        End If



End Select


'Cierra Transaccion
glogon.Conection.CommitTrans

'Indica si debe reprocesar el Plan de Pagos por registro de Abonos Extraordinario
If vExtraOrdinario Then
        strSQL = "exec spCxC_CuentaPlanPagos " & vOperacion
        Call ConectionExecute(strSQL)
End If

'Genera el Comprobante
Select Case True
  Case optAbono(0).Checked  'Abono Ordinario
      Call Bitacora("Registra", "Abono Ordinario a la Operacion : " & vOperacion)
      Call sbDocumentoAbono("ABONO ORDINARIO", vTipoDoc, vNumDoc, "CRD001", vCuenta)
  Case optAbono(1).Checked  'Abono Extraordinario
      Call Bitacora("Registra", "Abono ExtraOrd. " & IIf((chkRecalculaCuota.Value = 1), "Con Recal.", "Sin Recal") & " a la Op.: " & vOperacion)
      Call sbDocumentoAbono("ABONO EXTRAORDINARIO", vTipoDoc, vNumDoc, "CRD002", vCuenta)
  Case optAbono(2).Checked  'Abono De Cancelacion
      Call Bitacora("Registra", "Cancelación de la Operacion : " & vOperacion)
      Call sbDocumentoAbono("CANCELACION DE DEUDA", vTipoDoc, vNumDoc, "CRD003", vCuenta)
  Case optAbono(3).Checked  'Adelanto de Cuotas
      Call Bitacora("Registra", "Adelanto de Cuotas de la Operacion : " & vOperacion)
      Call sbDocumentoAbono("ADELANTO DE CUOTAS", vTipoDoc, vNumDoc, "CRD004", vCuenta)
End Select




'Imprime el Comprobante
Call sbImprimeRecibo(vNumDoc, vTipoDoc)

Me.MousePointer = vbDefault

strSQL = " - Abono aplicado, con : " & cboTipoDoc.Text & " ...No.: " & vNumDoc & vbCrLf _
       & " - Desea Realizar Otra Transacción a esta Operación ?"

i = MsgBox(strSQL, vbYesNo)
If i = vbYes Then
    Call sbConsultaOperacion
    txtTotalCajas.Text = 0
Else
    Unload Me
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 glogon.Conection.RollbackTrans
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxVerifica() As Boolean
Dim vMensaje As String, i As Integer

vMensaje = ""

Call sbSIFCleanTxtInject(txtNotas)

''Verifica el proceso
'If txtProceso.Tag = "J" Then
'   If Not fxCRDAbonosAutorizados(txtCodigo.Text, txtProceso.Tag) Then
'      vMensaje = vMensaje & "- El usuario actual no cuenta con permisos para realizar abonos a Creditos en Cobro Judicial, verifique..." & vbCrLf
'   End If
'End If

'Verifica que la diferencia del Monto a Cancelar no supere el Saldo
If CCur(txtDiferencia.Text) < 0 Then
 If CCur(lblSaldoR.Caption) + CCur(txtDiferencia.Text) < 0 Then
      vMensaje = vMensaje & "- La diferencia supera el saldo!, verifique..." & vbCrLf
 End If
End If


If vOperacion = 0 Then
  vMensaje = vMensaje & "- Número de Operacion no es válido..." & vbCrLf
End If
 
 
'Verifica Saldo Actual
If Not fxCxC_SaldoVerifica(vOperacion, CCur(lblSaldo.Caption)) Then
   vMensaje = vMensaje & "- Esta Operación ha sido modificada, actualice los datos nuevamente antes de realizar el abono..." & vbCrLf
End If
 
If CCur(txtDatosAmortiza) > CCur(lblSaldo.Caption) Then
   vMensaje = vMensaje & "- La Amortización es mayor al Saldo Actual..." & vbCrLf
End If

 If CCur(txtTotalCajas.Text) <= 0 Then
      vMensaje = vMensaje & "- Los valores Recibidos en Cajas no son válidos...verifique...!" & vbCrLf
 End If

'Abono Ordinario (Verificar Secuencia de Check's)
If optAbono.Item(0).Value Then
 For i = 1 To lsw.ListItems.Count
   If i = 1 And Not lsw.ListItems.Item(i).Checked Then
      vMensaje = vMensaje & "- No se ha especificado un orden válido de aplicación de cuotas...!" & vbCrLf
      Exit For
   End If
   
   If i > 1 Then
        If lsw.ListItems.Item(i).Checked And Not lsw.ListItems.Item(i - 1).Checked Then
               vMensaje = vMensaje & "- No se ha especificado un orden válido de aplicación de cuotas...!" & vbCrLf
               Exit For
        End If
   End If
 Next
End If


If chkFacturas.Value = xtpChecked Then
    vMensaje = vMensaje & " - " & chkFacturas.ToolTipText
    
End If

If Len(vMensaje) = 0 Then
  fxVerifica = True
Else
  fxVerifica = False
  MsgBox vMensaje, vbExclamation
End If

End Function


Private Sub CmdAbono_Click()
Dim iRespuesta As Integer

If Not fxVerifica Then Exit Sub

 iRespuesta = MsgBox("Esta seguro de realizar el abono a esta Operación " & vOperacion, vbYesNo)
 If iRespuesta = vbYes Then
  
  Call sbAbono
  Call sbConsultaOperacion
 
 Else 'Respuesta
  
  MsgBox "Transacción Cancelada...", vbInformation
 
 End If

End Sub



Private Sub dtpFechaCancelacion_Change()

If dtpFechaCancelacion.Enabled Then
   'Refresca información base para Cancelación y/o Abonos Extraordinarios
   Select Case True
      Case optAbono.Item(1).Checked 'Abono Extraordinario
            Call optAbono_Click(1)
      Case optAbono.Item(2).Checked 'Cancelación
            Call optAbono_Click(2)
   End Select
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim vNumCuota As Integer

On Error GoTo vError


vNumCuota = txtCuotas.Text

If vScroll Then
    If FlatScrollBar.Value = 1 Then
       vNumCuota = vNumCuota + 1
    Else
       vNumCuota = vNumCuota - 1
    End If
End If

If vNumCuota <= 0 Then vNumCuota = 1

txtCuotas.Text = vNumCuota

vScroll = False
    FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
 vModulo = 31

End Sub

Private Sub Form_Load()
Dim iDias As Integer

 vModulo = 31
 vOperacion = 0
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

 vFechaHoy = fxFechaServidor
 iDias = fxCrdParametro("32")
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 Call sbLimpiaDatos

dtpFechaCancelacion.Value = vFechaHoy
dtpFechaCancelacion.MinDate = DateAdd("d", (iDias * -1), dtpFechaCancelacion.Value)
dtpFechaCancelacion.MaxDate = dtpFechaCancelacion.Value

 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub sbConsultaOperacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vOperacion = 0 Then Exit Sub

Me.MousePointer = vbHourglass

Call sbLimpiaDatos
 
strSQL = "select R.Operacion,R.saldo,R.proceso,R.Tasa_Corriente,R.interesc,R.amortiza,dbo.fxSIFFechaProcesoConvert(isnull(R.Fecha_UltMov,dbo.MyGetdate())) as 'Fecha_UltMov'" _
       & ",R.cuota,R.cod_concepto,R.cedula,datediff(m,R.Activa_Fecha,dbo.MyGetdate()) as 'Meses'" _
       & ",S.nombre,R.Activa_Fecha,R.Autoriza_Usuario" _
       & ",C.descripcion as 'ConceptoDesc',Ofi.descripcion as 'OficinaDesc',dbo.MyGetdate() as 'FechaServer'" _
       & ",dbo.fxCajas_Valida_Auxiliar('" & ModuloCajas.mCaja & "','CxC',C.cod_Concepto) as 'Caja_Valida_Concepto'" _
       & ",dbo.fxCxC_Operacion_Facturas_Pending(R.Operacion) as 'Facturas'" _
       & " from CxC_Cuentas R inner join CxC_Conceptos C on R.cod_concepto = C.cod_concepto " _
       & " inner join CxC_Personas S on R.cedula = S.cedula" _
       & " left join Sif_Oficinas Ofi on R.cod_Oficina = Ofi.cod_Oficina" _
       & " left join vCxC_CuentasMora V on R.Operacion = V.Operacion" _
       & " where R.estado = 'A' and R.saldo > 0 and R.Operacion = " & vOperacion
      
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  txtOperacion.Text = rs!Operacion
  vOperacion = rs!Operacion
  
  vInteres = rs!Tasa_Corriente
  
  
    ModuloCajas.mClienteId = Trim(rs!Cedula)
    ModuloCajas.mCliente = Trim(rs!Nombre)
    ModuloCajas.mTiquete = Trim(rs!cod_Concepto) & "." & rs!Operacion & "." & Format(Time, "HH:mm:ss")
    
    ModuloCajas.mDivisa = "COL" 'RTrim(rs!Divisa)
    ModuloCajas.mConceptoValida = IIf((rs!Caja_Valida_Concepto > 0), True, False)
 
    ModuloCajas.mTotalDetallado = 0
    txtTotalCajas.Text = 0
     
    chkFacturas.Value = xtpUnchecked
    chkFacturas.ToolTipText = "La Operación no tiene facturas pendientes, está libre para cancelar!"
     
    If rs!facturas > 0 Then
       chkFacturas.Value = xtpChecked
       chkFacturas.ToolTipText = "La Operación tiene pendientes de cancelar " & rs!facturas & " facturas!"
    End If
     
     lblAmortiza.Caption = Format(rs!Amortiza, "Standard")
     lblAmortizaR.Caption = Format(0, "Standard")
     lblCuota = Format(rs!Cuota, "Standard")
     lblCuotaR.Caption = Format(0, "Standard")
     txtDatosAmortiza.Text = Format(0, "Standard")
     txtDatosInteres.Text = Format(0, "Standard")
     lblFecUltMov.Caption = IIf(IsNull(rs!Fecha_UltMov), fxFechaProcesoAnterior(GLOBALES.glngFechaCR), rs!Fecha_UltMov)
    If CLng(lblFecUltMov.Caption) < GLOBALES.glngFechaCR Then
       lblFecUltMov.Caption = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
    End If
     lblFecUltMovR.Caption = 0
     lblInteres.Caption = Format(rs!interesc, "Standard")
     lblInteresR.Caption = Format(0, "Standard")
    
     lblSaldo.Tag = rs!Activa_Fecha & ""
     lblSaldo.Caption = Format(rs!Saldo, "Standard")
     lblSaldoR.Caption = Format(0, "Standard")
    
     txtCuotas.Text = "0"
     txtOperacion.Text = CStr(rs!Operacion)

     fraAbono.Enabled = True
     fraDatosAbono.Enabled = False
    
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
    txtCodigo.Text = rs!cod_Concepto
    txtDescripcion.Text = rs!ConceptoDesc
    
    txtProceso.Tag = rs!Proceso
    Select Case rs!Proceso
      Case "N"
        txtProceso.Text = "Normal"
      Case "T"
        txtProceso.Text = "Traspaso Deuda"
      Case "J"
        txtProceso.Text = "Cobro Judicial"
      Case "I"
        txtProceso.Text = "Incobrable"
    End Select
  
    'Barra de Estado
   
    StatusBarX.Panels.Item(1).Text = rs!OficinaDesc & ""
    StatusBarX.Panels.Item(2).Text = rs!ConceptoDesc & ""
    StatusBarX.Panels.Item(3).Text = rs!Autoriza_Usuario & ""
        
        
       
    'Consulta Cuotas Activas
    strSQL = "select * from CxC_Cuentas_Mov where estado = 'A' and Operacion = " & rs!Operacion _
           & " order by Linea"
    Call OpenRecordSet(rs, strSQL)
    lsw.ListItems.Clear
    
    With lsw.ColumnHeaders
        .Clear
        .Add , , "Linea Id", 1200
        .Add , , "Corte", 1400, vbCenter
        .Add , , "Monto", 1400, vbRightJustify
        .Add , , "Estado", 1200, vbCenter
        .Add , , "Int.Cor.", 1400, vbRightJustify
        .Add , , "Int.Mor.", 1400, vbRightJustify
        .Add , , "Principal", 1400, vbRightJustify
        .Add , , "Cargos", 1400, vbRightJustify
        .Add , , "Dias", 1200, vbCenter
        .Add , , "Dias Mora", 1200, vbCenter
    
    End With
    
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!Linea)
          itmX.SubItems(1) = Format(rs!Fecha_Corte, "dd/mm/yyyy")
          itmX.SubItems(2) = Format(rs!Monto, "Standard")
          itmX.SubItems(3) = IIf((rs!Dias_Mora > 0), "En Mora", "Al Día")
          itmX.SubItems(4) = Format(rs!Int_Cor, "Standard")
          itmX.SubItems(5) = Format(rs!Int_Mor, "Standard")
          itmX.SubItems(6) = Format(rs!Principal, "Standard")
          itmX.SubItems(7) = Format(rs!Cargos, "Standard")
          itmX.SubItems(8) = rs!Dias
          itmX.SubItems(9) = rs!Dias_Mora
          
          itmX.Tag = rs!Linea
      rs.MoveNext
    Loop
    
    
    'Activacion de Tipos de Abonos
    
    Select Case lsw.ListItems.Count
      Case Is <= 0
            optAbono.Item(0).Enabled = False 'Ordinario
            optAbono.Item(1).Enabled = True 'Extraordinario
            optAbono.Item(2).Enabled = True 'Cancelacion
            optAbono.Item(3).Enabled = True 'Adelantos
            Call optAbono_Click(1)
      
      Case Is = 1
            optAbono.Item(0).Enabled = True 'Ordinario
            optAbono.Item(1).Enabled = True 'Extraordinario
            optAbono.Item(2).Enabled = True 'Cancelacion
            optAbono.Item(3).Enabled = True 'Adelantos
            Call optAbono_Click(0)
      
      Case Is > 1
            optAbono.Item(0).Enabled = True 'Ordinario
            optAbono.Item(1).Enabled = False 'Extraordinario
            optAbono.Item(2).Enabled = True 'Cancelacion
            optAbono.Item(3).Enabled = False 'Adelantos
            Call optAbono_Click(0)
    End Select
    

Else
 
 vOperacion = 0
 MsgBox "No se Encontró operación para abonos,puede que se encuentre cancelada ", vbInformation

End If
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub sbLimpiaDatos()
 
 lblAmortiza.Caption = 0
 lblAmortizaR.Caption = 0
 lblCuota = 0
 lblCuotaR.Caption = 0
 txtDatosAmortiza.Text = 0
 txtDatosInteres.Text = 0
 
 lblFecUltMov.Caption = 0
 lblFecUltMovR.Caption = 0
 lblInteres.Caption = 0
 lblInteresR.Caption = 0
 
 txtDescripcion.Text = ""
 
 
 
 lblSaldo.Caption = 0
 lblSaldoR.Caption = 0
 
 txtCedula = ""
 txtCodigo = ""
 txtCuotas = 0
 txtNombre = ""
 txtOperacion = ""


 cboDiferenciaApl.Text = "Adelanto de Cuota"
 cboDiferenciaApl.Enabled = False
 txtTotalCajas.Text = 0
 txtTotalCancela.Text = 0
 
 txtProceso.Tag = ""
 txtProceso.Text = ""
 
 fraAbono.Enabled = False
 fraDatosAbono.Enabled = False
 
 fraCuotas.Visible = False
 lsw.ListItems.Clear
 
 chkRecalculaCuota.Value = vbUnchecked
  
dtpFechaCancelacion.Enabled = False
lblFechaCancelacion.Enabled = False
dtpFechaCancelacion.Value = vFechaHoy
 
End Sub

Private Sub sbBusqueda()

On Error GoTo vError

gBusquedas.Convertir = "N"
gBusquedas.Consulta = "Select R.Operacion as Operacion,R.Codigo,S.Cedula,S.Nombre,C.Descripcion" _
          & " from CxC_Cuentas R inner join CxC_Personas S on R.cedula = S.cedula" _
          & " inner join CxC_Conceptos C on R.codigo = C.codigo"
gBusquedas.Columna = "R.CEDULA"
gBusquedas.Orden = "R.CEDULA"
gBusquedas.Filtro = " AND R.ESTADO = 'A'"

frmBusquedas.Show vbModal

txtOperacion = Trim(gBusquedas.Resultado)
vOperacion = txtOperacion

gBusquedas.Consulta = ""
gBusquedas.Columna = ""
gBusquedas.Orden = ""
gBusquedas.Resultado = ""
gBusquedas.Filtro = ""

Call sbConsultaOperacion

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curInteres As Currency, curPrincipal As Currency, curCargos As Currency
Dim i As Integer

curInteres = 0
curPrincipal = 0
curCargos = 0


With lsw.ListItems
  For i = 1 To .Count
    If .Item(i).Checked Then
       curInteres = curInteres + CCur(.Item(i).SubItems(4)) + CCur(.Item(i).SubItems(5))
       curPrincipal = curPrincipal + CCur(.Item(i).SubItems(6))
       curCargos = curCargos + CCur(.Item(i).SubItems(7))
    End If
  Next i
End With

txtDatosAmortiza.Text = Format(curPrincipal, "Standard")
txtDatosInteres.Text = Format(curInteres, "Standard")
txtDatosCargos.Text = Format(curCargos, "Standard")

 

txtTotalPagar.Text = Format(curPrincipal + curInteres + curCargos, "Standard")
txtTotalCancela.Text = txtTotalPagar.Text
End Sub






Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbCajaInicial

If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   Unload Me
   Exit Sub
End If

Call optAbono_Click(0)

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

Me.Caption = "Abonos a Cuentas por Cobrar    ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

txtTotalCajas.Text = 0
txtNotas.Text = ""
strSQL = "select rTrim(C.tipo_documento) as 'IdX', rtrim(D.Descripcion) as 'itmX'" _
       & " from SIF_DOCUMENTOS D inner join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
       & " Where C.cod_caja =  '" & ModuloCajas.mCaja & "' and D.Tipo_Movimiento in('A','C')" _
       & " order by C.tipo_documento"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)

ModuloCajas.mServicio = "Abonos a Cuentas por Cobrar"

If IsNumeric(ModuloCajas.mRef_01) Then

    txtOperacion.Text = ModuloCajas.mRef_01
    vOperacion = txtOperacion.Text

    Call sbConsultaOperacion
End If

End Sub




Private Sub optAbono_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim curInteres As Currency, curIntMor As Currency, curPrincipal As Currency, curCargos As Currency
Dim i As Integer


Me.MousePointer = vbHourglass

fraCuotas.Visible = False
fraDatosAbono.Enabled = True

chkRecalculaCuota.Enabled = False
chkRecalculaCuota.Value = vbUnchecked

'&H00C0FFC0&
txtTotalPagar.BackColor = &HC0FFC0
txtTotalPagar.Locked = True

txtCuotas.Enabled = False
FlatScrollBar.Enabled = txtCuotas.Enabled

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = False
Next


For i = 0 To 3
  If i = Index Then
      optAbono.Item(i).Checked = True
  Else
      optAbono.Item(i).Checked = False
  End If
Next i

dtpFechaCancelacion.Enabled = False
lblFechaCancelacion.Enabled = False

Select Case Index

 Case 0 'Ordinario
   txtDatosCargos.Text = 0
   txtDatosInteres.Text = 0
   txtDatosAmortiza = 0
      
   txtTotalCancela.Text = 0
   txtTotalPagar.Text = 0
   
   txtCuotas.Text = 0 'Inicializa
   
   fraCuotas.Visible = True
   
   txtTotalPagar.BackColor = vbWhite
   txtTotalPagar.Locked = False
'   txtTotalPagar.SetFocus
   
 Case 1 'Extraordinario
   txtCuotas = 0
   
   lblFechaCancelacion.Caption = "Fecha de Abono (Real) por parte del cliente para Ab.Extraordinario:"
   dtpFechaCancelacion.Enabled = True
   lblFechaCancelacion.Enabled = True
   
   txtDatosInteres.Text = 0
   txtDatosCargos.Text = 0
   
  
   txtDatosAmortiza.Text = 0
   txtTotalCancela.Text = 0
   
   txtTotalPagar.BackColor = vbWhite
   txtTotalPagar.Locked = False
'   txtTotalPagar.SetFocus
   
   chkRecalculaCuota.Enabled = True
'TODO: Revisar este codigo porque en creditos si va activo
'   strSQL = "select dbo.fxCrdPlanPagosDiasActivoFecha( " & txtOperacion.Text & ", '" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "') as 'Dias'"
'   Call OpenRecordSet(rs, strSQL)
'     vDiasActivo = rs!Dias
'   rs.Close
   
Case 2 'Cancelación
   
   txtDatosAmortiza.Text = 0
  
   lblFechaCancelacion.Caption = "Fecha de Abono (Real) por parte del cliente para cancelación...:"
   dtpFechaCancelacion.Enabled = True
   lblFechaCancelacion.Enabled = True
   
'   strSQL = "exec spCxC_CuentaPlanPagosInfoCancelacion " & txtOperacion.Text & ", '" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
'   Call OpenRecordSet(rs, strSQL)
'    txtDatosAmortiza.Text = Format(rs!Principal, "Standard")
'    txtDatosInteres.Text = Format(rs!IntCor + rs!IntMor, "Standard")
'    txtDatosCargos.Text = Format(rs!Cargos, "Standard")
'    txtTotalPagar.Text = Format(rs!Principal + rs!IntCor + rs!IntMor + rs!Cargos, "Standard")
'    txtTotalCancela.Text = txtTotalPagar.Text
'   rs.Close
   

 Case 3 'Adelantos
   txtDatosCargos.Text = 0
   txtDatosInteres.Text = 0
   txtDatosAmortiza.Text = 0
   
   txtCuotas.Enabled = True
   FlatScrollBar.Enabled = txtCuotas.Enabled
   
   txtCuotas.Text = 0 'Inicializa
   txtCuotas.Text = 1 'Inicializa
   txtCuotas.SetFocus
   
   txtTotalPagar.BackColor = vbWhite
   txtTotalPagar.Locked = False
'   txtTotalPagar.SetFocus

End Select

Call RefrescaTags(Me)


Me.MousePointer = vbDefault


End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtNombre = fxCxC_PersonaNombre(txtCedula)
  txtCodigo.SetFocus
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtCodigo = UCase(txtCodigo)
  txtDescripcion.Text = fxCxC_ConceptoDesc(txtCodigo)
   
  txtOperacion.SetFocus
End If
End Sub


Private Sub txtCuotas_Change()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngCuotas As Long

If vOperacion = 0 Then Exit Sub

On Error GoTo vError

If Not IsNumeric(txtCuotas.Text) Then
 lngCuotas = 1
Else
 lngCuotas = txtCuotas.Text
End If

If lngCuotas <= 0 Then lngCuotas = 1

strSQL = "select isnull(max(Linea),0) as 'SeqX', isnull(sum(Int_Cor + Int_Mor),0) as 'IntCor', isnull(sum(Principal),0) as 'Principal'" _
       & ",isnull(min(Saldo_Final),0) as 'Saldo', isnull(max(Fecha_Corte),0) as 'Fecha_Proceso'" _
       & " from CxC_Cuentas_Mov where Operacion = " & vOperacion _
       & " and Linea in(select Top " & lngCuotas & " Linea from CxC_Cuentas_Mov" _
       & " where estado in('A','P') and Operacion = " & vOperacion & " and Linea > 0  order by Linea)"
Call OpenRecordSet(rs, strSQL)
    txtDatosInteres.Text = Format(rs!IntCor, "Standard")
    lblFecUltMovR.Caption = rs!Fecha_Proceso
    
    lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - rs!Principal, "Standard")
    lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + rs!Principal, "Standard")
    lblInteresR.Caption = Format(CCur(lblInteres.Caption) + rs!IntCor, "Standard")

    'Se pone de ultimo porque activa otro sub
    txtDatosAmortiza.Text = Format(rs!Principal, "Standard")

strSQL = "select Monto as 'Cuota' from CxC_Cuentas_Mov where Linea = " & rs!SeqX & " and Operacion = " & vOperacion
rs.Close

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    lblCuotaR.Caption = Format(rs!Cuota, "Standard")
End If
rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtDatosAmortiza.SetFocus
End Sub

Private Sub txtDatosAmortiza_Change()
On Error Resume Next

lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - CCur(txtDatosAmortiza), "Standard")
lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + CCur(txtDatosAmortiza.Text), "Standard")
lblInteresR.Caption = Format(CCur(lblInteres.Caption) + CCur(txtDatosInteres.Text), "Standard")




txtTotalPagar.Text = Format(CCur(txtDatosAmortiza) + CCur(txtDatosInteres.Text) _
                   + CCur(txtDatosCargos.Text), "Standard")
txtTotalCancela.Text = txtTotalPagar.Text
txtDiferencia.Text = "0.00"


End Sub

Private Sub txtDatosAmortiza_GotFocus()
On Error Resume Next
txtDatosAmortiza = CCur(txtDatosAmortiza)
End Sub

Private Sub txtDatosAmortiza_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 txtDatosAmortiza = Format(txtDatosAmortiza, "Standard")
' cboTipoPago.SetFocus
End If
End Sub

Private Sub txtDatosAmortiza_LostFocus()
On Error Resume Next
txtDatosAmortiza = Format(txtDatosAmortiza, "Standard")
End Sub

Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion = xOpTemp
 Call txtOperacion_KeyPress(vbKeyReturn)
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 vOperacion = txtOperacion
 Call sbConsultaOperacion
End If
End Sub


Private Sub sbDocumentoAbono(pTipoAbono As String, pTipoDoc As String, pNumDoc As String _
                                , pConcepto As String, pCuenta As String)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim strCliente As String, vCuenta As String
Dim rsTmp As New ADODB.Recordset, vCuentaPoliza As String, pTipoCambio As Currency
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency

vCuenta = pCuenta

pTipoCambio = fxCajasTipoCambio(ModuloCajas.mDivisa)


'Cuentas
strSQL = "exec spCxC_OperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)


strSQL = "exec spCxC_DocumentoAfectacion '" & fxTipoASENumero(pTipoDoc) & "','" & pNumDoc & "','R'"
Call OpenRecordSet(rsTmp, strSQL, 0)
If rsTmp.EOF And rsTmp.BOF Then
  curIntC = 0
  curIntM = 0
  curAmortiza = 0
  curCargo = 0
Else
  curIntC = rsTmp!IntCor
  curIntM = rsTmp!IntMor
  curAmortiza = rsTmp!Principal
  curCargo = rsTmp!Cargos
End If
rsTmp.Close



'Lineas de Comprobante
strLinea(1) = "Saldo Anterior    ..: " & SIFGlobal.fxStringRelleno(lblSaldo.Caption, "I", pCharRelleno, 15) '
strLinea(2) = "Saldo Actual      ..: " & SIFGlobal.fxStringRelleno(Format(CCur(lblSaldo.Caption) - curAmortiza, "Standard"), "I", pCharRelleno, 15) '
strLinea(3) = "Interes Corriente ..: " & SIFGlobal.fxStringRelleno(Format(curIntC, "Standard"), "I", pCharRelleno, 15) '
strLinea(4) = "Interes Atrasado  ..: " & SIFGlobal.fxStringRelleno(Format(curIntM, "Standard"), "I", pCharRelleno, 15) '
strLinea(5) = "Amortización      ..: " & SIFGlobal.fxStringRelleno(Format(curAmortiza, "Standard"), "I", pCharRelleno, 15) '
strLinea(6) = "Cargos Totales    ..: " & SIFGlobal.fxStringRelleno(Format(curCargo, "Standard"), "I", pCharRelleno, 15) '
strLinea(7) = "Operacion/Concepto..: " & "Op.:" & txtOperacion.Text & " Cpt.:" & txtCodigo.Text

If cboDiferenciaApl.Enabled Then
    strLinea(8) = "Aplica Diferencia ..: " & cboDiferenciaApl.Text
Else
    strLinea(8) = "Descripción       ..: " & txtDescripcion.Text

End If



strLinea(9) = ""
strLinea(10) = "Num. Documento    ..:" & ""
strLinea(11) = ""

strSQL = "exec spCxC_OperacionFechaProxPago " & txtOperacion.Text
Call OpenRecordSet(rsTmp, strSQL, 0)
  If Not IsNull(rsTmp!Fecha_Corte) Then
       strLinea(9) = "Prox.Pago..:" & Format(rsTmp!Fecha_Corte, "dd/mm/yyyy") & " Cta.(" & rsTmp!Linea & ") " & Format(rsTmp!Monto, "Standard")
  Else
       strLinea(9) = "Prox.Pago..: >> <<"
  End If
  strLinea(10) = "Notas: " & rsTmp!Notas & ""
rsTmp.Close
      
strLinea(10) = Mid(strLinea(10), 1, 80)
      

If dtpFechaCancelacion.Enabled Then
   strLinea(11) = "Fecha Real Abono  ..: " & Format(dtpFechaCancelacion.Value, "dd/mm/yyyy")
End If

'Registro del Comprobante
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
         & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
         & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento)" _
         & " values('" & pNumDoc & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
         & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo & ",'P','" & txtOperacion.Text _
         & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
         & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
         & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
         & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
         & strLinea(11) & "','" & vAseDocDetalle & "','" & vAseDocDeposito & "')"
 Call ConectionExecute(strSQL)
 
 'ASIENTO
 If curIntC > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntC * pTipoCambio & ",'C','" & rs!cod_Divisa _
          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintc _
          & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
 
 If curIntM > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntM * pTipoCambio & ",'C','" & rs!cod_Divisa _
          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintm _
          & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
 
 If curCargo > 0 Then
 'Detallar Cargos
   strSQL = "exec spCxC_DocumentoAfectacionCargos '" & pTipoDoc & "','" & pNumDoc & "'"
   Call OpenRecordSet(rsTmp, strSQL, 0)
   Do While Not rsTmp.EOF
         strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rsTmp!Monto * pTipoCambio & ",'C','" & rs!cod_Divisa _
                & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!Cod_Centro_Costo & "','" & rsTmp!cod_cuenta _
                & "','" & rsTmp!Operacion & "','" & rsTmp!cod_Concepto & "','" & vAseDocDeposito & "'"
         Call ConectionExecute(strSQL)
         rsTmp.MoveNext
   Loop
   rsTmp.Close
 End If
 

 If curAmortiza > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curAmortiza * pTipoCambio & ",'C','" & rs!cod_Divisa _
          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
          & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
 
  
  If curIntC + curIntM + curCargo + curAmortiza > 0 Then
     'Procesa Formas de Pago (Registro Final / Asiento de Pago)
      strSQL = "exec spCajas_DesglocePagosDocFinal '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ",'" & ModuloCajas.mTiquete _
              & "','" & ModuloCajas.mUsuario & "','" & pTipoDoc & "','" & pNumDoc & "','" & ModuloCajas.mUnidad _
              & "','" & rs!Operacion & "','" & rs!cod_Concepto & "'"
      Call ConectionExecute(strSQL)
 End If

rs.Close


End Sub


Private Sub txtTotalPagar_GotFocus()

On Error GoTo vError
 txtTotalPagar.Text = CCur(txtTotalPagar.Text)
vError:

End Sub

Private Sub txtTotalCajas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

End Sub

Private Sub txtTotalPagar_LostFocus()
Dim curInteres As Currency, curAmortiza As Currency
Dim i As Integer, vChecks As Boolean
 
On Error GoTo vError
 
'ExtraOrdinario
If optAbono.Item(1).Value = True Then
   'Cobra intereses desde el ultimo corte
    txtTotalCancela.Text = Format(txtTotalPagar.Text, "Standard")
    curInteres = (CCur(txtTotalPagar.Text) * vInteres / 36000) * vDiasActivo
   'Se re-calculan intereses para ajustar y relacionar segun porcion amortizada
   'Previamente sobre el monto a cancelar
   
   If curInteres > 0 Then
      'Hacer 10 aproximaciones
      For i = 1 To 10
            curAmortiza = CCur(txtTotalPagar.Text) - curInteres
            curInteres = (curAmortiza * vInteres / 36000) * vDiasActivo
      Next i
   End If
   
   txtDatosInteres.Text = Format(curInteres, "Standard")
   txtDatosAmortiza.Text = Format(CCur(txtTotalPagar.Text) - curInteres, "Standard")
End If

txtTotalPagar.Text = Format(CCur(txtTotalPagar.Text), "Standard")

txtDiferencia.Text = Format(CCur(txtTotalCancela.Text) - CCur(txtTotalPagar.Text), "Standard")

cboDiferenciaApl.Enabled = False

If CCur(txtDiferencia.Text) < 0 Then
    Select Case True
      Case optAbono.Item(0).Value 'Abono Ordinario
           cboDiferenciaApl.Enabled = False
           'Verifica el Plazo sea menor que la ultima cuota marcada y que se hayan marcado todas con corte igual o menor a la fecha actual
           vChecks = True
           For i = 1 To lsw.ListItems.Count
             If Not lsw.ListItems.Item(i).Checked And DateDiff("d", CDate(lsw.ListItems.Item(i).SubItems(1)), vFechaHoy) >= 0 Then
               vChecks = False
             End If
           Next i
           
           'Verifica el Ultimo Plazo
           If vChecks Then
              If CCur(lblSaldoR.Caption) > CCur(txtDiferencia.Text) Then
                  cboDiferenciaApl.Enabled = True
              End If
           End If
           
      Case optAbono.Item(1).Value 'Abono ExtraOrdinario
           cboDiferenciaApl.Enabled = False
      
      Case optAbono.Item(2).Value 'Cancelacion
           cboDiferenciaApl.Enabled = False
      
      Case optAbono.Item(3).Value 'Adelanto
           'Verifica si el Saldo Resultante del Credito es mayor igual a la diferencia.
              If CCur(lblSaldoR.Caption) > CCur(txtDiferencia.Text) Then
                  cboDiferenciaApl.Enabled = True
              End If
    
    End Select
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

