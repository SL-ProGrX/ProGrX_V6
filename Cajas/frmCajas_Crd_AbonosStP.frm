VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCajas_Crd_AbonosStP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cajas..Abonos"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEstado 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6012
      Left            =   9720
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   9495
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2652
         Left            =   240
         TabIndex        =   26
         Top             =   2640
         Width           =   9012
         _Version        =   1572864
         _ExtentX        =   15896
         _ExtentY        =   4678
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
      Begin XtremeSuiteControls.FlatEdit txtCuotasMarcadas 
         Height          =   312
         Left            =   7560
         TabIndex        =   28
         Top             =   5400
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   9492
         _Version        =   1572864
         _ExtentX        =   16743
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Resumen en cambio de Estado y Cuotas pendientes:"
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total Cuotas Marcadas"
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
         Left            =   5400
         TabIndex        =   25
         Top             =   5400
         Width           =   2172
      End
      Begin VB.Label Label7 
         Caption         =   "Cuotas Pendientes ....:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   1692
      End
      Begin VB.Label lblFecUltMov 
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
         Left            =   1800
         TabIndex        =   22
         Top             =   1800
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
         Left            =   1800
         TabIndex        =   21
         Top             =   1560
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
         Left            =   1800
         TabIndex        =   20
         Top             =   1080
         Width           =   1572
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
         Left            =   1800
         TabIndex        =   19
         Top             =   840
         Width           =   1572
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
         Left            =   1800
         TabIndex        =   18
         Top             =   1320
         Width           =   1572
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ult.Mov."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   600
         TabIndex        =   17
         ToolTipText     =   "Si es menor a la fecha de proceso se Utiliza la Fecha de Proceso"
         Top             =   1800
         Width           =   1212
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   600
         TabIndex        =   16
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intereses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   600
         TabIndex        =   15
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amortización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   600
         TabIndex        =   14
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   600
         TabIndex        =   13
         Top             =   1320
         Width           =   1212
      End
      Begin VB.Label Label7 
         Caption         =   "Estado Actual ....:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   468
         Width           =   1692
      End
      Begin VB.Label Label7 
         Caption         =   "Estado Resultante....:"
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
         Height          =   252
         Index           =   1
         Left            =   3720
         TabIndex        =   11
         Top             =   480
         Width           =   1692
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4320
         TabIndex        =   10
         Top             =   1320
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
         Left            =   5520
         TabIndex        =   9
         Top             =   1320
         Width           =   1572
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amortización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4320
         TabIndex        =   8
         Top             =   840
         Width           =   1212
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
         Left            =   5520
         TabIndex        =   7
         Top             =   840
         Width           =   1572
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intereses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4320
         TabIndex        =   6
         Top             =   1080
         Width           =   1212
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
         Left            =   5520
         TabIndex        =   5
         Top             =   1080
         Width           =   1572
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4320
         TabIndex        =   4
         Top             =   1560
         Width           =   1212
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
         Left            =   5520
         TabIndex        =   3
         Top             =   1560
         Width           =   1572
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ult.Mov."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4320
         TabIndex        =   2
         Top             =   1800
         Width           =   1212
      End
      Begin VB.Label lblFecUltMovR 
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
         Left            =   5520
         TabIndex        =   1
         Top             =   1800
         Width           =   1572
      End
   End
   Begin XtremeSuiteControls.GroupBox fraAbono 
      Height          =   1092
      Left            =   120
      TabIndex        =   72
      Top             =   1680
      Width           =   9492
      _Version        =   1572864
      _ExtentX        =   16743
      _ExtentY        =   1926
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.DateTimePicker dtpFechaCancelacion 
         Height          =   312
         Left            =   6360
         TabIndex        =   74
         Top             =   600
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
      Begin XtremeSuiteControls.PushButton optAbono 
         Height          =   372
         Index           =   0
         Left            =   1560
         TabIndex        =   76
         Top             =   120
         Width           =   1572
         _Version        =   1572864
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
         TabIndex        =   77
         Top             =   120
         Width           =   1572
         _Version        =   1572864
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
         TabIndex        =   78
         Top             =   120
         Width           =   1572
         _Version        =   1572864
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
         TabIndex        =   75
         Top             =   600
         Width           =   4692
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   372
         Left            =   0
         TabIndex        =   73
         Top             =   120
         Width           =   9492
         _Version        =   1572864
         _ExtentX        =   16743
         _ExtentY        =   656
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
      Height          =   3252
      Left            =   120
      TabIndex        =   49
      Top             =   2760
      Width           =   9492
      _Version        =   1572864
      _ExtentX        =   16743
      _ExtentY        =   5736
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkRecalculaCuota 
         Height          =   252
         Left            =   6480
         TabIndex        =   58
         Top             =   1200
         Width           =   1932
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtCuotas 
         Height          =   312
         Left            =   2400
         TabIndex        =   59
         Top             =   360
         Width           =   852
         _Version        =   1572864
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosAmortiza 
         Height          =   312
         Left            =   1560
         TabIndex        =   60
         Top             =   720
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosInteres 
         Height          =   312
         Left            =   1560
         TabIndex        =   61
         Top             =   1080
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosCargos 
         Height          =   312
         Left            =   1560
         TabIndex        =   62
         Top             =   1440
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosAnticipo 
         Height          =   312
         Left            =   1560
         TabIndex        =   63
         Top             =   2040
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPolizas 
         Height          =   312
         Left            =   1560
         TabIndex        =   64
         Top             =   2640
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCompromiso 
         Height          =   312
         Left            =   6840
         TabIndex        =   65
         Top             =   2040
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDiferencia 
         Height          =   312
         Left            =   6840
         TabIndex        =   66
         Top             =   2640
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMoraCuotas 
         Height          =   312
         Left            =   3240
         TabIndex        =   67
         Top             =   360
         Width           =   852
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtMoraAmortiza 
         Height          =   312
         Left            =   3240
         TabIndex        =   68
         Top             =   720
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtMoraIntereses 
         Height          =   312
         Left            =   3240
         TabIndex        =   69
         Top             =   1080
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtMoraCargos 
         Height          =   312
         Left            =   3240
         TabIndex        =   70
         Top             =   1440
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
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
         Left            =   360
         TabIndex        =   55
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label Label26 
         Caption         =   "Intereses"
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
         Left            =   360
         TabIndex        =   54
         Top             =   1080
         Width           =   1092
      End
      Begin VB.Label Label25 
         Caption         =   "Amortización"
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
         Left            =   360
         TabIndex        =   53
         Top             =   720
         Width           =   1212
      End
      Begin VB.Label Label26 
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
         Index           =   2
         Left            =   360
         TabIndex        =   52
         Top             =   1440
         Width           =   1092
      End
      Begin VB.Label Label24 
         Caption         =   "Compromiso ..:"
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
         Left            =   6120
         TabIndex        =   51
         Top             =   1800
         Width           =   1932
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
         Left            =   6120
         TabIndex        =   50
         Top             =   2400
         Width           =   1692
      End
      Begin VB.Label Label26 
         Caption         =   "Cuotas de las Pólizas Asociadas..:"
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
         Index           =   3
         Left            =   360
         TabIndex        =   57
         Top             =   2400
         Width           =   2652
      End
      Begin VB.Label Label26 
         Caption         =   "Cargo por Cancelación Anticipada..:"
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
         Index           =   1
         Left            =   360
         TabIndex        =   56
         Top             =   1800
         Width           =   2652
      End
   End
   Begin XtremeSuiteControls.GroupBox fraFormaPago 
      Height          =   1572
      Left            =   120
      TabIndex        =   39
      Top             =   6000
      Width           =   9492
      _Version        =   1572864
      _ExtentX        =   16743
      _ExtentY        =   2773
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipoDoc 
         Height          =   312
         Left            =   1200
         TabIndex        =   40
         Top             =   240
         Width           =   2772
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalCajas 
         Height          =   312
         Left            =   4920
         TabIndex        =   44
         Top             =   240
         Width           =   1692
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1200
         TabIndex        =   45
         Top             =   600
         Width           =   5412
         _Version        =   1572864
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   732
         Index           =   0
         Left            =   6720
         TabIndex        =   46
         Top             =   600
         Width           =   852
         _Version        =   1572864
         _ExtentX        =   1503
         _ExtentY        =   1291
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
         Appearance      =   16
         Picture         =   "frmCajas_Crd_AbonosStP.frx":0000
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   732
         Index           =   1
         Left            =   7680
         TabIndex        =   47
         Top             =   600
         Width           =   852
         _Version        =   1572864
         _ExtentX        =   1503
         _ExtentY        =   1291
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
         Appearance      =   16
         Picture         =   "frmCajas_Crd_AbonosStP.frx":0462
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   732
         Index           =   2
         Left            =   8520
         TabIndex        =   48
         Top             =   600
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   1291
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
         Appearance      =   16
         Picture         =   "frmCajas_Crd_AbonosStP.frx":0C3A
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   240
         Width           =   1452
      End
   End
   Begin XtremeSuiteControls.PushButton btnEstado 
      Height          =   492
      Left            =   8160
      TabIndex        =   29
      Top             =   1080
      Width           =   1452
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Estado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCajas_Crd_AbonosStP.frx":1407
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8400
      Top             =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Crd_AbonosStP.frx":1BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Crd_AbonosStP.frx":8448
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Crd_AbonosStP.frx":8C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Crd_AbonosStP.frx":93ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Crd_AbonosStP.frx":9CB9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3120
      TabIndex        =   30
      Top             =   960
      Width           =   4812
      _Version        =   1572864
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
      Left            =   3120
      TabIndex        =   31
      Top             =   1320
      Width           =   4212
      _Version        =   1572864
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
      Left            =   1440
      TabIndex        =   32
      Top             =   960
      Width           =   1692
      _Version        =   1572864
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
      Left            =   1440
      TabIndex        =   33
      Top             =   1320
      Width           =   1692
      _Version        =   1572864
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
      Left            =   7320
      TabIndex        =   36
      Top             =   1320
      Width           =   612
      _Version        =   1572864
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
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   2520
      TabIndex        =   37
      Top             =   240
      Width           =   2052
      _Version        =   1572864
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   372
      Left            =   4560
      TabIndex        =   38
      Top             =   240
      Width           =   2052
      _Version        =   1572864
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   492
      Left            =   8160
      TabIndex        =   71
      Top             =   360
      Visible         =   0   'False
      Width           =   1452
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   868
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
      Appearance      =   16
      Picture         =   "frmCajas_Crd_AbonosStP.frx":A6D7
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   35
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   34
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   1080
      TabIndex        =   23
      Top             =   240
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmCajas_Crd_AbonosStP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOperacion As Long, vCuotasDeducidas As Integer, vCuotasDirectas As Integer
Dim vInteres As Currency, vPlazo As Integer, vSaldoMes As Currency, vUltimoRecibo As Long
Dim vRetencion As Boolean, vBaseCalculo As String, vPrideduc As Currency
Dim pDatos() As Currency, vFechaHoy As Date
Dim pCharRelleno As String
Dim vTipoAbono As Integer



Private Sub btnCajas_Click(Index As Integer)
On Error GoTo vError

Select Case Index
  Case 2 'Cancelar
     Call sbConsultaOperacion
     
  Case 0 'Desgloce de Pago
        If Not IsNumeric(txtCompromiso.Text) Then txtCompromiso.Text = 0
        If Not ModuloCajas.mConceptoValida Then
           MsgBox "Esta caja no está autorizada para registrar movimientos a este línea de crédito", vbExclamation
           Exit Sub
        End If
        
        ModuloCajas.mTotalAplicar = CCur(txtCompromiso.Text)
        
        If ModuloCajas.mTotalAplicar = 0 Then
            MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
            Exit Sub
        End If
        
        ModuloCajas.mServicio = "Abonos a Operación de Crédito"
        
        Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)
        
        txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")

  Case 1 'Aplicar
    Call CmdAplicar_Click
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnEstado_Click()
fraEstado.top = 1680
fraEstado.Left = 120

If fraEstado.Visible Then
   fraEstado.Visible = False
Else
   fraEstado.Visible = True
End If

End Sub

Private Sub chkRecalculaCuota_Click()

If vRetencion Then
   chkRecalculaCuota.Value = vbUnchecked
   MsgBox "Las retenciones no se pueden Ajustar para Recálculos, verifique...", vbExclamation
   Exit Sub
End If

'Call txtTotalCajas_Change

End Sub

Private Sub sbAbono()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNumDoc  As String, vTipoDoc As String
Dim vTipo As String, vFecha As Date, vCuenta As String
Dim vFechaProceso As Long, i As Integer, vConcepto As String


Me.MousePointer = vbHourglass

On Error GoTo vError

vNumDoc = 0
vFecha = fxFechaServidor

'Configuracion del Documento
vTipoDoc = cboTipoDoc.ItemData(cboTipoDoc.ListIndex)
vNumDoc = fxDocumentoConsecutivo(vTipoDoc)

vCuenta = ""

If CLng(lblFecUltMovR) < GLOBALES.glngFechaCR Then
  lblFecUltMovR.Caption = GLOBALES.glngFechaCR
End If
vFechaProceso = lblFecUltMovR.Caption


'Genera el Comprobante
Select Case True
  Case optAbono(0).Checked 'Abono Ordinario
      vConcepto = "CRD001"
      vTipo = "O"
  
  Case optAbono(1).Checked  'Abono Extraordinario
      vConcepto = "CRD002"
      vTipo = "E"
      vFechaProceso = Format(dtpFechaCancelacion.Value, "yyyymm")
  
  Case optAbono(2).Checked  'Abono De Cancelacion
      vConcepto = "CRD003"
      vTipo = "O"
End Select


'Aplica Abono
strSQL = "exec spCajas_CrdAbono " & vOperacion & "," & CCur(txtTotalCajas.Text) & ",'" & vTipoDoc & "','" & vNumDoc & "','" & vConcepto _
       & "','" & ModuloCajas.mUsuario & "','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura _
       & "," & chkRecalculaCuota.Value & "," & CCur(txtDatosAnticipo.Text)
       
       
       
If vTipo = "E" Then
   strSQL = strSQL & "," & CCur(txtDatosInteres.Text) & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
Else
  strSQL = strSQL & ",0,'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
End If

Call OpenRecordSet(rs, strSQL)
If rs!Pendiente <> 0 Then
    Me.MousePointer = vbDefault
    MsgBox "Quedó un monto pendiente de :" & Format(rs!Pendiente, "Standard"), vbExclamation
    Me.MousePointer = vbHourglass
End If
rs.Close

'Genera el Comprobante
Select Case True
  Case optAbono(0).Checked  'Abono Ordinario
      Call Bitacora("Registra", "Abono Ordinario a la Operacion : " & vOperacion)
      Call sbDocumentoAbono("ABONO ORDINARIO", vTipoDoc, vNumDoc, vConcepto, vCuenta)
  
  Case optAbono(1).Checked 'Abono Extraordinario
      Call Bitacora("Registra", "Abono ExtraOrd. " & IIf((chkRecalculaCuota.Value = 1), "Con Recal.", "Sin Recal") & " a la Op.: " & vOperacion)
      Call sbDocumentoAbono("ABONO EXTRAORDINARIO", vTipoDoc, vNumDoc, vConcepto, vCuenta)
  
  Case optAbono(2).Checked 'Abono De Cancelacion
      Call Bitacora("Registra", "Cancelación de la Operacion : " & vOperacion)
      Call sbDocumentoAbono("CANCELACION DE DEUDA", vTipoDoc, vNumDoc, vConcepto, vCuenta)
End Select

'IMPRIMIR RECIBO
Call sbImprimeRecibo(vNumDoc, vTipoDoc)

Me.MousePointer = vbDefault

strSQL = " - Abono aplicado, con : " & cboTipoDoc.Text & " ...No.: " & vNumDoc & vbCrLf _
       & " - Desea Realizar Otra Transacción a Este Operación ?"

i = MsgBox(strSQL, vbYesNo)
If i = vbYes Then
    Call sbConsultaOperacion
    txtTotalCajas.Text = 0
Else
    Unload Me
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
' glogon.Conection.RollbackTrans
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxVerifica() As Boolean
Dim vMensaje As String

vMensaje = ""

Call sbSIFCleanTxtInject(txtNotas)

'Verifica el proceso
If txtProceso.Tag = "J" Then
   If Not fxCajasAbonosCbrJud(ModuloCajas.mCaja, ModuloCajas.mUsuario) Then
      vMensaje = vMensaje & "- Esta CAJA no cuenta con permisos para realizar abonos a Creditos en Cobro Judicial, verifique..." & vbCrLf
   End If
End If


'Verificar Congelamiento
If fxgCongelamiento(txtCedula, "per_abono_cajas") Then
  vMensaje = vMensaje & "- Esta Persona se encuentra CONGELADA, verifique..." & vbCrLf
End If

If vOperacion = 0 Then
  vMensaje = vMensaje & "- Número de Operacion no es válido..." & vbCrLf
End If
 
 
'Verifica Saldo Actual
If Not fxCrdSaldoVerifica(vOperacion, CCur(lblSaldo.Caption)) Then
   vMensaje = vMensaje & "- Esta Operación ha sido modificada, actualice los datos nuevamente antes de realizar el abono..." & vbCrLf
End If
 
If Not vRetencion Then
    If (CCur(txtDatosAmortiza.Text) + CCur(txtMoraAmortiza.Text)) > CCur(lblSaldo.Caption) Then
       vMensaje = vMensaje & "- La Amortización es mayor al Saldo Actual..." & vbCrLf
    End If
Else
    If CCur(txtDatosAmortiza) > ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption)) And vPlazo <= 900 Then
        vMensaje = vMensaje & "- La Amortización es mayor que el Remanente a Recaudar : " _
              & ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption)) & vbCrLf
     End If
End If

If Not IsNumeric(txtCompromiso.Text) Then
  vMensaje = vMensaje & "- El compromiso de pago no es un dato válido...verifique...!" & vbCrLf
Else
 If CCur(txtCompromiso.Text) <= 0 Then
      vMensaje = vMensaje & "- El compromiso de pago no es un dato válido...verifique...!" & vbCrLf
 End If
End If

 If CCur(txtTotalCajas.Text) <= 0 Then
      vMensaje = vMensaje & "- Los valores Recibidos en Cajas no son válidos...verifique...!" & vbCrLf
 End If
 
 'No se permiten Abonos ordinarios parciales al menos que sea para abonos a cuotas atrasadas
 If optAbono(0).Checked And Len(vMensaje) = 0 Then

    If txtMoraCuotas.Text > 0 Then
        If CCur(txtCompromiso.Text) < CCur(txtTotalCajas.Text) _
             And CCur(txtTotalCajas.Text) > (CCur(txtMoraAmortiza.Text) + CCur(txtMoraCargos.Text) + CCur(txtMoraIntereses.Text)) Then
                vMensaje = vMensaje & "- Tiene que hacer una transaccion solo por la mora y luego otro extraordinario...!" & vbCrLf
        End If
    Else
      'Ordinario al día
      If CCur(txtCompromiso.Text) > CCur(txtTotalCajas.Text) Then
                vMensaje = vMensaje & "- Tiene que hacer una transaccion Cancelando el Compromiso...!" & vbCrLf
      End If
    
    End If 'txtMoraCuotas.Text > 0
 End If


 'Validacion para Abonos Extraordinarios
 If optAbono(1).Checked And Len(vMensaje) = 0 Then
    If CCur(txtDiferencia.Text) <> 0 Then
       vMensaje = vMensaje & "- El Monto detallado en formas de pago no cubre el compromiso de pago. SOLUCION: Copie el Monto detallado en el monto del abono extraordinario...!" & vbCrLf
    End If
 End If


 'Validacion para Cancelacion de Deudas
 If optAbono(2).Checked And Len(vMensaje) = 0 Then
    If CCur(txtDiferencia.Text) <> 0 Then
       vMensaje = vMensaje & "- El Monto detallado en formas de pago no cubre el compromiso de cancelación...!" & vbCrLf
    End If
 End If


If fxCajasAperturaEstado = "C" Then
   vMensaje = vMensaje & "- La apertura ..:" & ModuloCajas.mApertura & " de esta caja ha sido cerrada!" & vbCrLf
End If

If Len(vMensaje) = 0 Then
  fxVerifica = True
Else
  fxVerifica = False
  MsgBox vMensaje, vbExclamation
End If

End Function


Private Sub CmdAplicar_Click()
Dim iRespuesta As Integer

If Not fxVerifica Then Exit Sub

 iRespuesta = MsgBox("Esta seguro de realizar el abono a esta Operación " & vOperacion, vbYesNo)
 If iRespuesta = vbYes Then
  
  Call sbAbono
 
 Else 'Respuesta
  
  MsgBox "Transacción Cancelada...", vbInformation
 
 End If

End Sub


Private Sub dtpFechaCancelacion_Change()

Call sbCargaGridMorosidad

If dtpFechaCancelacion.Enabled Then
   Select Case True
      Case optAbono.Item(1).Checked 'Abono Extraordinario
            Call optAbono_Click(1)
      Case optAbono.Item(2).Checked 'Cancelación
            Call optAbono_Click(2)
   End Select
End If
End Sub

Private Sub Form_Activate()
 vModulo = 5

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

Me.Caption = "Abonos a Créditos     ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

txtTotalCajas.Text = 0
txtNotas.Text = ""
strSQL = "select rTrim(C.tipo_documento) as 'IdX', rtrim(D.Descripcion) as 'itmX'" _
       & " from SIF_DOCUMENTOS D inner join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
       & " Where C.cod_caja =  '" & ModuloCajas.mCaja & "' and D.Tipo_Movimiento in('A','C')" _
       & " order by C.tipo_documento"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)

ModuloCajas.mServicio = "Abonos a Operaciones de Crédito"

If IsNumeric(ModuloCajas.mRef_01) Then
    txtOperacion.Text = ModuloCajas.mRef_01
    vOperacion = txtOperacion.Text
    Call sbConsultaOperacion
End If


If ModuloCajas.mSesionId = 0 Then
   Call sbFormsCall("frmCajas_Sesion", vbModal, , , False, Me)
   If ModuloCajas.mSesionId = 0 Then
        MsgBox "No se ha iniciado ninguna sesión de Cliente para esta caja!", vbExclamation
        Unload Me
        Exit Sub
   End If
End If

End Sub

Private Sub Form_Load()
Dim iDias As Integer

vModulo = 5
vOperacion = 0
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "[Id]", 1500
    .Add , , "Operación", 1500
    .Add , , "Proceso", 1100, vbCenter
    .Add , , "Int.Cor.", 1500, vbRightJustify
    .Add , , "Int.Mor.", 1500, vbRightJustify
    .Add , , "Amortización", 1500, vbRightJustify
    .Add , , "Cargos", 1500, vbRightJustify
    .Add , , "Total", 1500, vbRightJustify
End With
 
vFechaHoy = fxFechaServidor
iDias = fxCrdParametro("32")

dtpFechaCancelacion.Value = vFechaHoy
dtpFechaCancelacion.MinDate = DateAdd("d", (iDias * -1), dtpFechaCancelacion.Value)
dtpFechaCancelacion.MaxDate = dtpFechaCancelacion.Value

Call Formularios(Me)
Call RefrescaTags(Me)

btnCajas.Item(1).Enabled = cmdAplicar.Enabled

End Sub

Private Sub sbCargaGridMorosidad()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTempCuota As Integer, vTempCargo As Currency
Dim vTempAmort As Currency, vTempIntCor As Currency
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

vTempCuota = 0
vTempAmort = 0
vTempCargo = 0
vTempIntCor = 0

strSQL = "exec spCajas_Crd_MoraConsulta " & vOperacion & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!id_moro)
      itmX.SubItems(1) = rs!Id_Solicitud
      itmX.SubItems(2) = Format(rs!fechap, "####-##")
      itmX.SubItems(3) = Format(rs!IntC, "Standard")
      itmX.SubItems(4) = Format(rs!IntM, "Standard")
      itmX.SubItems(5) = Format(rs!Amortiza, "Standard")
      itmX.SubItems(6) = Format(rs!Cargo, "Standard")
      itmX.SubItems(7) = Format(rs!IntC + rs!IntM + rs!Amortiza + rs!Cargo, "Standard")
      
      vTempCuota = vTempCuota + 1
      vTempAmort = vTempAmort + rs!Amortiza
      vTempCargo = vTempCargo + rs!Cargo
      vTempIntCor = vTempIntCor + rs!IntC + rs!IntM
      
      itmX.Checked = True
      
 rs.MoveNext
Loop
rs.Close

txtCuotasMarcadas.Text = Format(vTempAmort + vTempIntCor + vTempCargo, "Standard")

txtMoraCuotas.Text = vTempCuota
txtMoraAmortiza.Text = Format(vTempAmort, "Standard")
txtMoraIntereses.Text = Format(vTempIntCor, "Standard")
txtMoraCargos.Text = Format(vTempCargo, "Standard")

'Desactiva los abonos extraordinarios si está con cuotas pendientes.
If vTempCuota > 1 Then
    optAbono.Item(1).Enabled = False
Else
    optAbono.Item(1).Enabled = True
End If

'Saldo del Mes
If Not vRetencion Then
   vSaldoMes = CCur(lblSaldo.Caption) - vTempAmort
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub


Private Sub sbConsultaOperacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curSaldo As Currency

Me.MousePointer = vbHourglass

 
strSQL = "select R.id_solicitud,R.saldo, R.saldo - isnull(V.amortiza,0) As Saldo_mes,R.proceso, isnull(R.cod_Divisa,'COL') as 'Divisa'" _
       & ",R.interesv,R.int,R.plazo,R.interesc,R.amortiza,R.fecult,R.Prideduc" _
       & ",R.opex,R.cuota,R.codigo,R.cedula,R.cuotas_planilla,R.cuotas_directas,R.montoApr" _
       & ",S.nombre,C.descripcion,C.retencion,C.poliza,R.fechaforp,C.PORC_CARGO_CANCELACION,R.Base_Calculo" _
       & ",dbo.fxCajas_Valida_Auxiliar('" & ModuloCajas.mCaja & "','CRD',R.Codigo) as 'Caja_Valida_Concepto'" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " left join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " where R.estado = 'A' and R.saldo > 0" _
       & " and R.ID_SOLICITUD = " & vOperacion
       
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
    txtCodigo.Text = rs!Codigo
    txtDescripcion.Text = rs!Descripcion
    
    ModuloCajas.mClienteId = Trim(rs!Cedula)
    ModuloCajas.mCliente = Trim(rs!Nombre)
    ModuloCajas.mTiquete = Trim(rs!Codigo) & "." & rs!Id_Solicitud & "." & Format(Time, "HH:mm:ss")
    
    ModuloCajas.mDivisa = RTrim(rs!Divisa)
    ModuloCajas.mConceptoValida = IIf((rs!Caja_Valida_Concepto > 0), True, False)
    
    ModuloCajas.mTotalDetallado = 0
    txtTotalCajas.Text = 0

    vBaseCalculo = Trim(rs!Base_Calculo)
    vPrideduc = rs!PriDeduc
    vOperacion = rs!Id_Solicitud
    vPlazo = rs!Plazo
  
  
    vInteres = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
    If IsNull(rs!saldo_mes) Then
      vSaldoMes = rs!Saldo
      strSQL = "update reg_creditos set saldo_mes = saldo where id_solicitud = " & rs!Id_Solicitud
      Call ConectionExecute(strSQL)
    Else
      If rs!saldo_mes = 0 Then
          vSaldoMes = rs!Saldo
          strSQL = "update reg_creditos set saldo_mes = saldo where id_solicitud = " & rs!Id_Solicitud
          Call ConectionExecute(strSQL)
      Else
         vSaldoMes = rs!saldo_mes
      End If
    
    End If
  
    vCuotasDeducidas = IIf(IsNull(rs!cuotas_planilla), 0, rs!cuotas_planilla)
    vCuotasDirectas = IIf(IsNull(rs!cuotas_directas), 0, rs!cuotas_directas)
    
    lblAmortiza.Caption = Format(rs!Amortiza, "Standard")
    lblAmortizaR.Caption = 0
    lblCuota = Format(rs!Cuota, "Standard")
    lblCuotaR.Caption = 0
    
    txtDatosAmortiza = 0
    txtDatosInteres.Text = 0
    
    lblFecUltMov.Caption = IIf(IsNull(rs!FecUlt), fxFechaProcesoAnterior(GLOBALES.glngFechaCR), rs!FecUlt)
    
    If CLng(lblFecUltMov.Caption) < GLOBALES.glngFechaCR Then
      lblFecUltMov.Caption = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
    End If
    
    
    lblFecUltMovR.Caption = 0
    lblInteres.Caption = Format(rs!interesc, "Standard")
    lblInteresR.Caption = 0
    
    txtOpex.Text = IIf((rs!opex = 1), "OPEX", "")
    
    lblSaldo.Tag = rs!FechaForp
    lblSaldo.Caption = Format(rs!Saldo, "Standard")
    lblSaldoR.Caption = 0
    
    txtCuotas = 0
    txtOperacion = rs!Id_Solicitud
    
    fraAbono.Enabled = True
    fraDatosAbono.Enabled = False
    

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
    
    
    
    txtDatosAnticipo.ToolTipText = "% de Comision : " & rs!PORC_CARGO_CANCELACION
    txtDatosAnticipo.Tag = rs!PORC_CARGO_CANCELACION
    
    optAbono(0).Enabled = True
    optAbono(1).Enabled = True
    
    
    If rs!retencion = "S" Or rs!Poliza = "S" Then
      vRetencion = True
     If rs!Plazo < 900 Then
        lblSaldo.Caption = Format((rs!montoapr * rs!Plazo) - rs!Amortiza, "Standard")
        lblSaldoR.Caption = 0
        vSaldoMes = CCur(lblSaldo.Caption)
     End If
    Else
      vRetencion = False
    End If
        
        
    'Carga datos Mora antes del Click del Abono
    Call sbCargaGridMorosidad
        
        
        Select Case True
         Case optAbono(0).Checked
           Call optAbono_Click(0)
         Case optAbono(1).Checked
           Call optAbono_Click(1)
         Case optAbono(2).Checked
           Call optAbono_Click(2)
        End Select

Else
 
    vOperacion = 0
    vPlazo = 0
    vInteres = 0
    vSaldoMes = 0
    MsgBox "No se Encontró operación para abonos,puede que se encuentre cancelada ", vbInformation
    Call sbLimpiaDatos

End If
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub sbLimpiaDatos()
 
txtDatosAnticipo.Text = 0

lblAmortiza.Caption = 0
lblAmortizaR.Caption = 0
lblCuota = 0
lblCuotaR.Caption = 0
txtDatosAmortiza = 0
txtDatosInteres.Text = 0
lblFecUltMov.Caption = 0
lblFecUltMovR.Caption = 0
lblInteres.Caption = 0
lblInteresR.Caption = 0
txtDescripcion.Text = ""
txtOpex.Text = ""

lblSaldo.Caption = 0
lblSaldoR.Caption = 0
txtCedula = ""
txtCodigo = ""
txtCuotas = 0
txtNombre = ""
txtOperacion = ""

txtCompromiso.Text = 0

txtTotalCajas.Text = 0

txtProceso.Tag = ""
txtProceso.Text = ""

fraAbono.Enabled = False
fraDatosAbono.Enabled = False

chkRecalculaCuota.Value = vbUnchecked
 
dtpFechaCancelacion.Enabled = False
lblFechaCancelacion.Enabled = False
dtpFechaCancelacion.Value = vFechaHoy
 
End Sub

Private Sub sbBusqueda()

On Error GoTo vError

gBusquedas.Convertir = "N"
gBusquedas.Consulta = "Select R.id_solicitud as Operacion,R.Codigo,S.Cedula,S.Nombre,C.Descripcion" _
          & " from REG_CREDITOS R inner join SOCIOS S on R.cedula = S.cedula" _
          & " inner join Catalogo C on R.codigo = C.codigo"
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

Private Sub sbCargaOperacionCodCed(vCedula As String, vCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select R.id_solicitud,R.saldo,R.saldo_mes,R.interesv,R.int,R.plazo,R.interesc,R.amortiza,R.fecult" _
       & ",R.opex,R.cuota,R.codigo,R.cedula,R.cuotas_planilla,R.cuotas_directas,C.retencion,C.poliza " _
       & "from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & "where R.estado = 'A' and R.proceso <> 'N' and R.saldo > 0 " _
       & "and R.cedula = '" & txtCedula & "' and R.codigo = '" & txtCodigo & "'"
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  vOperacion = rs!Id_Solicitud
  vPlazo = rs!Plazo
  vInteres = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
  vSaldoMes = IIf(IsNull(rs!saldo_mes), rs!Saldo, rs!saldo_mes)
  vCuotasDeducidas = IIf(IsNull(rs!cuotas_planilla), 0, rs!cuotas_planilla)
  vCuotasDirectas = IIf(IsNull(rs!cuotas_directas), 0, rs!cuotas_directas)
     lblAmortiza.Caption = Format(rs!Amortiza, "Standard")
     lblAmortizaR.Caption = 0
     lblCuota = Format(rs!Cuota, "Standard")
     lblCuotaR.Caption = 0
     txtDatosAmortiza = 0
     txtDatosInteres.Text = 0
    
     lblFecUltMov.Caption = IIf(IsNull(rs!FecUlt), fxFechaProcesoAnterior(GLOBALES.glngFechaCR), rs!FecUlt)
    If CLng(lblFecUltMov.Caption) < GLOBALES.glngFechaCR Then
       lblFecUltMov.Caption = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
    End If
     lblFecUltMovR.Caption = 0
     lblInteres.Caption = Format(rs!interesc, "Standard")
     lblInteresR.Caption = 0
     txtOpex.Text = IIf((rs!opex = 1), "OPEX", "")
     lblSaldo.Caption = Format(vSaldoMes, "Standard")
     lblSaldoR.Caption = 0
     txtCuotas = 0
     txtOperacion = rs!Id_Solicitud
     fraAbono.Enabled = True
     fraDatosAbono.Enabled = False
    
    optAbono(0).Enabled = True
    optAbono(1).Enabled = True
    
    
    If rs!retencion = "S" Or rs!Poliza = "S" Then
      vRetencion = True
    Else
      vRetencion = False
    End If
        
    If Not vRetencion Then
        Select Case True
         Case optAbono(0).Checked
           Call optAbono_Click(0)
           
         Case optAbono(1).Checked
           Call optAbono_Click(1)
        End Select
    Else
           Call optAbono_Click(0)
           optAbono(1).Enabled = False
    End If
    
    
Else
 
 vOperacion = 0
 vPlazo = 0
 vInteres = 0
 vSaldoMes = 0
 MsgBox "No se Encontrarón operaciones para abonos con esta cédula y código", vbInformation
End If
rs.Close

End Sub










Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim i As Integer, curMonto As Currency

On Error GoTo vError

If Item.Checked Then
    curMonto = CCur(Item.SubItems(7))
Else
    curMonto = CCur(Item.SubItems(7)) * -1
End If

If Not IsNumeric(txtCuotasMarcadas.Text) Then
    txtCuotasMarcadas.Text = 0
End If

txtCuotasMarcadas.Text = Format(CCur(txtCuotasMarcadas.Text) + curMonto, "Standard")

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tblDesgloce_ButtonClick(ByVal Button As MSComctlLib.Button)
If Not IsNumeric(txtCompromiso.Text) Then txtCompromiso.Text = 0

ModuloCajas.mTotalAplicar = CCur(txtCompromiso.Text)

If ModuloCajas.mTotalAplicar = 0 Then
    MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
    Exit Sub
End If

ModuloCajas.mServicio = "Abonos a Operación de Crédito"

Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)

txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")

End Sub

Private Sub opt_Click(Index As Integer)

End Sub

Private Sub TimerX_Timer()

TimerX.Enabled = False
TimerX.Interval = 0

Call sbCajaInicial

If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   Unload Me
   Exit Sub
End If

If ModuloCajas.mSesionId = 0 Or ModuloCajas.mSesionId = Empty Then
   Unload Me
   Exit Sub
End If

Call optAbono_Click(0)
End Sub



Private Sub txtCompromiso_Change()
Dim strSQL As String, rs As New ADODB.Recordset
Dim ProcesosTmp As Currency, lngFecha As Currency, iPlazoRst As Integer, curCuota As Currency

On Error Resume Next

If chkRecalculaCuota.Value = vbChecked Then
  
       lngFecha = lblFecUltMovR.Caption
       If lngFecha < vPrideduc Then lngFecha = vPrideduc
      
       ProcesosTmp = vPrideduc
       iPlazoRst = 1
        Do While ProcesosTmp < lngFecha
          ProcesosTmp = fxFechaProcesoSiguiente(ProcesosTmp)
          iPlazoRst = iPlazoRst + 1
        Loop
       iPlazoRst = vPlazo - iPlazoRst
       
       If iPlazoRst <= 0 Then iPlazoRst = 1
       
       curCuota = fxCalcula_Cuota(CDbl(lblSaldoR.Caption), iPlazoRst, vInteres)
       lblCuotaR.Caption = Format(curCuota, "Standard")
Else
  lblCuotaR.Caption = lblCuota.Caption
End If

txtDiferencia.Text = Format(CCur(txtTotalCajas.Text) - CCur(txtCompromiso.Text), "Standard")

End Sub

Private Sub txtCompromiso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTotalCajas.SetFocus
End Sub

Private Sub txtTotalCajas_Change()
On Error GoTo vError

txtDiferencia.Text = Format(CCur(txtTotalCajas.Text) - CCur(txtCompromiso.Text), "Standard")

vError:
End Sub

Private Sub optAbono_Click(Index As Integer)
Dim curInteres As Currency, vFecha As Date
Dim vProceso As Long, i As Integer


vTipoAbono = Index

fraDatosAbono.Enabled = True

chkRecalculaCuota.Enabled = False
chkRecalculaCuota.Value = vbUnchecked

txtCompromiso.BackColor = &HC0FFC0
txtCuotas.BackColor = txtCompromiso.BackColor

txtCompromiso.Locked = True
txtCuotas.Enabled = False

dtpFechaCancelacion.Enabled = False
lblFechaCancelacion.Enabled = False


For i = 0 To 2
  If i = Index Then
      optAbono.Item(i).Checked = True
  Else
      optAbono.Item(i).Checked = False
  End If
Next i

Select Case Index

 Case 0 'Ordinario
   txtDatosAnticipo.Text = 0
   txtDatosAmortiza = 0
   txtCuotas.Enabled = True
   txtCuotas.Text = 0 'Inicializa
   txtCuotas.Text = 1 'Inicializa
   txtCuotas.SetFocus
   
   txtCuotas.BackColor = vbWhite
   
 Case 1 'Extraordinario
 
   lblFechaCancelacion.Caption = "Fecha de Abono (Real) por parte del cliente para Ab.Extraordinario:"
   dtpFechaCancelacion.Enabled = True
   lblFechaCancelacion.Enabled = True
 
   txtCuotas = 0
   txtDatosInteres.Text = 0
   txtDatosAnticipo.Text = 0
   txtDatosAmortiza = 0
   txtCuotas.Enabled = False
   
   txtCompromiso.BackColor = vbWhite
   txtCompromiso.Locked = False
   txtCompromiso.SetFocus
   
   chkRecalculaCuota.Value = vbChecked
   chkRecalculaCuota.Enabled = True
   
Case 2 'Cancelacion
   'Le Calcula los intereses del proceso mensual + el saldo
   
   lblFechaCancelacion.Caption = "Fecha de Abono (Real) por parte del cliente para cancelación...:"
   dtpFechaCancelacion.Enabled = True
   lblFechaCancelacion.Enabled = True
   
   txtDatosAmortiza = 0
   txtCuotas.Enabled = False
   txtCuotas.Text = 0 'Inicializa
   txtCuotas.Text = 1 'Inicializa
'   txtCuotas.SetFocus
   
   txtDatosAmortiza = Format(vSaldoMes, "Standard")
   
   'Cobra intereses del mes, pero verificar la fecha de proceso que sea igual
   'o menor
   vFecha = dtpFechaCancelacion.Value
   vProceso = Year(vFecha) & Format(Month(vFecha), "00")
   
   
   If (vProceso >= vPrideduc) And (vProceso > CLng(lblFecUltMov.Caption)) Then
     curInteres = (vSaldoMes * vInteres / 36000) * Day(vFecha)
   Else
     curInteres = 0
   End If
   
   
   '3er Paso de Validacion de Pago de Intereses
   'Verifica que no sea un credito del mismo mes
   If curInteres > 0 And Month(CDate(lblSaldo.Tag)) = Month(vFecha) _
        And Year(CDate(lblSaldo.Tag)) = Year(vFecha) Then
      curInteres = 0
   End If
   
   If vRetencion Then
      txtDatosAnticipo.Text = "0.00"
   Else
      txtDatosAnticipo.Text = Format(vSaldoMes * (CCur(txtDatosAnticipo.Tag) / 100), "Standard")
   End If
   
   txtDatosInteres.Text = Format(curInteres, "Standard")
   txtCompromiso.Text = Format(CCur(lblSaldo.Caption) + curInteres + CCur(txtMoraIntereses.Text) + CCur(txtMoraCargos.Text) + CCur(txtDatosAnticipo.Text), "Standard")
   
End Select


lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - (CCur(txtDatosAmortiza) + CCur(txtMoraAmortiza.Text)), "Standard")

Call RefrescaTags(Me)

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtNombre = fxNombre(txtCedula)
  If txtCodigo <> "" Then Call sbCargaOperacionCodCed(txtCedula, txtCodigo)
  txtCodigo.SetFocus
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtCodigo = UCase(txtCodigo)
  txtDescripcion.Text = fxDescribeCodigo(txtCodigo)
  If txtCedula <> "" Then Call sbCargaOperacionCodCed(txtCedula, txtCodigo)
  txtOperacion.SetFocus
End If

End Sub

Private Sub txtCuotas_Change()
Dim curSaldo As Currency, curAmortiza As Currency, curInteres As Currency
Dim curTmpAmortiza As Currency, curTmpInteres As Currency, i As Integer
Dim lngFecha As Currency, lngCuotas As Long, lngCuotaMaxima As Long


Dim iDias As Integer, vFecha As Date, curCuota As Currency, iPlazoRst As Integer, ProcesosTmp As Currency

On Error Resume Next

If txtCuotas = "" Or Not IsNumeric(txtCuotas.Text) Then
 lngCuotas = 0
Else
 lngCuotas = txtCuotas
End If

If txtOperacion.Text = "" Then Exit Sub


ReDim pDatos(lngCuotas, 5) As Currency


lngFecha = CCur(lblFecUltMov.Caption)

If Not vRetencion Then
    curSaldo = vSaldoMes
Else
  'En las retenciones hay que proyectar el saldo del mes
  If vPlazo > 900 Then
      curSaldo = CCur(lblCuota.Caption) * 20 'Soporta 20 Cuotas Retencion Indefinida
  Else
      curSaldo = ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption))
  End If
End If

curAmortiza = 0
curInteres = 0
curCuota = lblCuota.Caption


If lngFecha < vPrideduc Then lngFecha = fxFechaProcesoAnterior(vPrideduc)

If vBaseCalculo = "01" Then
    For i = 1 To lngCuotas
    '360 / 360
        If curSaldo > 0 Then
          
          lngCuotaMaxima = i
          curTmpInteres = (curSaldo * vInteres) / 1200
          curTmpAmortiza = CCur(lblCuota.Caption) - curTmpInteres
          
          curAmortiza = curAmortiza + curTmpAmortiza
          curInteres = curInteres + curTmpInteres
          
          curSaldo = curSaldo - curTmpAmortiza
          lngFecha = fxFechaProcesoSiguiente(lngFecha)
        
          pDatos(i, 1) = curTmpInteres
          pDatos(i, 2) = curTmpAmortiza
          pDatos(i, 3) = lngFecha
          pDatos(i, 4) = curSaldo
          pDatos(i, 5) = curCuota
        
        End If
        
        If curSaldo < 0 Then
            pDatos(i, 1) = 0
            pDatos(i, 2) = curSaldo
            pDatos(i, 3) = lngFecha
            pDatos(i, 4) = 0
            pDatos(i, 5) = curSaldo
           
           curAmortiza = curAmortiza + curSaldo
           curSaldo = 0
        End If
     Next i
 
 Else
   '365 / 360
   
       'Calcula el Plazo Restante
       ProcesosTmp = vPrideduc
       iPlazoRst = 0
        Do While ProcesosTmp < lngFecha
          ProcesosTmp = fxFechaProcesoSiguiente(ProcesosTmp)
          iPlazoRst = iPlazoRst + 1
        Loop
       iPlazoRst = vPlazo - iPlazoRst
       
       'Saca el formato fecha del ultimo movimiento para calculo de dias
       vFecha = Mid(CStr(lngFecha), 1, 4) & "/" & Right(CStr(lngFecha), 2) & "/01"
       
       For i = 1 To lngCuotas
          lngCuotaMaxima = i

          If iPlazoRst = 1 Or iPlazoRst = vPlazo Then
            iDias = 30
          Else
            iDias = fxMesDias(Month(vFecha), Year(vFecha))
          End If
        
          curTmpInteres = curSaldo * (vInteres / 100) * iDias / 360
          curTmpAmortiza = curCuota - curTmpInteres
          
          curAmortiza = curAmortiza + curTmpAmortiza
          curInteres = curInteres + curTmpInteres
          
          curSaldo = curSaldo - curTmpAmortiza
          lngFecha = fxFechaProcesoSiguiente(lngFecha)
          vFecha = DateAdd("m", 1, vFecha)
          
          pDatos(i, 1) = curTmpInteres
          pDatos(i, 2) = curTmpAmortiza
          pDatos(i, 3) = lngFecha
          pDatos(i, 4) = curSaldo
          pDatos(i, 5) = curCuota
          
          
          iPlazoRst = iPlazoRst - 1
          curCuota = fxCalcula_Cuota(CDbl(curSaldo), iPlazoRst, vInteres)
          
       Next i
    
   
 End If 'Base

txtDatosInteres.Text = Format(curInteres, "Standard")
txtDatosAmortiza = Format(curAmortiza, "Standard")
lblFecUltMovR.Caption = lngFecha
lblCuotaR.Caption = Format(curCuota, "Standard")

If Not vRetencion Then 'El proceso nuevo de retenciones no toca los saldos
    lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - curAmortiza, "Standard")
End If


'Estado Resultante (+ Mora Registrada)
curAmortiza = curAmortiza + CCur(txtMoraAmortiza.Text)
curInteres = curInteres + CCur(txtMoraIntereses.Text)

lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + curAmortiza, "Standard")
lblInteresR.Caption = Format(CCur(lblInteres.Caption) + curInteres, "Standard")

If lngCuotas > lngCuotaMaxima Then txtCuotas = lngCuotaMaxima

End Sub

Private Sub txtCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then cboTipoDoc.SetFocus
End Sub

Private Sub txtDatosAmortiza_Change()
On Error GoTo vError

If Not vRetencion Then
    lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - (CCur(txtDatosAmortiza) + CCur(txtMoraAmortiza.Text)), "Standard")
Else
    lblSaldoR.Caption = lblCuota.Caption
End If

lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + (CCur(txtDatosAmortiza) + CCur(txtMoraAmortiza.Text)), "Standard")
lblInteresR.Caption = Format(CCur(lblInteres.Caption) + (CCur(txtDatosInteres) + CCur(txtMoraIntereses.Text)), "Standard")

txtCompromiso.Text = Format(CCur(txtDatosAmortiza) + CCur(txtMoraAmortiza.Text) + CCur(txtDatosAnticipo.Text) _
                    + CCur(txtDatosInteres.Text) + CCur(txtMoraIntereses.Text) + CCur(txtMoraCargos.Text) _
                    + CCur(txtPolizas.Text), "Standard")

vError:

End Sub

Private Sub txtDatosAmortiza_GotFocus()
On Error GoTo vError

txtDatosAmortiza = CCur(txtDatosAmortiza.Text)
vError:
End Sub

Private Sub txtDatosAmortiza_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
    txtDatosAmortiza = Format(txtDatosAmortiza.Text, "Standard")
    cboTipoDoc.SetFocus
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


Private Sub sbDocumentoAbono(pTipoAbono As String, vTipoDoc As String, vNumDoc As String _
                                , pConcepto As String, pCuenta As String)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim strCliente As String, vCuenta As String
Dim rsTmp As New ADODB.Recordset
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency, curPoliza As Currency
Dim curSaldo As Currency, pTipoCambio As Currency


vCuenta = pCuenta

pTipoCambio = fxCajasTipoCambio(ModuloCajas.mDivisa)


strSQL = "exec spCrdDocumentoAfectacionStP '" & vTipoDoc & "','" & vNumDoc & "','R'"
Call OpenRecordSet(rsTmp, strSQL, 0)
If rsTmp.EOF And rsTmp.BOF Then
  curIntC = 0
  curIntM = 0
  curAmortiza = 0
  curCargo = 0
  curPoliza = 0
Else
  curIntC = rsTmp!IntCor
  curIntM = rsTmp!IntMor
  curAmortiza = rsTmp!Principal
  curCargo = rsTmp!Cargos
  curPoliza = rsTmp!Polizas
End If
rsTmp.Close

'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
      


strLinea(1) = "Saldo Anterior    ..: " & SIFGlobal.fxStringRelleno(lblSaldo.Caption, "I", pCharRelleno, 15) '
strLinea(2) = "Saldo Actual      ..: " & SIFGlobal.fxStringRelleno(Format(CCur(lblSaldo.Caption) - curAmortiza, "Standard"), "I", pCharRelleno, 15) '
strLinea(3) = "Interes Corriente ..: " & SIFGlobal.fxStringRelleno(Format(curIntC, "Standard"), "I", pCharRelleno, 15) '
strLinea(4) = "Interes Atrasado  ..: " & SIFGlobal.fxStringRelleno(Format(curIntM, "Standard"), "I", pCharRelleno, 15) '
strLinea(5) = "Amortización      ..: " & SIFGlobal.fxStringRelleno(Format(curAmortiza, "Standard"), "I", pCharRelleno, 15) '
strLinea(6) = "Cargos Totales    ..: " & SIFGlobal.fxStringRelleno(Format(curCargo, "Standard"), "I", pCharRelleno, 15) '
strLinea(7) = "Pólizas           ..: " & SIFGlobal.fxStringRelleno(Format(curPoliza, "Standard"), "I", pCharRelleno, 15) '


strLinea(8) = "Operacion/Línea   ..: " & "Op.:" & txtOperacion.Text & " L.:" & txtCodigo & "-" & UCase(txtOpex.Text)
strLinea(9) = "Descripción       ..: " & txtDescripcion.Text
strLinea(10) = "Proc. Retencion   ..: " & IIf(vRetencion, "SI", "NO")

If dtpFechaCancelacion.Enabled Then
    strLinea(11) = "Fecha Real Abono " & Format(dtpFechaCancelacion.Value, "dd/mm/yyyy")
End If

  'Control de Documentos v2
   
        strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
                & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
                & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento,cod_caja,cod_apertura)" _
                & " values('" & vNumDoc & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
                & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo & ",'P','" & txtOperacion.Text _
                & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
                & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
                & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
                & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" & strLinea(11) & "','" _
                & txtNotas.Text & "','" & vAseDocDeposito & "','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ")"
'        Call ConectionExecute(strSQL)
        
        If curIntC > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curIntC * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintc _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          Call ConectionExecute(strSQL)
        End If
        
        If curIntM > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curIntM * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintm _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          Call ConectionExecute(strSQL)
        End If
        
        If curCargo > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curCargo * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!CtaCargos _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          Call ConectionExecute(strSQL)
        End If
        
        
        If curAmortiza > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curAmortiza * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          Call ConectionExecute(strSQL)
        End If

       If curIntC + curIntM + curPoliza + curCargo + curAmortiza > 0 Then
            'Procesa Formas de Pago (Registro Final / Asiento de Pago)
             strSQL = strSQL & Space(10) & "exec spCajas_DesglocePagosDocFinal '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ",'" & ModuloCajas.mTiquete _
                     & "','" & ModuloCajas.mUsuario & "','" & vTipoDoc & "','" & vNumDoc & "','" & ModuloCajas.mUnidad _
                     & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "'"
'             Call ConectionExecute(strSQL)
       End If
       
       'Aplicación en una sola llamada
       Call ConectionExecute(strSQL)


rs.Close


End Sub


Private Sub txtTotalCajas_GotFocus()

On Error GoTo vError
 txtCompromiso.Text = CCur(txtCompromiso.Text)
vError:

End Sub

Private Sub txtCompromiso_LostFocus()
Dim vFecha As Date, vProceso As Long
Dim curInteres As Currency, curAmortiza As Currency
Dim i As Integer
 
On Error GoTo vError
 
'ExtraOrdinario
If optAbono.Item(1).Checked Then
   'Cobra intereses del mes, pero verificar la fecha de proceso que sea igual o menor
  
   vFecha = dtpFechaCancelacion.Value
   vProceso = Year(vFecha) & Format(Month(vFecha), "00")
   
   If vProceso >= vPrideduc And vProceso > CLng(lblFecUltMov.Caption) Then
     curInteres = (CCur(txtCompromiso.Text) * vInteres / 36000) * Day(vFecha)
   Else
     curInteres = 0
   End If
   
'   '2do Paso de Validacion de Pago de Intereses
'   'Que la fecha de Primer Deduccion sea mayor al ultimo abono (No ha iniciado plan de pago)
'   If curInteres > 0 And (vPrideduc > vProceso Or vPrideduc > CLng(lblFecUltMov.Caption)) Then
'     curInteres = 0
'   End If
   
   
   'Verifica que no sea un credito del mismo mes
   If curInteres > 0 And Month(CDate(lblSaldo.Tag)) = Month(vFecha) _
        And Year(CDate(lblSaldo.Tag)) = Year(vFecha) Then
      curInteres = 0
   End If
   
   'Se re-calculan intereses para ajustar y relacionar segun porcion amortizada
   'Previamente sobre el monto a cancelar
   
   If curInteres > 0 Then
      'Hacer 10 aproximaciones
      For i = 1 To 10
            curAmortiza = CCur(txtCompromiso.Text) - curInteres
            curInteres = (curAmortiza * vInteres / 36000) * Day(vFecha)
      Next i
   End If
   
   txtDatosInteres.Text = Format(curInteres, "Standard")
   txtDatosAmortiza.Text = Format(CCur(txtCompromiso.Text) - curInteres, "Standard")

End If


txtCompromiso.Text = Format(CCur(txtCompromiso.Text), "Standard")

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


