VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmActivos_Revaluaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Revaluaciones de Activos"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4212
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   10092
      _Version        =   1572864
      _ExtentX        =   17801
      _ExtentY        =   7429
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
      Item(0).Caption =   "Registro"
      Item(0).ControlCount=   25
      Item(0).Control(0)=   "txtDescripcion"
      Item(0).Control(1)=   "Label6(0)"
      Item(0).Control(2)=   "Label6(1)"
      Item(0).Control(3)=   "cbo"
      Item(0).Control(4)=   "dtpFecha"
      Item(0).Control(5)=   "Label6(2)"
      Item(0).Control(6)=   "Label6(3)"
      Item(0).Control(7)=   "Label6(4)"
      Item(0).Control(8)=   "Label6(5)"
      Item(0).Control(9)=   "Label6(6)"
      Item(0).Control(10)=   "cboVidaUtil"
      Item(0).Control(11)=   "txtMeses"
      Item(0).Control(12)=   "txtMonto"
      Item(0).Control(13)=   "txtRevaluacion"
      Item(0).Control(14)=   "Label6(7)"
      Item(0).Control(15)=   "Label6(8)"
      Item(0).Control(16)=   "Label6(9)"
      Item(0).Control(17)=   "Label6(10)"
      Item(0).Control(18)=   "Label6(11)"
      Item(0).Control(19)=   "lblPeriodo"
      Item(0).Control(20)=   "lblHistorico"
      Item(0).Control(21)=   "lblRescate"
      Item(0).Control(22)=   "lblDepreciacion"
      Item(0).Control(23)=   "lblLibros"
      Item(0).Control(24)=   "lblID"
      Item(1).Caption =   "Histórico"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "lbl"
      Item(1).Control(1)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3492
         Left            =   -70000
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1572864
         _ExtentX        =   17801
         _ExtentY        =   6159
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
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   912
         Left            =   2520
         TabIndex        =   7
         Top             =   960
         Width           =   6252
         _Version        =   1572864
         _ExtentX        =   11028
         _ExtentY        =   1609
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
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   2520
         TabIndex        =   11
         Top             =   480
         Width           =   6252
         _Version        =   1572864
         _ExtentX        =   11033
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
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   312
         Left            =   2520
         TabIndex        =   12
         Top             =   2280
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.ComboBox cboVidaUtil 
         Height          =   312
         Left            =   2520
         TabIndex        =   18
         Top             =   2640
         Width           =   1572
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtMeses 
         Height          =   315
         Left            =   2520
         TabIndex        =   19
         Top             =   3000
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   2520
         TabIndex        =   20
         Top             =   3360
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRevaluacion 
         Height          =   312
         Left            =   2520
         TabIndex        =   21
         Top             =   3720
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Text            =   "0"
         BackColor       =   16777152
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label lblID 
         Height          =   252
         Left            =   8520
         TabIndex        =   8
         Top             =   0
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Id.:"
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
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLibros 
         Height          =   312
         Left            =   6960
         TabIndex        =   31
         Top             =   3720
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "0"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label lblDepreciacion 
         Height          =   312
         Left            =   6960
         TabIndex        =   30
         Top             =   3360
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "0"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label lblRescate 
         Height          =   312
         Left            =   6960
         TabIndex        =   29
         Top             =   3000
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "0"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label lblHistorico 
         Height          =   312
         Left            =   6960
         TabIndex        =   28
         Top             =   2640
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "0"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label lblPeriodo 
         Height          =   312
         Left            =   6960
         TabIndex        =   27
         Top             =   2280
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "0"
         BackColor       =   14737632
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
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   11
         Left            =   5040
         TabIndex        =   26
         Top             =   3720
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Valor en Libros"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   10
         Left            =   5040
         TabIndex        =   25
         Top             =   3360
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Depreciación Acu."
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   9
         Left            =   5040
         TabIndex        =   24
         Top             =   3000
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Total Rescate"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   8
         Left            =   5040
         TabIndex        =   23
         Top             =   2640
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Total Historico"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   7
         Left            =   5040
         TabIndex        =   22
         Top             =   2280
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ultimo Periodo"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   6
         Left            =   960
         TabIndex        =   17
         Top             =   3720
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Revaluación"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   5
         Left            =   960
         TabIndex        =   16
         Top             =   3360
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Monto"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   4
         Left            =   960
         TabIndex        =   15
         Top             =   3000
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Meses V.U."
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   3
         Left            =   960
         TabIndex        =   14
         Top             =   2640
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Vida Util"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   2
         Left            =   960
         TabIndex        =   13
         Top             =   2280
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fecha"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   960
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   480
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Motivo"
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
      End
      Begin XtremeShortcutBar.ShortcutCaption lbl 
         Height          =   372
         Left            =   -70000
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1572864
         _ExtentX        =   17801
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Revaluaciones Registradas al Activo"
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
         VisualTheme     =   3
         Alignment       =   1
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5565
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario que Registro Activo"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   3422
            MinWidth        =   3422
            Object.ToolTipText     =   "Fecha de Registro Real"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   1482
            MinWidth        =   1482
            Object.ToolTipText     =   "Ultimo Periodo Depreciado"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   2541
            MinWidth        =   2541
            Object.ToolTipText     =   "Depreciación Acumulada"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   2187
            MinWidth        =   2187
            Object.ToolTipText     =   "Depreciación del Mes"
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   432
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   2412
      _Version        =   1572864
      _ExtentX        =   4254
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   1440
      TabIndex        =   32
      ToolTipText     =   "Nuevo"
      Top             =   45
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Revaluaciones.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   2520
      TabIndex        =   33
      ToolTipText     =   "Editar"
      Top             =   45
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Revaluaciones.frx":0632
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   2880
      TabIndex        =   34
      ToolTipText     =   "Eliminar"
      Top             =   45
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Revaluaciones.frx":0C2D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   3480
      TabIndex        =   35
      ToolTipText     =   "Guardar"
      Top             =   45
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Revaluaciones.frx":11D1
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   3840
      TabIndex        =   36
      ToolTipText     =   "Deshacer"
      Top             =   45
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Revaluaciones.frx":1902
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   4320
      TabIndex        =   37
      ToolTipText     =   "Reporte"
      Top             =   45
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Revaluaciones.frx":2002
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.Label LabelX 
      Height          =   192
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   339
      _StockProps     =   79
      Caption         =   "No. Placa"
      ForeColor       =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblActivo 
      Height          =   432
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   6012
      _Version        =   1572864
      _ExtentX        =   10604
      _ExtentY        =   762
      _StockProps     =   79
      Caption         =   "xx"
      ForeColor       =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmActivos_Revaluaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vPaso As Boolean

Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub btnBarra_Click(Index As Integer)
Dim strSQL As String


Select Case Index
    Case 0 'NUEVO
        vEdita = False
        Call sbLimpiaPantalla
        txtCodigo.SetFocus

        Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
      If lblID.Tag = "" Then
        MsgBox "Seleccione una Mejora o Retiro de la lista del activo para modificacion...", vbInformation
      Else
        vEdita = True
        txtCodigo.SetFocus
        Call sbBarra_Accion("Editar")
      End If
      
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
      Call sbBarra_Accion("Editar")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      End If
    
    Case 5 'REPORTES
   
End Select

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub cboVidaUtil_Click()
txtMeses = fxMeses
End Sub

Private Sub cboVidaUtil_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboVidaUtil.SetFocus
End Sub

Public Sub sbArbolShow()
  Call sbConsulta(txtCodigo)
End Sub


Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 36
 
 vEdita = True

cboVidaUtil.Clear
cboVidaUtil.AddItem "Restante del Activo"
cboVidaUtil.AddItem "Suplementaria del Activo"

 With lsw.ColumnHeaders
        .Add , , "Boleta Id", 1200
        .Add , , "Tipo Mov", 1600, vbCenter
        .Add , , "Fecha", 1800, vbCenter
        .Add , , "Monto", 1600, vbRightJustify
        .Add , , "Motivo", 2600
        .Add , , "Descripción", 2600
 End With
 
 Call sbLimpiaPantalla
 Call sbBarra_Accion("Nuevo")

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim strSQL As String

tcMain.Item(0).Selected = True

vCodigo = ""
txtCodigo = ""

strSQL = "select rtrim(cod_justificacion) as 'IdX',rtrim(descripcion) as 'ItmX'" _
       & " from Activos_justificaciones where tipo = 'V'"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

lblActivo.Caption = ""
lblID.Tag = ""
lblID.Visible = False

txtDescripcion = ""
dtpFecha.Value = gActivos.Periodo


txtMonto.Text = "0"
txtRevaluacion.Text = "0"

cboVidaUtil.Text = "Restante del Activo"

StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""
StatusBarX.Panels(3).Text = 0
StatusBarX.Panels(4).Text = 0
StatusBarX.Panels(5).Text = 0


End Sub



Private Sub sbDatosActivo()
Dim strSQL As String, rs As New ADODB.Recordset

If Len(txtCodigo) = 0 Then Exit Sub


strSQL = "exec spActivos_InfoDepreciacion '" & txtCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  lblPeriodo.Caption = Format(rs!depreciacion_periodo & "", "dd/mm/yyyy")
  lblDepreciacion.Caption = Format(rs!depreciacion_acum, "Standard")
  lblHistorico.Caption = Format(rs!valor_historico, "Standard")
  lblRescate.Caption = Format(rs!valor_desecho, "Standard")
  lblLibros.Caption = Format(rs!VALOR_LIBROS, "Standard")
Else
  lblPeriodo.Caption = "????"
  lblDepreciacion.Caption = 0
  lblHistorico.Caption = 0
  lblRescate.Caption = 0
  lblLibros.Caption = 0
End If
rs.Close

 
txtMeses.Text = CStr(fxMeses)
  

End Sub





Public Sub sbConsultaExterna(pNumPlaca As String)
If pNumPlaca <> "" Then
 Call sbConsulta(pNumPlaca)
End If
End Sub


Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select X.*,rtrim(J.cod_justificacion) as 'Motivo_Id', rtrim(J.descripcion) as 'Motivo_Desc'" _
       & ",A.nombre,P.cod_proveedor,P.descripcion as Proveedor" _
       & " from Activos_retiro_adicion X inner join Activos_Principal A on X.num_placa = A.num_placa" _
       & " inner join Activos_justificaciones J on X.cod_justificacion = J.cod_justificacion" _
       & " left join Activos_proveedores P on X.compra_proveedor = P.cod_proveedor" _
       & " where X.id_AddRet = " & lblID.Tag & " and X.num_placa = '" & txtCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbBarra_Accion("Activo")
  vEdita = True
  
  tcMain.Item(0).Selected = True
  
  vCodigo = rs!num_placa
  txtCodigo = rs!num_placa
  
  lblActivo.Caption = rs!Nombre
  
  txtDescripcion.Text = rs!Descripcion
  dtpFecha.Value = rs!fecha
  txtMonto.Text = Format(rs!monto, "Standard")
  txtRevaluacion.Text = Format(rs!monto + rs!VALOR_LIBROS, "Standard")
    
  If rs!tipo_vidautil = "R" Then
    cboVidaUtil.Text = "Restante del Activo"
  Else
    cboVidaUtil.Text = "Suplementaria del Activo"
  End If
    
  Call sbCboAsignaDato(cbo, rs!Motivo_Desc, True, rs!Motivo_ID)
  
  
  lblID.Caption = "Revaluación Id: " & rs!id_AddRet
  lblID.Visible = True
  
  StatusBarX.Panels(1).Text = rs!creacion_user & ""
  StatusBarX.Panels(2).Text = rs!creacion_fecha & ""
  StatusBarX.Panels(3).Text = rs!depreciacion_periodo
  StatusBarX.Panels(4).Text = Format(rs!depreciacion_acum, "Standard")
  StatusBarX.Panels(5).Text = Format(rs!DEPRECIACION_MES, "Standard")
  
  Call sbDatosActivo
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""
fxValida = True

'1. Verificar Periodo / Si esta cerrado no puede registrarse
'2. No se puede modificar si ya se le ha calculo depreciacion
'3. Verifica que la fecha de la adicion o retiro sea mayor a la fecha de adquisicion
'4. del activo
'5. No puede Modificar un Activo Retirado


strSQL = "select fecha_adquisicion from Activos_Principal where num_placa = '" _
       & txtCodigo & "' and estado <> 'R'"
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
  vMensaje = vMensaje & vbCrLf & " - El Activo no existe, o ya fue retirado ..."
Else
  If DateDiff("d", rs!fecha_adquisicion, dtpFecha.Value) < 1 Then
      vMensaje = vMensaje & vbCrLf & " - La fecha del Movimiento no es válida, ya que es menor a la del activo ..."
  End If
End If
rs.Close

strSQL = "select estado, dbo.fxActivos_PeriodoActual() as 'PeriodoActual'" _
       & " from Activos_periodos where anio = " & Year(dtpFecha.Value) _
       & " and mes = " & Month(dtpFecha.Value)
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
 If rs!Estado <> "P" Then
      vMensaje = vMensaje & vbCrLf & " - El Periodo del Movimiento ya fue cerrado ..."
 End If
 
 If Year(dtpFecha.Value) <> Year(rs!PeriodoActual) Or Month(dtpFecha.Value) <> Month(rs!PeriodoActual) Then
      vMensaje = vMensaje & vbCrLf & " - La fecha de aplicación del movimiento no corresponde al periodo abierto!"
 End If

End If
rs.Close

If CCur(StatusBarX.Panels(3).Text) > 0 Then vMensaje = vMensaje & vbCrLf & " - No se puede registrar porque ya inicio ciclo de depreciacion..."
If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion del tipo de Movimiento no es válido ..."
If Not IsNumeric(txtMonto) Then vMensaje = vMensaje & vbCrLf & " - El Monto del movimiento no es válido ..."
If Not IsNumeric(txtRevaluacion) Then vMensaje = vMensaje & vbCrLf & " - El Monto de la Revaluacion no es válido ..."
If Not IsNumeric(lblLibros.Caption) Then vMensaje = vMensaje & vbCrLf & " - El valor en libros no es válido ..."
If cbo.ListCount <= 0 Then vMensaje = vMensaje & vbCrLf & " - No existe ninguna Justificación ..."

If Len(vMensaje) = 0 Then
 If CCur(lblLibros.Caption) > CCur(txtRevaluacion) Then vMensaje = vMensaje & vbCrLf & " - El valor en libros puede ser mayor que la revaluacion ..."
End If


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Function fxMeses() As Integer
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select dbo.fxActivos_VidaUtilPendiente('" & txtCodigo & "','" & Mid(cboVidaUtil.Text, 1, 1) _
        & "','" & Format(dtpFecha.Value, "yyyy/mm/dd") & "') as 'VidaUtil'"
Call OpenRecordSet(rs, strSQL, 0)
    fxMeses = rs!VidaUtil
rs.Close

Exit Function

vError:
 fxMeses = 1

End Function





Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vDepInicial As Currency, dbFactor As Double
Dim vSuperAvit As Currency

On Error GoTo vError

Call sbDatosActivo
'

'   dbFactor = CCur(txtMonto) / CCur(lblLibros.Caption)
'   vDepInicial = CCur(lblDepreciacion.Caption) * dbFactor
'   vSuperAvit = CCur(txtMonto) - vDepInicial

vCodigo = txtCodigo.Text

strSQL = "exec spActivos_AdicionRetiro '" & txtCodigo.Text & "','V','" & cbo.ItemData(cbo.ListIndex) & "','" & txtDescripcion.Text _
        & "','" & Format(dtpFecha.Value, "yyyy/mm/dd") & "'," & CCur(txtMonto.Text) & "," & CLng(txtMeses.Text) & ",'" & glogon.Usuario _
        & "','', '', '', '' "
Call OpenRecordSet(rs, strSQL, 0)

lblID.Tag = rs!Linea
 
rs.Close

strSQL = "Revaluación (Placa: " & vCodigo & ") Id.: " & lblID.Tag & "_" & cbo.Text
Call Bitacora("Registra", strSQL)
  
MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(vCodigo)

Call sbBarra_Accion("Activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

If lblID.Tag = "" Then Exit Sub

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete Activos_retiro_adicion where num_placa = '" & vCodigo _
        & "' and id_AddRet = " & lblID.Tag
  If CCur(StatusBarX.Panels(3).Text) = 0 Then
    Call ConectionExecute(strSQL)
    Call Bitacora("Elimina", "Revaluación, Placa: " & vCodigo & ", Id: " & lblID.Tag)
  End If
  
  Call sbLimpiaPantalla
  Call sbBarra_Accion("Nuevo")
  
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
 lblID.Tag = Item.Text
 Call sbConsulta(txtCodigo)
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Select Case Item.Index
 Case 1
    If txtCodigo.Text = "" Then
      MsgBox "Ingrese Un número de placa primero...", vbInformation
      tcMain.Item(0).Selected = True
      Exit Sub
    End If
    
    
    strSQL = "select X.*,rtrim(J.cod_justificacion) + '..' + J.descripcion as Justifica" _
          & ",A.nombre,P.cod_proveedor,P.descripcion as Proveedor" _
          & " from Activos_retiro_adicion X inner join Activos_Principal A on X.num_placa = A.num_placa" _
          & " inner join Activos_justificaciones J on X.cod_justificacion = J.cod_justificacion" _
          & " left join Activos_proveedores P on X.compra_proveedor = P.cod_proveedor" _
          & " where X.num_placa = '" & txtCodigo.Text & "' and X.tipo = 'V'"
    Call OpenRecordSet(rs, strSQL, 0)
    lsw.ListItems.Clear
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!id_AddRet)
          itmX.SubItems(1) = "Revaluación"
          itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
          itmX.SubItems(3) = Format(rs!monto, "Standard")
          itmX.SubItems(4) = rs!Justifica
          itmX.SubItems(5) = rs!Descripcion
      rs.MoveNext
    Loop
    rs.Close


End Select

vError:


End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Num_Placa"
  gBusquedas.Orden = "Num_Placa"
  
  gBusquedas.Col1Name = "Id Placa"
  gBusquedas.Col2Name = "Id Alterna"
  gBusquedas.Col3Name = "Nombre"
  
  gBusquedas.Consulta = "select num_placa, Placa_Alterna, Nombre from Activos_Principal"
  
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  If txtCodigo.Text <> "" Then
    lblActivo.Caption = gBusquedas.Resultado3
    Call sbDatosActivo
  End If
End If

End Sub

Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select nombre from Activos_Principal where num_placa = '" & txtCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
 lblActivo.Caption = rs!Nombre
 Call sbDatosActivo
End If
rs.Close
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFecha.SetFocus
End Sub


Private Sub txtRevaluacion_Change()
On Error GoTo vError
 txtMonto.Text = Format(CCur(txtRevaluacion.Text) - CCur(lblLibros.Caption), "Standard")
vError:
End Sub

Private Sub txtRevaluacion_GotFocus()
On Error GoTo vError
 txtRevaluacion.Text = CStr(CCur(txtRevaluacion.Text))
vError:
End Sub

Private Sub txtRevaluacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboVidaUtil.SetFocus
End Sub

Private Sub txtRevaluacion_LostFocus()
On Error GoTo vError
 txtRevaluacion.Text = Format(CCur(txtRevaluacion.Text), "Standard")
vError:
End Sub

