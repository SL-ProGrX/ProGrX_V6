VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCntX_DiferidosCreacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Creación de Movimientos Diferidos"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   9135
      _Version        =   1310723
      _ExtentX        =   16113
      _ExtentY        =   7011
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
      Item(0).Caption =   "Datos"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "txtAnio"
      Item(0).Control(1)=   "txtPeriodo"
      Item(0).Control(2)=   "txtMes"
      Item(0).Control(3)=   "txtDocumento"
      Item(0).Control(4)=   "txtDetalle"
      Item(0).Control(5)=   "Label1(1)"
      Item(0).Control(6)=   "Label6(9)"
      Item(0).Control(7)=   "Label7(1)"
      Item(0).Control(8)=   "GroupBox1"
      Item(1).Caption =   "Historial"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3615
         Left            =   -70000
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1310723
         _ExtentX        =   16113
         _ExtentY        =   6376
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   8895
         _Version        =   1310723
         _ExtentX        =   15690
         _ExtentY        =   3413
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtMontoDiferir 
            Height          =   315
            Left            =   1920
            TabIndex        =   22
            Top             =   360
            Width           =   2055
            _Version        =   1310723
            _ExtentX        =   3625
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAcumulado 
            Height          =   315
            Left            =   1920
            TabIndex        =   24
            Top             =   1080
            Width           =   2055
            _Version        =   1310723
            _ExtentX        =   3625
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlazo 
            Height          =   315
            Left            =   1920
            TabIndex        =   23
            Top             =   720
            Width           =   1695
            _Version        =   1310723
            _ExtentX        =   2990
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
            Text            =   "1"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ScrollBar Scroll_Plazo 
            Height          =   255
            Left            =   3720
            TabIndex        =   26
            Top             =   730
            Width           =   255
            _Version        =   1310723
            _ExtentX        =   445
            _ExtentY        =   0
            _StockProps     =   64
            Orientation     =   1
            UseVisualStyle  =   0   'False
            Appearance      =   2
         End
         Begin XtremeSuiteControls.ComboBox cbo 
            Height          =   330
            Left            =   1920
            TabIndex        =   27
            Top             =   1440
            Width           =   2055
            _Version        =   1310723
            _ExtentX        =   3625
            _ExtentY        =   582
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtCreadoFecha 
            Height          =   315
            Left            =   6120
            TabIndex        =   25
            Top             =   360
            Width           =   2535
            _Version        =   1310723
            _ExtentX        =   4471
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtCreadoUser 
            Height          =   315
            Left            =   6120
            TabIndex        =   32
            Top             =   720
            Width           =   2535
            _Version        =   1310723
            _ExtentX        =   4471
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtProcesoFecha 
            Height          =   315
            Left            =   6120
            TabIndex        =   33
            Top             =   1080
            Width           =   2535
            _Version        =   1310723
            _ExtentX        =   4471
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtProcesoUser 
            Height          =   315
            Left            =   6120
            TabIndex        =   34
            Top             =   1440
            Width           =   2535
            _Version        =   1310723
            _ExtentX        =   4471
            _ExtentY        =   556
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
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Creación"
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
            Left            =   4320
            TabIndex        =   31
            Top             =   360
            Width           =   1965
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario Creación"
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
            Index           =   6
            Left            =   4320
            TabIndex        =   30
            Top             =   720
            Width           =   1965
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Ult.Proceso"
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
            Index           =   7
            Left            =   4320
            TabIndex        =   29
            Top             =   1080
            Width           =   1965
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario Ult. Proc."
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
            Index           =   8
            Left            =   4320
            TabIndex        =   28
            Top             =   1440
            Width           =   1965
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Monto a Diferir"
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
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1725
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Plazo"
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
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Acumulado"
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
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   1725
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Estado del Diferido"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Width           =   1725
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtAnio 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   600
         Width           =   735
         _Version        =   1310723
         _ExtentX        =   1296
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
      Begin XtremeSuiteControls.FlatEdit txtPeriodo 
         Height          =   315
         Left            =   3240
         TabIndex        =   10
         Top             =   600
         Width           =   5535
         _Version        =   1310723
         _ExtentX        =   9763
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
      Begin XtremeSuiteControls.FlatEdit txtMes 
         Height          =   315
         Left            =   2760
         TabIndex        =   11
         Top             =   600
         Width           =   495
         _Version        =   1310723
         _ExtentX        =   873
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
         Left            =   2040
         TabIndex        =   12
         Top             =   960
         Width           =   6735
         _Version        =   1310723
         _ExtentX        =   11880
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Left            =   2040
         TabIndex        =   13
         Top             =   1320
         Width           =   6735
         _Version        =   1310723
         _ExtentX        =   11880
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
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
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
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
         Index           =   9
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicia"
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
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   945
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
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
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8520
      TabIndex        =   2
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodPlantilla 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1720
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
   Begin XtremeSuiteControls.FlatEdit txtDesPlantilla 
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   6015
      _Version        =   1310723
      _ExtentX        =   10610
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1720
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
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Width           =   6015
      _Version        =   1310723
      _ExtentX        =   10610
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Plantilla"
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
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Diferido Id:"
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
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   885
   End
End
Attribute VB_Name = "frmCntX_DiferidosCreacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vBusca As Integer
Dim vCodigo As Long, vScroll As Boolean

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not IsNumeric(txtCodPlantilla.Text) Then
    txtCodPlantilla.Text = "1"
End If
       
If Not IsNumeric(txtCodigo.Text) Then
    txtCodigo.Text = "0"
End If

If vScroll Then
    strSQL = "select Top 1 cod_DifPlantilla,cod_diferido from CntX_diferido_plantilla" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_diferido = " & txtCodPlantilla.Text
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_DifPlantilla > " & txtCodigo.Text & " order by cod_DifPlantilla asc"
    Else
       strSQL = strSQL & " and cod_DifPlantilla < " & txtCodigo.Text & " order by cod_DifPlantilla desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
       txtCodigo.Text = rs!cod_DifPlantilla
       txtCodPlantilla.Text = rs!cod_diferido
       Call sbConsulta(rs!cod_DifPlantilla, rs!cod_diferido)
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

Private Sub Form_Load()


 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "No. Asiento", 1800
    .Add , , "Tipo", 1000, vbCenter
    .Add , , "Fecha", 1400, vbCenter
    .Add , , "Año", 900, vbCenter
    .Add , , "Mes", 900, vbCenter
    .Add , , "Descripción", 3500
End With

Call sbToolBarIconos(tlb)
vEdita = False
Call sbLimpiaPantalla
Call sbToolBar(tlb, "activo")
 
 
If gCntX_Arbol.ArbolActivo Then
 Call sbConsulta(Val(gCntX_Arbol.AsientoNumr), Val(gCntX_Arbol.AsientoTipo))
End If

 
Call Formularios(Me)
Call RefrescaTags(Me)
 
End Sub


Private Sub sbLimpiaPantalla()
vBusca = 1

vCodigo = 0
txtCodigo = ""
txtDescripcion = ""

txtCodPlantilla.Text = ""
txtDesPlantilla.Text = ""

txtAnio.Text = gCntX_Parametros.PeriodoAnio
txtMes.Text = gCntX_Parametros.PeriodoMes
txtPeriodo.Text = fxCntX_PeriodoDesc(txtAnio, txtMes)

txtDocumento = ""
txtDetalle = ""

txtCreadoFecha = ""
txtCreadoUser = ""
txtProcesoFecha = ""
txtProcesoUser = ""

txtMontoDiferir = 0
txtAcumulado = 0

cbo.Clear
cbo.AddItem "Activo"
cbo.AddItem "Cancelado"

cbo.Text = "Activo"

Scroll_Plazo.Value = txtPlazo.Text

tcMain.Item(0).Selected = True

End Sub





Private Sub Scroll_Plazo_Change()
txtPlazo.Text = Scroll_Plazo.Value
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If Item.Index = 1 And txtCodigo.Text <> "" Then
  
  lsw.ListItems.Clear
  
  strSQL = "select * from CntX_Diferido_Historico where cod_DifPlantilla = " _
         & txtCodigo & " and cod_diferido = " & txtCodPlantilla _
         & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
         & " order by anio,mes"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Num_Asiento)
       itmX.SubItems(1) = rs!tipo_asiento
       itmX.SubItems(2) = Format(rs!fecha, "yyyy/mm/dd")
       itmX.SubItems(3) = rs!Anio
       itmX.SubItems(4) = rs!Mes
   rs.MoveNext
  Loop
  rs.Close

End If

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      Call sbToolBar(tlb, "edicion")
    
      txtCodPlantilla.SetFocus
    
    Case "MODIFICAR", "EDITAR"
        vEdita = True
        txtDescripcion.SetFocus
        Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
        Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
      Call sbGuardar
    
    Case "DESHACER"
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
    
    Case "CONSULTAR"
       Select Case vBusca
         Case 1, 2 'Codigo del Diferido
            If vBusca = 1 Then
                gBusquedas.Columna = "cod_DifPlantilla"
                gBusquedas.Orden = "cod_DifPlantilla"
            Else
                gBusquedas.Columna = "Descripcion"
                gBusquedas.Orden = "Descripcion"
            End If
            If IsNumeric(txtCodPlantilla) Then
                gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                       & " and cod_diferido = " & txtCodPlantilla
            Else
                gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
            End If
            gBusquedas.Consulta = "select cod_DifPlantilla,cod_diferido,descripcion from CntX_diferido_plantilla"
            frmBusquedas.Show vbModal
            Call sbLimpiaPantalla
            txtCodigo = gBusquedas.Resultado
            txtCodPlantilla = gBusquedas.Resultado2
            txtCodigo.SetFocus
            
         Case 3, 4 'Codigo o Descripcion  de Plantilla
            If vBusca = 3 Then
                gBusquedas.Columna = "cod_diferido"
                gBusquedas.Orden = "cod_diferido"
            Else
                gBusquedas.Columna = "Descripcion"
                gBusquedas.Orden = "Descripcion"
            End If
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
            gBusquedas.Consulta = "select cod_diferido,descripcion,case when tipo = 'I' Then 'INGRESOS'" _
                      & " when tipo = 'G' Then 'GASTOS' end as Tipo from CntX_Diferidos"
            frmBusquedas.Show vbModal
            txtCodPlantilla = gBusquedas.Resultado
            txtDesPlantilla = gBusquedas.Resultado2
            txtCodPlantilla.SetFocus
       
       End Select

    Case "REPORTES"
      
'      strSQL = "{Cntx_Asientos.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
'             & " AND {Cntx_Asientos.TIPO_ASIENTO} = '" & txtCAsiento & "' AND " _
'             & " {Cntx_Asientos.NUM_ASIENTO} = '" & txtNAsiento & "'"
'
'      Call sbCntX_Reportes("ASIENTO", strSQL)
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
    
    Case "CERRAR"
      UnLoad Me
End Select

End Sub

Private Sub sbConsulta(vCodDiferido As Long, vPlantilla As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select D.*,P.descripcion as DescPlantilla" _
       & " from CntX_diferido_plantilla D inner join CntX_Diferidos P on D.cod_diferido = P.cod_diferido" _
       & " and D.cod_contabilidad = P.cod_contabilidad and D.tipo_asiento = P.tipo_asiento" _
       & " where D.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and D.cod_DifPlantilla = " & vCodDiferido & " and D.cod_diferido = " & vPlantilla
       
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
 
  'llenar datos en pantalla
    vCodigo = rs!cod_DifPlantilla
    txtAnio = rs!Anio
    txtMes = rs!Mes
    txtPeriodo = fxCntX_PeriodoDesc(rs!Anio, rs!Mes)
    
    txtCodigo = rs!cod_DifPlantilla
    txtDescripcion = rs!Descripcion & ""
    
    txtCodPlantilla = rs!cod_diferido
    txtDesPlantilla = rs!DescPlantilla & ""
    
    txtDetalle = rs!Detalle & ""
    txtDocumento = rs!documento & ""

    txtCreadoFecha = rs!fecha_crea & ""
    txtCreadoUser = rs!user_crea & ""
    txtProcesoFecha = rs!fecha_procesa & ""
    txtProcesoUser = rs!user_procesa & ""
    
    txtMontoDiferir = Format(rs!monto_diferir, "Standard")
    txtAcumulado = Format(rs!acumulado, "Standard")
    txtPlazo = rs!plazo
    
    Select Case rs!Estado
       Case "A"
        cbo.Text = "Activo"
       Case "C"
        cbo.Text = "Cancelado"
    End Select
    
    Scroll_Plazo.Value = rs!plazo

End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxPlantilla(vPlantilla As String, vTipo As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select consecutivo,tipo_asiento from CntX_Diferidos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_diferido = " & vPlantilla
Call OpenRecordSet(rs, strSQL, 0)
If vTipo = "C" Then
    fxPlantilla = rs!Consecutivo + 1
    strSQL = "update CntX_Diferidos set consecutivo = consecutivo + 1 where cod_contabilidad = " _
           & gCntX_Parametros.CodigoConta & " and cod_diferido = " & vPlantilla
    Call ConectionExecute(strSQL, 0)
Else
 'Tipo Asiento
    fxPlantilla = rs!tipo_asiento
End If
rs.Close

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset, lng As Long

On Error GoTo vError

If vEdita Then

     strSQL = "update CntX_diferido_plantilla set descripcion = '" & UCase(txtDescripcion) _
            & "',detalle = '" & txtDetalle _
            & "',documento = '" & txtDocumento & "',estado = '" & Mid(cbo.Text, 1, 1) & "'"
     If CCur(txtAcumulado) = 0 Then
       strSQL = strSQL & ",anio =" & txtAnio & ",mes = " & txtMes _
              & ",monto_diferir = " & CCur(txtMontoDiferir) _
              & ",acumulado = 0" _
              & ",plazo = " & txtPlazo _
              & ",cod_diferido = " & txtCodPlantilla _
              & ",tipo_asiento = '" & fxPlantilla(txtCodPlantilla, "T") & "'"
     End If
     strSQL = strSQL & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
            & " and cod_DifPlantilla = " & vCodigo _
            & " and cod_diferido = " & txtCodPlantilla
     Call ConectionExecute(strSQL, 0)

     Call Bitacora("Modifica", "DIFERIDO NUM: " & vCodigo & " Emp" & gCntX_Parametros.CodigoConta)


Else 'Inserta


      vCodigo = fxPlantilla(txtCodPlantilla, "C")
      
      strSQL = "insert CntX_diferido_plantilla(cod_diferido,cod_contabilidad,tipo_asiento,cod_DifPlantilla,Anio,Mes" _
             & ",fecha_crea,user_crea,monto_diferir,plazo,acumulado,consecutivo,detalle,documento,estado,descripcion)" _
             & " values(" & txtCodPlantilla & "," & gCntX_Parametros.CodigoConta & ",'" & fxPlantilla(txtCodPlantilla, "T") _
             & "'," & vCodigo & "," & txtAnio & "," & txtMes & ",getdate(),'" & glogon.Usuario & "'," _
             & CCur(txtMontoDiferir) & "," & txtPlazo & ",0,0,'" _
             & txtDetalle & "','" & txtDocumento & "','A','" & Trim(UCase(txtDescripcion)) & "')"
      Call ConectionExecute(strSQL, 0)
      
      txtCodigo = vCodigo
      
      Call Bitacora("Registra", "DIFERIDO NUM : " & txtCodigo & " Emp" & gCntX_Parametros.CodigoConta)

End If 'Si Inserta o Actualiza

Call sbToolBar(tlb, "activo")
Call sbConsulta(txtCodigo, txtCodPlantilla)

vEdita = True

MsgBox "Información guardada satisfactoriamente...", vbInformation


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
  strSQL = "delete CntX_Plantilla_detalle where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and cod_plantilla = " & txtCodPlantilla
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "delete CntX_Plantilla_Asientos where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and cod_plantilla = " & txtCodPlantilla
  Call ConectionExecute(strSQL, 0)
  

  Call Bitacora("Elimina", "Plantilla Asiento : " & txtCodPlantilla & " EMP:" _
                  & gCntX_Parametros.CodigoConta)

  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtAnio_Change()
On Error GoTo vError
  txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 Call sbConsulta(txtCodigo, txtCodPlantilla)
 txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))

Exit Sub

vError:
  Call sbLimpiaPantalla
End Sub

Private Sub txtCodPlantilla_GotFocus()
vBusca = 3
End Sub

Private Sub txtCodigo_GotFocus()
vBusca = 1
End Sub

Private Sub txtDescripcion_GotFocus()
vBusca = 2
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodPlantilla.SetFocus
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub

Private Sub txtDesPlantilla_GotFocus()
vBusca = 4
End Sub

Private Sub txtDesPlantilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMes.SetFocus
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub

Private Sub txtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMontoDiferir.SetFocus
End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub txtMes_Change()
On Error GoTo vError
  txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtCodPlantilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDesPlantilla.SetFocus
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAnio.SetFocus
End Sub

Private Sub txtMontoDiferir_GotFocus()
On Error GoTo vError
txtMontoDiferir = CCur(txtMontoDiferir)
vError:
End Sub

Private Sub txtMontoDiferir_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub txtMontoDiferir_LostFocus()
On Error GoTo vError
txtMontoDiferir = Format(CCur(txtMontoDiferir), "Standard")
vError:
End Sub


