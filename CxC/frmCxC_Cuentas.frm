VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCxC_Cuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Cuentas por Cobrar"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11070
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   11070
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8760
      Top             =   600
   End
   Begin XtremeSuiteControls.TabControl ssTab 
      Height          =   7092
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   10812
      _Version        =   1572864
      _ExtentX        =   19071
      _ExtentY        =   12509
      _StockProps     =   68
      Appearance      =   4
      Color           =   32
      ItemCount       =   4
      SelectedItem    =   1
      Item(0).Caption =   "Recepción"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "fraOperacion"
      Item(0).Control(1)=   "GroupBox1"
      Item(1).Caption =   "Facturas"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "gbFacturaRegistra"
      Item(1).Control(1)=   "tcFacturas"
      Item(2).Caption =   "Activación"
      Item(2).ControlCount=   10
      Item(2).Control(0)=   "txtTasaMora"
      Item(2).Control(1)=   "chkTesoreria"
      Item(2).Control(2)=   "lswOpciones"
      Item(2).Control(3)=   "lsw"
      Item(2).Control(4)=   "Label1(8)"
      Item(2).Control(5)=   "cboOficina"
      Item(2).Control(6)=   "txtDocNum"
      Item(2).Control(7)=   "Label1(2)"
      Item(2).Control(8)=   "Label1(18)"
      Item(2).Control(9)=   "TituloOpcionesSub"
      Item(3).Caption =   "Bitácora"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "GroupBox2(0)"
      Item(3).Control(1)=   "GroupBox2(1)"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3612
         Left            =   -66520
         TabIndex        =   95
         Top             =   720
         Visible         =   0   'False
         Width           =   7212
         _Version        =   1572864
         _ExtentX        =   12721
         _ExtentY        =   6371
         _StockProps     =   77
         View            =   3
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcFacturas 
         Height          =   4692
         Left            =   0
         TabIndex        =   97
         Top             =   2400
         Width           =   10812
         _Version        =   1572864
         _ExtentX        =   19071
         _ExtentY        =   8276
         _StockProps     =   68
         Appearance      =   4
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Registradas"
         Item(0).ControlCount=   10
         Item(0).Control(0)=   "lswFacturas"
         Item(0).Control(1)=   "btnArchivoBusca"
         Item(0).Control(2)=   "btnArchivoCarga"
         Item(0).Control(3)=   "feArchivo"
         Item(0).Control(4)=   "Label2(7)"
         Item(0).Control(5)=   "txtFacturas_Casos"
         Item(0).Control(6)=   "txtFacturas_Total"
         Item(0).Control(7)=   "Label2(15)"
         Item(0).Control(8)=   "txtFacturaFiltro(0)"
         Item(0).Control(9)=   "Label2(18)"
         Item(1).Caption =   "Adelantadas/Pendientes"
         Item(1).ControlCount=   10
         Item(1).Control(0)=   "lswAdelantadas"
         Item(1).Control(1)=   "bntFacturas_Vincular"
         Item(1).Control(2)=   "txtFacturas_Adelanto_Casos"
         Item(1).Control(3)=   "txtFacturas_Adelanto_Total"
         Item(1).Control(4)=   "Label2(16)"
         Item(1).Control(5)=   "txtFacturas_Adelanto_Casos_Sel"
         Item(1).Control(6)=   "txtFacturas_Adelanto_Total_Sel"
         Item(1).Control(7)=   "Label2(17)"
         Item(1).Control(8)=   "txtFacturaFiltro(1)"
         Item(1).Control(9)=   "Label2(19)"
         Begin XtremeSuiteControls.ListView lswAdelantadas 
            Height          =   3132
            Left            =   -70000
            TabIndex        =   103
            Top             =   360
            Visible         =   0   'False
            Width           =   10812
            _Version        =   1572864
            _ExtentX        =   19071
            _ExtentY        =   5524
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
            GridLines       =   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.ListView lswFacturas 
            Height          =   3012
            Left            =   0
            TabIndex        =   98
            Top             =   360
            Width           =   10812
            _Version        =   1572864
            _ExtentX        =   19071
            _ExtentY        =   5313
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnArchivoBusca 
            Height          =   312
            Left            =   8280
            TabIndex        =   99
            ToolTipText     =   "Detalle de Facturas"
            Top             =   4080
            Width           =   372
            _Version        =   1572864
            _ExtentX        =   656
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
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
         Begin XtremeSuiteControls.PushButton btnArchivoCarga 
            Height          =   432
            Left            =   9120
            TabIndex        =   100
            ToolTipText     =   "Detalle de Facturas"
            Top             =   4080
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   762
            _StockProps     =   79
            Caption         =   "Cargar Archivo"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.PushButton bntFacturas_Vincular 
            Height          =   312
            Left            =   -61960
            TabIndex        =   104
            ToolTipText     =   "Detalle de Facturas"
            Top             =   4320
            Visible         =   0   'False
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Vincular"
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtFacturas_Casos 
            Height          =   312
            Left            =   9960
            TabIndex        =   109
            Top             =   3480
            Width           =   852
            _Version        =   1572864
            _ExtentX        =   1503
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFacturas_Total 
            Height          =   312
            Left            =   8040
            TabIndex        =   110
            Top             =   3480
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
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
         Begin XtremeSuiteControls.FlatEdit txtFacturas_Adelanto_Casos 
            Height          =   312
            Left            =   -63040
            TabIndex        =   111
            Top             =   3960
            Visible         =   0   'False
            Width           =   732
            _Version        =   1572864
            _ExtentX        =   1291
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFacturas_Adelanto_Total 
            Height          =   312
            Left            =   -64720
            TabIndex        =   112
            Top             =   3960
            Visible         =   0   'False
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFacturas_Adelanto_Casos_Sel 
            Height          =   312
            Left            =   -63040
            TabIndex        =   115
            Top             =   4320
            Visible         =   0   'False
            Width           =   732
            _Version        =   1572864
            _ExtentX        =   1291
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFacturas_Adelanto_Total_Sel 
            Height          =   312
            Left            =   -64720
            TabIndex        =   116
            Top             =   4320
            Visible         =   0   'False
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFacturaFiltro 
            Height          =   312
            Index           =   0
            Left            =   1320
            TabIndex        =   118
            Top             =   3480
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
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
         Begin XtremeSuiteControls.FlatEdit txtFacturaFiltro 
            Height          =   312
            Index           =   1
            Left            =   -64720
            TabIndex        =   120
            Top             =   3600
            Visible         =   0   'False
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feArchivo 
            Height          =   432
            Left            =   1320
            TabIndex        =   101
            Top             =   4080
            Width           =   6852
            _Version        =   1572864
            _ExtentX        =   12086
            _ExtentY        =   762
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
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. Factura ?"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   19
            Left            =   -65920
            TabIndex        =   121
            Top             =   3600
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. Factura ?"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   18
            Left            =   120
            TabIndex        =   119
            Top             =   3480
            Width           =   1092
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Seleccionadas:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   17
            Left            =   -66400
            TabIndex        =   117
            Top             =   4320
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Totales:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   16
            Left            =   -66400
            TabIndex        =   114
            Top             =   3960
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Totales:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   15
            Left            =   6840
            TabIndex        =   113
            Top             =   3480
            Width           =   1092
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Archivo:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   7
            Left            =   120
            TabIndex        =   102
            Top             =   4080
            Width           =   1092
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   3972
         Index           =   0
         Left            =   -68680
         TabIndex        =   55
         Top             =   2640
         Visible         =   0   'False
         Width           =   7932
         _Version        =   1572864
         _ExtentX        =   13991
         _ExtentY        =   7006
         _StockProps     =   79
         Caption         =   "Tramite"
         ForeColor       =   4210752
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtAutorizaNota 
            Height          =   852
            Left            =   2760
            TabIndex        =   145
            Top             =   2640
            Width           =   5172
            _Version        =   1572864
            _ExtentX        =   9123
            _ExtentY        =   1503
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Resolución"
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   3
            Left            =   1320
            TabIndex        =   73
            Top             =   3600
            Width           =   1212
         End
         Begin VB.Label lblAutorizaEstado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   2760
            TabIndex        =   72
            Top             =   3600
            Width           =   5172
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nota"
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   7
            Left            =   1320
            TabIndex        =   71
            Top             =   2640
            Width           =   1212
         End
         Begin VB.Label lblFechaAuto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   5400
            TabIndex        =   70
            Top             =   2280
            Width           =   2532
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Autorizada"
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   6
            Left            =   1320
            TabIndex        =   69
            Top             =   2280
            Width           =   1212
         End
         Begin VB.Label lblAutorizada 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   2760
            TabIndex        =   68
            Top             =   2280
            Width           =   2532
         End
         Begin VB.Label lblTesoreria 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   2760
            TabIndex        =   67
            Top             =   1440
            Width           =   2532
         End
         Begin VB.Label lblFormaliza 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   2760
            TabIndex        =   66
            Top             =   1080
            Width           =   2532
         End
         Begin VB.Label lblResoluciona 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   2760
            TabIndex        =   65
            Top             =   360
            Width           =   2532
         End
         Begin VB.Label lblRecibe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   2760
            TabIndex        =   64
            Top             =   720
            Width           =   2532
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Bancos"
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   5
            Left            =   1320
            TabIndex        =   63
            Top             =   1440
            Width           =   1212
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Activa"
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   4
            Left            =   1320
            TabIndex        =   62
            Top             =   1080
            Width           =   1212
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Recibe"
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   2
            Left            =   1320
            TabIndex        =   61
            Top             =   720
            Width           =   1212
         End
         Begin VB.Label lblFechaTes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   5400
            TabIndex        =   60
            Top             =   1440
            Width           =   2532
         End
         Begin VB.Label lblFechaFor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   5400
            TabIndex        =   59
            Top             =   1080
            Width           =   2532
         End
         Begin VB.Label lblFechaRes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   5400
            TabIndex        =   58
            Top             =   360
            Width           =   2532
         End
         Begin VB.Label lblFechaRec 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Left            =   5400
            TabIndex        =   57
            Top             =   720
            Width           =   2532
         End
      End
      Begin XtremeSuiteControls.GroupBox gbFacturaRegistra 
         Height          =   1695
         Left            =   0
         TabIndex        =   31
         Top             =   480
         Width           =   10695
         _Version        =   1572864
         _ExtentX        =   18865
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "Registro de Facturas:"
         ForeColor       =   4210752
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit feFactura 
            Height          =   330
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   2292
            _Version        =   1572864
            _ExtentX        =   4043
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
         Begin XtremeSuiteControls.FlatEdit feFacturaImporte 
            Height          =   330
            Left            =   4320
            TabIndex        =   37
            Top             =   600
            Width           =   1815
            _Version        =   1572864
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cbofacturasDivisas 
            Height          =   312
            Left            =   2400
            TabIndex        =   38
            Top             =   600
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboFacturaEstado 
            Height          =   312
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   2292
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feFacturaMonto 
            Height          =   330
            Left            =   6120
            TabIndex        =   44
            Top             =   600
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFacturaEmite 
            Height          =   330
            Left            =   8040
            TabIndex        =   45
            Top             =   600
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.DateTimePicker dtpFacturaPago 
            Height          =   330
            Left            =   9360
            TabIndex        =   46
            Top             =   600
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.PushButton btnFacturaRegistra 
            Height          =   315
            Left            =   8160
            TabIndex        =   47
            ToolTipText     =   "Detalle de Facturas"
            Top             =   1200
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Registra"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.PushButton btnFacturaElimina 
            Height          =   315
            Left            =   9360
            TabIndex        =   48
            ToolTipText     =   "Detalle de Facturas"
            Top             =   1200
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2138
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Elimina"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.FlatEdit feTipoCambio 
            Height          =   330
            Left            =   3480
            TabIndex        =   50
            Top             =   600
            Width           =   852
            _Version        =   1572864
            _ExtentX        =   1503
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboAdelantoTipo 
            Height          =   330
            Left            =   4320
            TabIndex        =   82
            Top             =   1200
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feAdelanto 
            Height          =   330
            Left            =   6120
            TabIndex        =   83
            Top             =   1200
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkAdelanta 
            Height          =   252
            Left            =   2880
            TabIndex        =   39
            Top             =   1200
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Adelanta"
            BackColor       =   -2147483633
            UseVisualStyle  =   -1  'True
            Alignment       =   1
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo"
            Height          =   252
            Index           =   14
            Left            =   4320
            TabIndex        =   85
            Top             =   960
            Width           =   852
         End
         Begin VB.Label Label2 
            Caption         =   "Adelanto"
            Height          =   255
            Index           =   13
            Left            =   6120
            TabIndex        =   84
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "T.C."
            Height          =   252
            Index           =   8
            Left            =   3480
            TabIndex        =   51
            Top             =   360
            Width           =   492
         End
         Begin VB.Label Label2 
            Caption         =   "Monto"
            Height          =   255
            Index           =   6
            Left            =   6120
            TabIndex        =   43
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Divisa"
            Height          =   252
            Index           =   5
            Left            =   2400
            TabIndex        =   41
            Top             =   360
            Width           =   732
         End
         Begin VB.Label Label2 
            Caption         =   "Estado"
            Height          =   252
            Index           =   4
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   852
         End
         Begin VB.Label Label2 
            Caption         =   "Pago"
            Height          =   255
            Index           =   3
            Left            =   9480
            TabIndex        =   35
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Emite"
            Height          =   255
            Index           =   2
            Left            =   8400
            TabIndex        =   34
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Importe"
            Height          =   252
            Index           =   1
            Left            =   4320
            TabIndex        =   33
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label Label2 
            Caption         =   "No. Factura"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   1092
         End
      End
      Begin XtremeSuiteControls.GroupBox fraOperacion 
         Height          =   5532
         Left            =   -69520
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   10572
         _Version        =   1572864
         _ExtentX        =   18648
         _ExtentY        =   9758
         _StockProps     =   79
         Caption         =   "Información General"
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
         BorderStyle     =   2
         Begin MSComCtl2.FlatScrollBar FlatScrollBarCnt 
            Height          =   252
            Left            =   8760
            TabIndex        =   12
            Top             =   960
            Width           =   492
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1638401
         End
         Begin MSComCtl2.FlatScrollBar FlatScrollBarPagador 
            Height          =   252
            Left            =   8760
            TabIndex        =   13
            Top             =   1320
            Width           =   492
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1638401
         End
         Begin XtremeSuiteControls.GroupBox gbCalculo 
            Height          =   2532
            Left            =   5880
            TabIndex        =   18
            Top             =   2040
            Width           =   3852
            _Version        =   1572864
            _ExtentX        =   6794
            _ExtentY        =   4466
            _StockProps     =   79
            Caption         =   "Cálculo"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.CheckBox chkCtaApl 
               Height          =   252
               Left            =   3480
               TabIndex        =   147
               ToolTipText     =   "Aplica cobro por Cuota Mensual"
               Top             =   1440
               Width           =   252
               _Version        =   1572864
               _ExtentX        =   444
               _ExtentY        =   444
               _StockProps     =   79
               BackColor       =   -2147483633
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.ComboBox cboEstado 
               Height          =   312
               Left            =   1560
               TabIndex        =   108
               Top             =   1800
               Width           =   1812
               _Version        =   1572864
               _ExtentX        =   3201
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   1973790
               BackColor       =   16777215
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtMonto 
               Height          =   312
               Left            =   1560
               TabIndex        =   134
               Top             =   360
               Width           =   1812
               _Version        =   1572864
               _ExtentX        =   3196
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPlazo 
               Height          =   312
               Left            =   2640
               TabIndex        =   135
               Top             =   720
               Width           =   732
               _Version        =   1572864
               _ExtentX        =   1291
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtTasa 
               Height          =   312
               Left            =   2640
               TabIndex        =   136
               Top             =   1080
               Width           =   732
               _Version        =   1572864
               _ExtentX        =   1291
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCuota 
               Height          =   312
               Left            =   1560
               TabIndex        =   137
               Top             =   1440
               Width           =   1812
               _Version        =   1572864
               _ExtentX        =   3196
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.DateTimePicker dtpInicioFecha 
               Height          =   330
               Left            =   2040
               TabIndex        =   149
               Top             =   2160
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2355
               _ExtentY        =   582
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
            Begin VB.Label lblInicioCuota 
               Caption         =   "Inicio de Cuota:"
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
               Left            =   600
               TabIndex        =   148
               Top             =   2160
               Width           =   1332
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
               Index           =   10
               Left            =   600
               TabIndex        =   89
               Top             =   1800
               Width           =   732
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   252
               Left            =   600
               TabIndex        =   22
               Top             =   1440
               Width           =   732
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   252
               Index           =   0
               Left            =   600
               TabIndex        =   21
               Top             =   1080
               Width           =   492
            End
            Begin VB.Label Label1 
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
               Index           =   4
               Left            =   600
               TabIndex        =   20
               Top             =   360
               Width           =   852
            End
            Begin VB.Label lblPlazo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   252
               Left            =   600
               TabIndex        =   19
               Top             =   720
               Width           =   732
            End
         End
         Begin XtremeSuiteControls.GroupBox gbFactoreo 
            Height          =   2172
            Left            =   0
            TabIndex        =   23
            Top             =   2040
            Width           =   3732
            _Version        =   1572864
            _ExtentX        =   6583
            _ExtentY        =   3831
            _StockProps     =   79
            Caption         =   "Factoreo"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.CheckBox chkAdelantoComisionApl 
               Height          =   252
               Left            =   3240
               TabIndex        =   92
               Top             =   1080
               Width           =   732
               _Version        =   1572864
               _ExtentX        =   1291
               _ExtentY        =   444
               _StockProps     =   79
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Appearance      =   16
            End
            Begin XtremeSuiteControls.PushButton btnFacturas 
               Height          =   312
               Left            =   2520
               TabIndex        =   24
               ToolTipText     =   "Detalle de Facturas"
               Top             =   360
               Width           =   1092
               _Version        =   1572864
               _ExtentX        =   1926
               _ExtentY        =   550
               _StockProps     =   79
               Caption         =   "Facturas"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.PushButton btnActualizaPorcentaje 
               Height          =   312
               Left            =   2520
               TabIndex        =   86
               ToolTipText     =   "Detalle de Facturas"
               Top             =   720
               Width           =   1092
               _Version        =   1572864
               _ExtentX        =   1926
               _ExtentY        =   550
               _StockProps     =   79
               Caption         =   "Actualiza"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.FlatEdit txtFacturasNo 
               Height          =   312
               Left            =   1800
               TabIndex        =   138
               Top             =   360
               Width           =   612
               _Version        =   1572864
               _ExtentX        =   1080
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
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtAdelantoPorcentaje 
               Height          =   312
               Left            =   1800
               TabIndex        =   139
               Top             =   720
               Width           =   612
               _Version        =   1572864
               _ExtentX        =   1080
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
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtAdelantoComision 
               Height          =   312
               Left            =   1800
               TabIndex        =   140
               Top             =   1080
               Width           =   612
               _Version        =   1572864
               _ExtentX        =   1080
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
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtAdelantoDias 
               Height          =   312
               Left            =   2520
               TabIndex        =   141
               Top             =   1080
               Width           =   612
               _Version        =   1572864
               _ExtentX        =   1080
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
               Text            =   "5"
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtAdelantoTotal 
               Height          =   312
               Left            =   1800
               TabIndex        =   142
               Top             =   1440
               Width           =   1812
               _Version        =   1572864
               _ExtentX        =   3196
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtFacturasPendiente 
               Height          =   312
               Left            =   1800
               TabIndex        =   143
               Top             =   1800
               Width           =   1812
               _Version        =   1572864
               _ExtentX        =   3196
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label1 
               Caption         =   "(%) Comisión"
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
               Index           =   6
               Left            =   360
               TabIndex        =   91
               Top             =   1080
               Width           =   1452
            End
            Begin VB.Label Label1 
               Caption         =   "No. Facturas"
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
               Left            =   360
               TabIndex        =   28
               Top             =   360
               Width           =   1452
            End
            Begin VB.Label Label1 
               Caption         =   "(%) Adelanto"
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
               Left            =   360
               TabIndex        =   27
               Top             =   720
               Width           =   1452
            End
            Begin VB.Label Label1 
               Caption         =   "Total Adelanto"
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
               Left            =   360
               TabIndex        =   26
               Top             =   1440
               Width           =   1452
            End
            Begin VB.Label Label1 
               Caption         =   "Monto Pendiente"
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
               TabIndex        =   25
               Top             =   1800
               Width           =   1452
            End
         End
         Begin XtremeSuiteControls.GroupBox gbDesembolso 
            Height          =   1812
            Left            =   0
            TabIndex        =   29
            Top             =   4560
            Width           =   9732
            _Version        =   1572864
            _ExtentX        =   17166
            _ExtentY        =   3196
            _StockProps     =   79
            Caption         =   "Desembolso y Estado de la Operación"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox cboCuenta 
               Height          =   312
               Left            =   5880
               TabIndex        =   105
               Top             =   600
               Width           =   3852
               _Version        =   1572864
               _ExtentX        =   6800
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   1973790
               BackColor       =   16777215
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboBanco 
               Height          =   312
               Left            =   120
               TabIndex        =   106
               Top             =   600
               Width           =   3852
               _Version        =   1572864
               _ExtentX        =   6800
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   1973790
               BackColor       =   16777215
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
               Height          =   312
               Left            =   3960
               TabIndex        =   107
               Top             =   600
               Width           =   1932
               _Version        =   1572864
               _ExtentX        =   3413
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   1973790
               BackColor       =   16777215
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label1 
               Caption         =   "Cuenta / Cliente"
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
               Index           =   7
               Left            =   5880
               TabIndex        =   146
               Top             =   360
               Width           =   2652
            End
            Begin VB.Label Label1 
               Caption         =   "Cuenta / Desembolso"
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
               Left            =   120
               TabIndex        =   90
               Top             =   360
               Width           =   2292
            End
            Begin VB.Label Label1 
               Caption         =   "Emitir"
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
               Left            =   3960
               TabIndex        =   30
               Top             =   360
               Width           =   852
            End
         End
         Begin MSComCtl2.FlatScrollBar FlatScrollBarAutorizado 
            Height          =   252
            Left            =   8760
            TabIndex        =   87
            Top             =   1680
            Width           =   492
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1638401
         End
         Begin MSComCtl2.FlatScrollBar FlatScrollBarConcepto 
            Height          =   252
            Left            =   8760
            TabIndex        =   96
            Top             =   600
            Width           =   492
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1638401
         End
         Begin XtremeSuiteControls.FlatEdit txtCedula 
            Height          =   312
            Left            =   1800
            TabIndex        =   122
            Top             =   240
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   312
            Left            =   3600
            TabIndex        =   123
            Top             =   240
            Width           =   5052
            _Version        =   1572864
            _ExtentX        =   8911
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
         Begin XtremeSuiteControls.FlatEdit txtConceptoCod 
            Height          =   312
            Left            =   1800
            TabIndex        =   125
            Top             =   600
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConceptoDesc 
            Height          =   312
            Left            =   3600
            TabIndex        =   126
            Top             =   600
            Width           =   5052
            _Version        =   1572864
            _ExtentX        =   8911
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
         Begin XtremeSuiteControls.FlatEdit txtContratoCod 
            Height          =   312
            Left            =   1800
            TabIndex        =   127
            Top             =   960
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtContratoDesc 
            Height          =   312
            Left            =   3600
            TabIndex        =   128
            Top             =   960
            Width           =   5052
            _Version        =   1572864
            _ExtentX        =   8911
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
         Begin XtremeSuiteControls.FlatEdit txtPagadorCed 
            Height          =   312
            Left            =   1800
            TabIndex        =   129
            Top             =   1320
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPagadorNom 
            Height          =   312
            Left            =   3600
            TabIndex        =   130
            Top             =   1320
            Width           =   5052
            _Version        =   1572864
            _ExtentX        =   8911
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
         Begin XtremeSuiteControls.FlatEdit txtAutorizadoCed 
            Height          =   312
            Left            =   1800
            TabIndex        =   131
            Top             =   1680
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAutorizadoNom 
            Height          =   312
            Left            =   3600
            TabIndex        =   132
            Top             =   1680
            Width           =   5052
            _Version        =   1572864
            _ExtentX        =   8911
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
         Begin VB.Label lblAutorizado 
            Caption         =   "Autorizado"
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
            Left            =   480
            TabIndex        =   88
            Top             =   1680
            Width           =   1092
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto"
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
            Left            =   480
            TabIndex        =   17
            Top             =   600
            Width           =   852
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
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
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Width           =   852
         End
         Begin VB.Label lblContrato 
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
            Height          =   252
            Left            =   480
            TabIndex        =   15
            Top             =   960
            Width           =   1092
         End
         Begin VB.Label lblPagador 
            Caption         =   "Pagador"
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
            Left            =   480
            TabIndex        =   14
            Top             =   1320
            Width           =   1092
         End
      End
      Begin MSComctlLib.ListView lswOpciones 
         Height          =   3252
         Left            =   -69400
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   2856
         _ExtentX        =   5027
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   16711680
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   5029
         EndProperty
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   852
         Left            =   -69520
         TabIndex        =   11
         Top             =   6120
         Visible         =   0   'False
         Width           =   9732
         _Version        =   1572864
         _ExtentX        =   17166
         _ExtentY        =   1503
         _StockProps     =   79
         Caption         =   "Notas"
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
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   552
            Left            =   360
            TabIndex        =   133
            Top             =   240
            Width           =   9372
            _Version        =   1572864
            _ExtentX        =   16531
            _ExtentY        =   974
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
      End
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   312
         Left            =   -64720
         TabIndex        =   49
         Top             =   4560
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1572864
         _ExtentX        =   9340
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTesoreria 
         Height          =   372
         Left            =   -61480
         TabIndex        =   53
         Top             =   5760
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Procesa Desembolsos?"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1452
         Index           =   1
         Left            =   -68800
         TabIndex        =   56
         Top             =   600
         Visible         =   0   'False
         Width           =   8052
         _Version        =   1572864
         _ExtentX        =   14203
         _ExtentY        =   2561
         _StockProps     =   79
         Caption         =   "Desembolsos"
         ForeColor       =   4210752
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit feDesembolsoTotal 
            Height          =   312
            Left            =   2760
            TabIndex        =   74
            Top             =   360
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feDesembolsoRealizados 
            Height          =   312
            Left            =   2760
            TabIndex        =   76
            Top             =   720
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feDesembolsoPendiente 
            Height          =   312
            Left            =   2760
            TabIndex        =   78
            Top             =   1080
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit feDesembolsoTransito 
            Height          =   312
            Left            =   5760
            TabIndex        =   80
            Top             =   1080
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3831
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Liberado en Transito:"
            Height          =   252
            Index           =   12
            Left            =   6000
            TabIndex        =   81
            Top             =   840
            Width           =   1812
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Pendiente"
            Height          =   252
            Index           =   11
            Left            =   1080
            TabIndex        =   79
            Top             =   1080
            Width           =   1452
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Realizados"
            Height          =   252
            Index           =   10
            Left            =   1080
            TabIndex        =   77
            Top             =   720
            Width           =   1452
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Total"
            Height          =   252
            Index           =   9
            Left            =   1080
            TabIndex        =   75
            Top             =   360
            Width           =   1452
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtDocNum 
         Height          =   312
         Left            =   -61360
         TabIndex        =   93
         Top             =   5040
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasaMora 
         Height          =   312
         Left            =   -61360
         TabIndex        =   94
         Top             =   5400
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption TituloOpcionesSub 
         Height          =   340
         Left            =   -69400
         TabIndex        =   144
         Top             =   720
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
         _ExtentY        =   600
         _StockProps     =   14
         Caption         =   "Opciones"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   6
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tasa Moratoria"
         Height          =   252
         Index           =   18
         Left            =   -63400
         TabIndex        =   54
         Top             =   5400
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "No. Documento"
         Height          =   252
         Index           =   2
         Left            =   -63400
         TabIndex        =   52
         Top             =   5040
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Index           =   8
         Left            =   -65800
         TabIndex        =   9
         Top             =   4620
         Visible         =   0   'False
         Width           =   972
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   5
      Top             =   8328
      Width           =   11064
      _ExtentX        =   19526
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario Registra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario Autoriza"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fecha Autoriza"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario -> Tesoreria "
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fecha -> Tesoreria "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
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
            Picture         =   "frmCxC_Cuentas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Cuentas.frx":0101
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Cuentas.frx":0220
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Cuentas.frx":0340
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Cuentas.frx":0455
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Cuentas.frx":0573
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Cuentas.frx":069D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Cuentas.frx":07C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Cuentas.frx":08DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Cuentas.frx":09DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_Cuentas.frx":0AF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   688
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   11070
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlbPrincipal"
      MinHeight1      =   330
      Width1          =   2955
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   2520
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   3120
         TabIndex        =   6
         Top             =   30
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   582
         ButtonWidth     =   2487
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Autorización"
               Key             =   "Autorizacion"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Activar"
               Key             =   "Activar"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anular"
               Key             =   "Anular"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
               Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
               Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
               Object.ToolTipText     =   "Imprime el listado seleccionado"
               Object.Tag             =   "1"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Boleta"
                     Text            =   "Boleta de Activación"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Nota"
                     Text            =   "Nota de Activación"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Sep1"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Cesion"
                     Text            =   "Cesión de Facturas"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
               Object.ToolTipText     =   "Ayuda General"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cierra esta ventana"
               Object.Tag             =   "1"
            EndProperty
         EndProperty
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   432
      Left            =   1320
      TabIndex        =   124
      Top             =   480
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   10440
      TabIndex        =   150
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   720
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCxC_Cuentas.frx":0C1B
   End
   Begin VB.Image imgId_Cambio 
      Height          =   252
      Left            =   4320
      Picture         =   "frmCxC_Cuentas.frx":0CA4
      Stretch         =   -1  'True
      ToolTipText     =   "Ajustes y Correcciones"
      Top             =   480
      Width           =   252
   End
   Begin VB.Image ImgAutorizacion 
      Height          =   252
      Left            =   3960
      Top             =   480
      Width           =   252
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      Height          =   252
      Index           =   0
      Left            =   -120
      TabIndex        =   4
      Top             =   480
      Width           =   1212
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   252
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Width           =   4932
   End
End
Attribute VB_Name = "frmCxC_Cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje        As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita          As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso           As Boolean 'Control de Activacion de Controles en proceso de carga
Dim vScroll         As Boolean
Dim vFecha          As Date
Dim mMonto          As Currency, mAutoriza As String, mConcepto As String
Dim mFacturas As Long, mFacturaTotal As Currency, mFacturaAdelanto As Currency, mFacturaPendiente As Currency
Dim mAdelantoPermite As Boolean, mAdelantoModifica As Boolean, mAdelantoPorc As Currency
Dim mAdelantoComision As Currency, mAdelantoComisionApl As Integer, mAdelantoComisionDias As Integer
Dim mCreditoLimite As Currency, mCreditoCerrado As Boolean
Dim mCntPagadorAbierto As Boolean

Private Sub sbCambioFechas()

'If vPaso Then Exit Sub
'
'dtpDocFechaVence.Value = DateAdd("yyyy", 1, dtpDocFechaEmite.Value)
'
'If IsNumeric(txtPagoDias.Tag) Then
'   dtpDocFechaPago.Value = DateAdd("d", CLng(txtPagoDias.Tag), vFecha)
'End If
'
' dtpDocFechaEmite.ToolTipText = "Días Transcurridos desde su emisión: " & DateDiff("d", dtpDocFechaEmite.Value, vFecha)
' dtpDocFechaVence.ToolTipText = "Días Pendientes para su vencimiento: " & DateDiff("d", vFecha, dtpDocFechaVence.Value)
' dtpDocFechaPago.ToolTipText = "Días Pendientes para su pago: " & DateDiff("d", vFecha, dtpDocFechaPago.Value)
' txtPagoDias.Text = DateDiff("d", vFecha, dtpDocFechaPago.Value)
End Sub

Function fxCxC_PersonaNombre(strCedula As String) As String
Dim rsX As New ADODB.Recordset

rsX.Open "select nombre from cxc_Personas where cedula = '" & strCedula & "'", glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 fxCxC_PersonaNombre = ""
Else
 fxCxC_PersonaNombre = IIf(IsNull(rsX!Nombre), "", rsX!Nombre)
End If
rsX.Close

End Function


Function fxCxC_CuentaAhorros(strCedula As String, lngID_Banco As Long) As String
Dim strSQL As String, rsX As New ADODB.Recordset

strSQL = "select dbo.fxSys_Cuentas_Bancarias('" & strCedula & "'," & lngID_Banco & ",0) as 'Cuenta'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
  fxCxC_CuentaAhorros = ""
Else
  fxCxC_CuentaAhorros = IIf(IsNull(rsX!Cuenta), "", rsX!Cuenta)
End If
rsX.Close
End Function



Private Sub bntFacturas_Vincular_Click()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = ""

With lswAdelantadas.ListItems
  For i = 1 To .Count
    If .Item(i).Checked = True Then
        strSQL = strSQL & Space(10) & "exec spCxC_Operacion_Factura_Registra " & txtOperacion.Text & ",'" & .Item(i).Text _
                & "','" & .Item(i).SubItems(2) & "','T" _
                & "'," & CCur(.Item(i).SubItems(3)) & "," & CCur(.Item(i).SubItems(4)) & "," & CCur(.Item(i).SubItems(5)) _
                & ",0,'M'," & CCur(.Item(i).SubItems(8)) _
                & ",'" & .Item(i).SubItems(6) & "','" & .Item(i).SubItems(7) _
                & "','" & glogon.Usuario & "','I',1,0," & .Item(i).SubItems(1)
         
         'Procesa Lote
         If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            If Not glogon.error Then
                strSQL = ""
            End If
         End If
      
    End If
  Next i
    
    'Procesa Lote Final
    If Len(strSQL) > 0 Then
       Call ConectionExecute(strSQL)
       If Not glogon.error Then
           strSQL = ""
       End If
    End If
    
End With

'Calcular Totales
strSQL = "exec spCxC_Operacion_Facturas_Actualiza " & txtOperacion.Text & ",1,'" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

txtMonto.Text = Format(rs!Total, "Standard")

txtFacturasNo.Text = rs!facturas
txtAdelantoTotal.Text = Format(rs!Adelanto, "Standard")
txtFacturasPendiente.Text = Format(rs!Pendiente, "Standard")

rs.Close



Me.MousePointer = vbDefault

Call sbFacturas_Load

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnActualizaPorcentaje_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Mid(cboEstado.Text, 1, 1) = "A" Or Mid(cboEstado.Text, 1, 1) = "D" Then
    MsgBox "La Operación no está Pendiente o Recibida, no pueden realizarse los cambios!", vbInformation
    Exit Sub
End If


If Mid(lblAutorizaEstado.Caption, 1, 1) <> "P" Then
    MsgBox "La Operación ya fue autorizada o denegada!", vbInformation
    Exit Sub
End If



Me.MousePointer = vbHourglass

strSQL = "update cxc_cuentas set ADELANTO_PORCENTAJE = " & CCur(txtAdelantoPorcentaje.Text) _
       & ", ADELANTO_COMISION_APL = " & chkAdelantoComisionApl.Value _
       & ", ADELANTO_COMISION = " & CCur(txtAdelantoComision.Text) _
       & ", ADELANTO_COMISION_DIAS = " & CLng(txtAdelantoDias.Text) _
       & " where Operacion = " & txtOperacion.Text & " and estado in('R','P')"
Call ConectionExecute(strSQL)

strSQL = "exec spCxC_Operacion_Facturas_Actualiza " & txtOperacion.Text & ",1"
Call OpenRecordSet(rs, strSQL)

txtMonto.Text = Format(rs!Total, "Standard")

txtFacturasNo.Text = rs!facturas
txtAdelantoTotal.Text = Format(rs!Adelanto, "Standard")
txtFacturasPendiente.Text = Format(rs!Pendiente, "Standard")

rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnAdjuntos_Click()

 gGA.Modulo = "CXC"
 gGA.Llave_01 = txtCedula.Text
 gGA.Llave_02 = txtOperacion.Text
 gGA.Llave_03 = txtConceptoCod.Text
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End Sub

Private Sub btnArchivoBusca_Click()

With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Depósitos del Banco [Microsoft EXCEL]"
    .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
    .ShowOpen

    If .FileName = "" Then
        MsgBox "Archivo no válido...", vbExclamation
        Exit Sub
    End If

    If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
        'Ok
    Else
        MsgBox "La Extensión del Archivo no es válido...", vbExclamation
        Exit Sub
    End If

    feArchivo.Text = .FileName
End With

End Sub

Private Sub btnArchivoCarga_Click()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset

Dim i As Integer, iCampos As Integer

Dim pFactura As String, pFecha As Date, pFechaPago As Date, pDivisa As String
Dim pImporte As Currency, pMonto As Currency, pTC As Currency, pEstado As String, pAdelanto As Currency

On Error GoTo vError

If feArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If


Me.MousePointer = vbHourglass

Set rsExcel = Excel_Load(feArchivo.Text, "Import")

'Verifica Estructura del Archivo
iCampos = 0
For i = 0 To rsExcel.Fields.Count - 1
   Select Case UCase(rsExcel.Fields(i).Name)
      Case "FACTURA", "FECHA_EMITE", "DIVISA", "TIPO_CAMBIO", "IMPORTE", "MONTO", "FECHA_PAGO", "ESTADO", "ADELANTO"
        iCampos = iCampos + 1
      Case Else
      
   End Select
Next i


If iCampos < 8 Then
   Me.MousePointer = vbDefault
   MsgBox "1. No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "2. Los campos son Factura, Fecha_Emite, Divisa, Tipo_Cambio, Importe, Monto, Fecha_Pago, Estado, Adelanto" & vbCrLf & _
          "3. La Hoja de Carga tiene que tener el nombre de: Import", vbExclamation
 
   Exit Sub
End If

Dim pAdelantoInd As Integer, pAdelantoTipo As String


strSQL = ""
Do While Not rsExcel.EOF
    pAdelantoInd = 0
    pAdelantoTipo = "M"
    
    pFactura = Trim(rsExcel!factura)
    pFecha = rsExcel!fecha_emite
    pFechaPago = rsExcel!Fecha_Pago
    pImporte = rsExcel!Importe
    pTC = rsExcel!TIPO_CAMBIO
    pDivisa = Trim(rsExcel!Divisa & "")
    pEstado = Trim(rsExcel!Estado & "")
    pAdelanto = rsExcel!Adelanto
    
'    pMonto = rsExcel!Monto
    pMonto = pImporte * pTC
        
    If pAdelanto > pMonto Then
       pAdelanto = pMonto
    End If
    
    If pAdelanto > 0 Or pEstado = "A" Then
        pAdelantoInd = 1
    End If
    
    'Adelanta
    If pEstado = "A" And pAdelanto = 0 Then
        pAdelantoTipo = "P"
        pAdelantoInd = 1
    End If
        
    strSQL = strSQL & Space(10) & "exec spCxC_Operacion_Factura_Registra " & txtOperacion.Text & ",'" & pFactura _
            & "','" & pDivisa _
            & "','" & pEstado _
            & "'," & pImporte & "," & pTC & "," & pMonto _
            & "," & pAdelantoInd & ",'" & pAdelantoTipo & "'," & pAdelanto _
            & ",'" & Format(pFecha, "yyyy/mm/dd") & "','" & Format(pFechaPago, "yyyy/mm/dd") _
            & "','" & glogon.Usuario & "','I',0,0"
  
   rsExcel.MoveNext
Loop
rsExcel.Close
    
'Procesa Lote
Call ConectionExecute(strSQL)

    
'Actualiza Datos
strSQL = "exec spCxC_Operacion_Facturas_Actualiza " & txtOperacion.Text & ",1"
Call OpenRecordSet(rs, strSQL)

    txtMonto.Text = Format(rs!Total, "Standard")
    
    txtFacturasNo.Text = rs!facturas
    txtAdelantoTotal.Text = Format(rs!Adelanto, "Standard")
    txtFacturasPendiente.Text = Format(rs!Pendiente, "Standard")


rs.Close
    
       
'Totales
feArchivo.Text = ""

Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente", vbInformation


'Cargar Facturas Registradas (Totalizar)
Call sbFacturas_Load


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

    
End Sub

Private Sub sbFactura_Mantenimiento(Optional pMovimiento As String = "I")
Dim strSQL As String, rs As New ADODB.Recordset
Dim pAdelanta As Integer, pOpOrigen As Long

On Error GoTo vError


'Valida el Estado de la Operación

If Mid(cboEstado.Text, 1, 1) = "A" Or Mid(cboEstado.Text, 1, 1) = "D" Then
    MsgBox "La Operación no está Pendiente o Recibida, no pueden realizarse los cambios!", vbInformation
    Exit Sub
End If


If Mid(lblAutorizaEstado.Caption, 1, 1) <> "P" Then
    MsgBox "La Operación ya fue autorizada o denegada!", vbInformation
    Exit Sub
End If

If Trim(feFactura.Text) = "" Or Not IsNumeric(feFacturaImporte.Text) Then
    MsgBox "El número de factura o el Importe no es válido!", vbInformation
    Exit Sub
End If

If CCur(feFacturaImporte.Text) <= 0 Then
    MsgBox "Importe no es válido!", vbInformation
    Exit Sub
End If

'Valida la Factura
If pMovimiento <> "E" Then
    strSQL = "select dbo.fxCxC_FacturaValida(" & txtOperacion.Text & ",'" & feFactura.Text & "') as 'Pass'"
    Call OpenRecordSet(rs, strSQL)
    If rs!Pass = 0 Then
        MsgBox "Esta factura ya ha sido utiliza anteriormente con este cliente!", vbInformation
        Exit Sub
    End If
    rs.Close
End If

Me.MousePointer = vbHourglass


If chkAdelanta.Enabled Then
   pAdelanta = chkAdelanta.Value
Else
  pAdelanta = 0
End If

If IsNumeric(feFactura.Tag) Then
   pOpOrigen = feFactura.Tag
Else
   pOpOrigen = 0
End If

strSQL = "exec spCxC_Operacion_Factura_Registra " & txtOperacion.Text & ",'" & feFactura.Text _
        & "','" & cbofacturasDivisas.ItemData(cbofacturasDivisas.ListIndex) _
        & "','" & cboFacturaEstado.ItemData(cboFacturaEstado.ListIndex) _
        & "'," & CCur(feFacturaImporte.Text) & "," & CCur(feTipoCambio.Text) & "," & CCur(feFacturaMonto.Text) _
        & "," & pAdelanta & ",'" & Mid(cboAdelantoTipo.Text, 1, 1) & "'," & CCur(feAdelanto.Text) _
        & ",'" & Format(dtpFacturaEmite.Value, "yyyy/mm/dd") & "','" & Format(dtpFacturaPago.Value, "yyyy/mm/dd") _
        & "','" & glogon.Usuario & "','" & pMovimiento & "',1,1," & pOpOrigen
Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then
    txtMonto.Text = Format(rs!Total, "Standard")
    
    txtFacturasNo.Text = rs!facturas
    txtAdelantoTotal.Text = Format(rs!Adelanto, "Standard")
    txtFacturasPendiente.Text = Format(rs!Pendiente, "Standard")
    
    rs.Close
End If

feFactura.Tag = ""

Me.MousePointer = vbDefault

Call sbFacturas_Load

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnFacturaElimina_Click()
Call sbFactura_Mantenimiento("E")
End Sub

Private Sub btnFacturaRegistra_Click()
Call sbFactura_Mantenimiento("I")
End Sub

Private Sub btnFacturas_Click()

If txtOperacion.Text = "" Then
  If fxVerificaRecepcion Then
    Call sbGuardar
    txtOperacion.Enabled = True
    
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    
    fraOperacion.Enabled = False
    txtNotas.Locked = True
  
    ssTab.Item(1).Selected = True
  
  Else
    MsgBox vMensaje, vbCritical
  End If
  
End If



End Sub

Private Sub cboAdelantoTipo_Change()
If Mid(cboAdelantoTipo.Text, 1, 1) = "M" Then
  feAdelanto.Enabled = True
Else
  feAdelanto.Enabled = False
End If

feAdelanto.Text = 0
  
End Sub

Private Sub cboAdelantoTipo_Click()
If Mid(cboAdelantoTipo.Text, 1, 1) = "M" Then
  feAdelanto.Enabled = True
Else
  feAdelanto.Enabled = False
End If

feAdelanto.Text = 0

End Sub

Private Sub cboAdelantoTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And feAdelanto.Enabled Then feAdelanto.SetFocus

End Sub

Private Sub cboBanco_Click()

If vPaso Or cboBanco.ListCount = 0 Then Exit Sub

Dim strSQL As String

On Error GoTo vError

'spSys_Cuentas_Bancarias(@Identificacion varchar(30), @BancoId int, @DivisaCheck smallint = 0)

strSQL = "exec spSys_Cuentas_Bancarias '" & txtCedula.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCuenta.SetFocus

End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboBanco.SetFocus
End Sub


Private Sub cbofacturasDivisas_Click()
If vPaso Then Exit Sub

With glogon
  .strSQL = "select [dbo].[fxCntXTipoCambio](" & GLOBALES.gEnlace & ",'" & cbofacturasDivisas.ItemData(cbofacturasDivisas.ListIndex) _
            & "','" & Format(vFecha, "yyyy/mm/dd") & "','V') AS 'TipoCambio'"
  Call OpenRecordSet(.Recordset, .strSQL)
  
  feTipoCambio.Text = CStr(.Recordset!TipoCambio)
  
  .Recordset.Close
End With


End Sub

Private Sub cbofacturasDivisas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then feFacturaImporte.SetFocus
End Sub

Private Sub cboTipoDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboBanco.SetFocus
End Sub


Private Sub sbReporte_Boleta()
Dim strRuta As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

strRuta = SIFGlobal.fxPathReportes("CxC_BoletaActivacion.rpt")

With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "CxC...: Boleta de Activación"
 .ReportFileName = strRuta
 
 .Connect = glogon.ConectRPT
 
 .SelectionFormula = "{CXC_CUENTAS.OPERACION}=" & txtOperacion.Text
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy HH:MM:ss") & "'"
 .Formulas(1) = "Usuario='" & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"

 .SubreportToChange = "sbAsiento"
 .StoredProcParam(0) = "CxC_FRM"
 .StoredProcParam(1) = txtOperacion.Text
 .StoredProcParam(2) = 0

 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub


Private Sub sbReporte_Cesion()
Dim strRuta As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

strRuta = SIFGlobal.fxPathReportes("CxC_Notificacion_Cesion.rpt")

With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "CxC...: Notificación de Cesión de Facturas"
 .ReportFileName = strRuta
 
 .Connect = glogon.ConectRPT
 

 .StoredProcParam(0) = txtOperacion.Text
 .StoredProcParam(1) = glogon.Usuario
 
 .SubreportToChange = "sbFacturas"
 .SelectionFormula = "{CXC_CUENTAS_FACTURAS.OPERACION}=" & txtOperacion.Text


 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub


Private Sub sbActivar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curRebajos As Currency, curMonto As Currency, curIngresos As Currency
Dim vTransac As Boolean, vFecha As Date, vEmitirTipo As String
Dim vTesoreria As Integer

Me.MousePointer = vbHourglass

curMonto = CCur(txtMonto.Text)

vEmitirTipo = fxTipoDocumento(cboTipoDocumento)
vTransac = False

'Calcula Rebajos Totales
strSQL = "select dbo.fxCxC_CuentaRebajos(" & txtOperacion.Text & ",'TOT') as 'Rebajos'" _
       & ",isnull(dbo.fxCxC_CuentaIngresos(" & txtOperacion.Text & "),0) as 'Ingresos'" _
       & ",Genera_Desembolso,dbo.MyGetdate() as 'Fecha_Server'" _
       & " from CxC_Conceptos where cod_concepto = '" & txtConceptoCod.Text & "'"
Call OpenRecordSet(rs, strSQL)
    curRebajos = rs!Rebajos
    curIngresos = rs!Ingresos
    vTesoreria = rs!Genera_Desembolso
    vFecha = rs!Fecha_Server
rs.Close



On Error GoTo vError

'Activa la Operacion
strSQL = "update CxC_Cuentas set Estado = 'A', Activa_Fecha = dbo.MyGetdate(), Activa_Usuario = '" & glogon.Usuario _
       & "', Rebajos_total = " & curRebajos & ",INGRESOS_TOTAL = " & curIngresos _
       & ",Desembolso_Monto = Monto + " & curIngresos & " - " & curRebajos _
       & ", Num_documento = '" & txtDocNum.Text & "'" _
       & " where Operacion = " & txtOperacion.Text

'Ajuste del Desembolso Pendiente

If gbFactoreo.Visible = True Then
    strSQL = strSQL & Space(10) & "update CxC_Cuentas set Desembolso_Pendiente = " _
           & "   case when (Desembolso_Realizado + Desembolso_Pendiente) > Desembolso_Monto then Desembolso_Monto - Desembolso_Realizado else Desembolso_Pendiente end " _
           & " where Operacion = " & txtOperacion.Text
Else
    strSQL = strSQL & Space(10) & "update CxC_Cuentas set Desembolso_Pendiente = Desembolso_Monto - Desembolso_Realizado" _
           & " where Operacion = " & txtOperacion.Text
End If

Call ConectionExecute(strSQL)


'Procesa Activación>


'Inicia Transacciones
glogon.Conection.BeginTrans
vTransac = True



'Actualizar el Registro de Envio a Tesoreria
'Or vEmitirTipo = "ND"
If (curMonto + curIngresos) <= curRebajos Or vTesoreria = 0 Then
    strSQL = "update CxC_Cuentas set Tesoreria_Fecha = dbo.MyGetdate(), Tesoreria_Solicitud = 0 " _
           & ",Tesoreria_Estado = 'C', Tesoreria_Usuario = '" & glogon.Usuario _
           & "' where Operacion = " & txtOperacion.Text
    Call ConectionExecute(strSQL)
Else
    strSQL = "update CxC_Cuentas set Tesoreria_Estado = 'P'" _
           & " where Operacion = " & txtOperacion.Text
    Call ConectionExecute(strSQL)
End If


'Cambia Los Procesos Anteriores x StoreProcedures
strSQL = "exec spCxC_CuentaActivaDetalle " & txtOperacion.Text & ",'" & Format(vFecha, "yyyy/mm/dd") _
       & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

'Cierra Transacciones
glogon.Conection.CommitTrans
vTransac = False


'BITACORA
Call Bitacora("Registra", "Activación de la OP: " & txtOperacion)

'Imprime Boleta de Activacion
Call sbReporte_Boleta

Me.MousePointer = vbDefault

MsgBox "Activación aplicada Satisfactoriamente...", vbInformation

Exit Sub

vError:
 If vTransac Then glogon.Conection.RollbackTrans
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
  
End Sub

Private Sub sbAnular()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spCxC_Cuenta_Anulacion " & txtOperacion.Text & ",'" & glogon.Usuario & "',''"
Call ConectionExecute(strSQL)


'BITACORA
Call Bitacora("Registra", "Anulación de la Operacion CxC No." & txtOperacion.Text)

Me.MousePointer = vbDefault

MsgBox "Anulación de la Operacion de CxC realizada satisfactoriamente!", vbInformation

'Imprime Nota de Reversion
Call sbImprimeRecibo(Trim(txtOperacion.Text) & "-A", "CxC_FRM")

'Imprime Boleta de Activacion/Anulacion
Call sbReporte_Boleta


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxVerificaRecepcion() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String, rs As New ADODB.Recordset
Dim Porcentaje As Currency, pOperacion As Long

fxVerificaRecepcion = True
vMensaje = ""


pOperacion = IIf(IsNumeric(txtOperacion.Text), txtOperacion.Text, 0)

If IsNumeric(txtPlazo) Then
 If txtPlazo < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado es Inválido"
End If

If IsNumeric(txtTasa) Then
 If txtTasa < 0 Or txtTasa > 100 Then vMensaje = vMensaje & vbCrLf & "- La Tasa solicitada no es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Interés Solicitado es Inválido"
End If

If IsNumeric(txtMonto.Text) Then
 If txtMonto.Text < 1 And gbFactoreo.Visible = False Then vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado es Inválido"
End If


'VERIFICAR SI TIENE CODIFICACION CONTABLE

strSQL = "select * from CxC_Conceptos where cod_Concepto ='" & txtConceptoCod.Text & "'"
Call OpenRecordSet(rsX, strSQL, 0)
 If rsX.EOF And rsX.BOF Then
   vMensaje = vMensaje & vbCrLf & "- No existe o no está activo el código de concepto de CxC a utilizar"
 Else
   
  
''   'Verifica el documento
''   If rsX!Requiere_Documento = 1 Or rsX!Proceso_Descuento = 1 Then
''      If txtDocNum.Text = "" Then
''        vMensaje = vMensaje & vbCrLf & "- No se especificó el No. de Documento de referencia para esta cuenta"
''      Else
''        strSQL = "select count(*) as 'Existe' from CxC_Cuentas where cedula = '" & txtCedula.Text _
''               & "' and num_documento = '" & txtDocNum.Text & "' and Operacion <> " & pOperacion
''        Call OpenRecordSet(rs, strSQL)
''        If rs!Existe >= 1 Then vMensaje = vMensaje & vbCrLf & "- El No. de Documento ya ha sido utilizado antes, verifique!"
''        rs.Close
''
''        If DateDiff("d", dtpDocFechaEmite.Value, dtpDocFechaVence.Value) <= 0 Then vMensaje = vMensaje & vbCrLf & "- Existe un error Entre la Fecha de Emisión y la de Vencimiento"
''        If DateDiff("d", dtpDocFechaEmite.Value, dtpDocFechaPago.Value) <= 0 Then vMensaje = vMensaje & vbCrLf & "- Existe un error Entre la Fecha de Emisión y la de Pago"
''        If DateDiff("d", dtpDocFechaPago.Value, dtpDocFechaVence.Value) <= 0 Then vMensaje = vMensaje & vbCrLf & "- Existe un error Entre la Fecha de Pago y la de vencimiento"
''
'''        If DateDiff("d", dtpDocFechaEmite.Value, vFecha) <= 0 Then vMensaje = vMensaje & vbCrLf & "- La fecha de emisión del documento no es válida"
''        If DateDiff("d", vFecha, dtpDocFechaVence.Value) <= 0 Then vMensaje = vMensaje & vbCrLf & "- El documento se encuentra vencido (fecha de vencimiento), verifique!"
'''        If DateDiff("d", vFecha, dtpDocFechaPago.Value) <= 0 Then vMensaje = vMensaje & vbCrLf & "- La fecha de pago no está pendiente?"
''
''
''      End If
''
''   End If
  
   'Verifica el Contrato
   If (rsX!Requiere_Contrato = 1 Or rsX!Proceso_Descuento = 1) And txtContratoCod.Text = "" Then
      vMensaje = vMensaje & vbCrLf & "- Es necesario el uso de algún contrato activo para esta cuenta"
   End If
   If (rsX!Requiere_Contrato = 1 Or rsX!Proceso_Descuento = 1) And txtContratoCod.Text <> "" Then
      strSQL = "select count(*) as 'Existe'" _
             & " from CxC_Contratos Cnt left join CxC_Personas_Contratos Per on Cnt.Cod_Contrato = Per.cod_contrato" _
             & " and Per.cedula = '" & txtCedula.Text _
             & "'    inner join cxc_Conceptos_Contratos Cc on Cnt.Cod_Contrato = Cc.cod_Contrato" _
             & " Where Cnt.cod_Contrato = '" & txtContratoCod.Text & "' and Cc.cod_concepto = '" & txtConceptoCod.Text _
             & "' and ( Per.Cedula is not null or Cnt.Suscripcion_Abierta = 1)"
      
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El concepto de CxC requiere que exista un contrato registrado a la persona y asociado a este concepto."
        rs.Close
   End If
   
   
   'Verifica Proceso de Descuento (revisa al pagador)
   If rsX!Proceso_Descuento = 1 And Len(vMensaje) = 0 Then
     
     'Verifica al Pagador
     If Len(txtPagadorCed.Text) = 0 Then
       vMensaje = vMensaje & vbCrLf & "- El Pagador no está registrado bajo el contrato individual (x Persona)"
     Else
       
        strSQL = "select count(*) as 'Existe'" _
               & " from CxC_Contratos_Pagadores Cp inner join  CxC_Contratos Cn on Cp.Cod_Contrato = Cn.Cod_Contrato" _
               & " inner join CxC_Personas Per on Cp.cedula = Per.cedula" _
               & "  left join CxC_Personas_Contratos_Pagadores PcP on Cp.Cod_Contrato = PcP.cod_Contrato" _
               & " and Cp.Cedula = PcP.cedula_Pagador and PcP.cedula = '" & txtCedula.Text & "'" _
               & " Where Cn.Cod_Contrato = '" & txtContratoCod.Text & "'" _
               & " and (PcP.cedula is not null or Cn.Pagadores_Abierto = 1)" _
               & " and Cp.Cedula = '" & txtPagadorCed.Text & "'"
   
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El Pagador no está registrado bajo el contrato individual (x Persona)"
        rs.Close
      End If
      
      
     'Verifica al Autorizador
     If Len(txtAutorizadoCed.Text) = 0 Then
       vMensaje = vMensaje & vbCrLf & "- No se localizó al Autorizador de la Cesión!"
     Else
       
        strSQL = "select count(*) as 'Existe'" _
               & " from CXC_PERSONAS_AUTORIZADOS " _
               & " Where Cedula_Autorizado = '" & txtAutorizadoCed.Text & "' and cedula = '" & txtCedula.Text & "'"
   
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El Autorizador no está registrado"
        rs.Close
      End If
      
      
   End If
  
  
 End If
rsX.Close

'Validaciones Finales> Consolida Varias Del disponible y Contabilizacion
strSQL = "select dbo.fxCxC_Persona_Disponible_Valida('" & Trim(txtCedula.Text) & "', " & CCur(txtMonto.Text) _
       & ", '" & txtConceptoCod.Text & "') as 'Resultado'"
Call OpenRecordSet(rs, strSQL)
If Len(rs!Resultado) > 0 Then
    vMensaje = vMensaje & vbCrLf & rs!Resultado
End If

If Len(vMensaje) > 0 Then
    fxVerificaRecepcion = False
    MsgBox vMensaje, vbExclamation
Else
    fxVerificaRecepcion = True
End If

End Function

Private Function fxActivacionVerifica() As Boolean
Dim rs As New ADODB.Recordset, strSQL As String
Dim lngPriDeduc As Long, vFecha As Date
Dim Porcentaje As Double, vMontoRefunde As Currency
Dim curDisponible As Currency, curGiros As Currency
Dim curMontoTmp As Currency, vPriDeducCorte As Long

vMensaje = ""
fxActivacionVerifica = True

vFecha = fxFechaServidor

'Verifica que si la salida es por transferencia, la cuenta de ahorros no este en blanco
If fxTipoDocumento(cboTipoDocumento.Text) = "TE" Then
  If Len(Trim(cboCuenta.Text)) = 0 Then
    vMensaje = vMensaje & vbCrLf & "- No se ha especificado una cuenta de ahorros para realizarle la transferencia..."
  End If
End If



'Calcula Rebajos Totales
strSQL = "select Monto,dbo.fxCxC_CuentaRebajos(" & txtOperacion.Text & ",'TOT') as 'Rebajos'" _
       & ",isnull(dbo.fxCxC_CuentaIngresos(" & txtOperacion.Text & "),0) as 'Ingresos'" _
       & " from CxC_Cuentas where Operacion = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
If rs!Rebajos > (rs!Monto + rs!Ingresos) Then
 vMensaje = vMensaje & vbCrLf & "- El monto de los rebajos es mayor que el monto de la operación más otros ingresos"
End If

'If rs!Rebajos > CCur(feDesembolsoTransito.Text) Then
' vMensaje = vMensaje & vbCrLf & "- El monto de los rebajos es mayor que el Desembolso Liberado!"
'End If

rs.Close

'Verifica que no Existan Facturas Duplicadas
strSQL = "exec spCxC_Operacion_Facturas_Verifica " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vMensaje = vMensaje & vbCrLf & "- Factura No.: " & Trim(rs!cod_Factura) & ", se encuentra registrada en la Operación: " & rs!Operacion
  rs.MoveNext
Loop
rs.Close
 
 
If Len(vMensaje) = 0 Then
'    curDisponible = 0
'    strSQL = "exec spCRDDisponibleRecurso '" & fxCodigoCbo(cboRecursos) & "','" & Format(dtpDesembolso.Value, "yyyy/mm/dd") & "'"
'    Call OpenRecordSet(rsX, strSQL, 0)
'    If Not rsX.EOF And Not rsX.BOF Then
'        curDisponible = IIf(IsNull(rsX!disponible), 0, rsX!disponible)
'    End If
'    rsX.Close
'
'    Call sbResumenOperacion
'    curGiros = CCur(lsw.ListItems.Item(10).SubItems(1)) - CCur(lsw.ListItems.Item(8).SubItems(1)) + CCur(lsw.ListItems.Item(3).SubItems(1)) - CCur(lsw.Tag)
'
'    If curGiros > 0 Then
'        If curDisponible < curGiros Then
'           vMensaje = vMensaje & vbCrLf & " - No Hay disponible en el Recurso, para desembolsar esta Operación..."
'           vMensaje = vMensaje & vbCrLf & " - Monto a Girar : " & Format(curGiros, "Standard") & " - Disponible :  " & Format(curDisponible, "Standard")
'           vMensaje = vMensaje & vbCrLf & " - Monto Faltante para Girar: " & Format(curGiros - curDisponible, "Standard")
'        End If
'    End If
End If

If Len(vMensaje) > 0 Then fxActivacionVerifica = False


End Function

Private Function fxAnulacionVerifica() As Boolean
Dim rs As New ADODB.Recordset, strSQL As String
Dim lngPriDeduc As Long, vFecha As Date
Dim Porcentaje As Double, vMontoRefunde As Currency
Dim rsTmp As New ADODB.Recordset

vMensaje = ""
fxAnulacionVerifica = True


''0. Verificacion base / Solo se pueden anular las formalizaciones del día
'strSQL = "select fechaforp,datediff(d,fechaforp,dbo.MyGetdate()) as Dias from CxC_Cuentas where Operacion = " & Operacion.Operacion
'Call OpenRecordSet(rs, strSQL)
'    vFecha = rs!fechaforp
'    If Abs(rs!Dias) > 0 Then
'      vMensaje = vMensaje & vbCrLf & "- Esta operación fue formalizada un día diferente al actual..."
'    End If
'rs.Close
'
'
''2. Verifica que no se le registren desembolsos, Se deben de anular o eliminar
'strSQL = "select isnull(count(*),0) as Existe from Tes_Transacciones where op = " & Operacion.Operacion _
'       & " and estado <> 'A'"
'Call OpenRecordSet(rs, strSQL)
'If rs!Existe > 0 Then
'  vMensaje = vMensaje & vbCrLf & "- Existen solicitudes o documentos emitidos (Cheques/Transferencias) en Tesorería (Proceda a Anularlos)"
'End If
'rs.Close
'
'If GLOBALES.SysPlanPagos = 0 Then
'    '3. Verificar si se le han realizado movimientos a la Operacion despues de su formalizacion
'    strSQL = "select isnull(count(*),0) as Existe from creditos_dt where Operacion = " & Operacion.Operacion _
'           & " and ncon <> Operacion"
'    Call OpenRecordSet(rs, strSQL)
'    If rs!Existe > 0 Then
'      vMensaje = vMensaje & vbCrLf & "- Existen movimientos a esta operación después de su formalización"
'    End If
'    rs.Close
'
'    'No tiene porque tener ningun registro de mora
'    strSQL = "select isnull(count(*),0) as Existe from MOROSIDAD where Operacion = " & Operacion.Operacion
'    Call OpenRecordSet(rs, strSQL)
'    If rs!Existe > 0 Then
'      vMensaje = vMensaje & vbCrLf & "- Existen movimientos a esta operación después de su formalización"
'    End If
'    rs.Close
'
'
'    '3a. Verificar si se le han realizado movimientos a las refundiciones (Abonadas o Canceladas)
'    strSQL = "select Operacion,consec from creditos_dt where tcon = 3 and ncon = " & Operacion.Operacion
'    Call OpenRecordSet(rs, strSQL)
'    Do While Not rs.EOF
'     strSQL = "select isnull(count(*),0) as Existe from creditos_dt where Operacion = " _
'            & rs!Operacion & " and consec > " & rs!consec
'     Call OpenRecordSet(rsTmp, strSQL, 0)
'        If rsTmp!Existe > 0 Then
'          vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a la op:" & rs!Operacion _
'                   & " posterior a su refundicion"
'        End If
'     rsTmp.Close
'     rs.MoveNext
'    Loop
'    rs.Close
'
'    '3a. a la fecha de formalizacion (Doble verificacion para movimientos en mora no reflejados)
'    strSQL = "select isnull(count(*),0) as Existe from creditos_dt" _
'            & " where fechas > '" & Format(vFecha, "yyyy/mm/dd") & "' and Operacion in(select Operacion" _
'            & " from creditos_dt where tcon = 3 and Operacion = " & Operacion.Operacion & ")"
'    Call OpenRecordSet(rs, strSQL)
'    If rs!Existe > 0 Then
'       vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a refundiciones posterior a la formalizacion"
'    End If
'    rs.Close
'
'    '3b. a la fecha de formalizacion para Morosidad
'    strSQL = "select isnull(count(*),0) as Existe from morosidad" _
'            & " where Estado = 'C' and fecUlt > '" & Format(vFecha, "yyyy/mm/dd") & "' and Operacion in(select Operacion" _
'            & " from morosidad where estado = 'C' and tcon = 3 and Operacion = " & Operacion.Operacion & ")"
'    Call OpenRecordSet(rs, strSQL)
'    If rs!Existe > 0 Then
'       vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a Mora de refundiciones posterior a la formalizacion"
'    End If
'    rs.Close
'
'End If 'SysPlanPagos = 0
'
''4. No puede anular retenciones
'strSQL = "select retencion from catalogo where codigo = '" & Operacion.Codigo & "'"
'Call OpenRecordSet(rs, strSQL)
'If rs!retencion = "S" Then
'   vMensaje = vMensaje & vbCrLf & "- Este es un código de retención No se puede Anular..."
'End If
'rs.Close


If Len(vMensaje) > 0 Then fxAnulacionVerifica = False


End Function



Public Sub sbGXSegTraIniTlb()
 Call tlbPrincipal_ButtonClick(tlbPrincipal.Buttons.Item(1))
 txtCedula = GLOBALES.gCedulaActual
 txtCedula_LostFocus
 txtConceptoCod.SetFocus
End Sub


Private Sub chkAdelanta_Click()
If chkAdelanta.Value = xtpChecked Then
   cboAdelantoTipo.Enabled = True
   cboAdelantoTipo.Text = "Porcentual"
   Call cboAdelantoTipo_Change
Else
   cboAdelantoTipo.Enabled = False
   feAdelanto.Enabled = False
   feAdelanto.Text = 0
End If

End Sub

Private Sub chkCtaApl_Click()
If chkCtaApl.Value = xtpChecked Then
   lblInicioCuota.Visible = True
   dtpInicioFecha.Visible = True
Else
   lblInicioCuota.Visible = False
   dtpInicioFecha.Visible = False

End If
End Sub

Private Sub feAdelanto_GotFocus()
On Error GoTo vError

feAdelanto.Text = CCur(feAdelanto.Text)

vError:

End Sub

Private Sub feAdelanto_LostFocus()
On Error GoTo vError

feAdelanto.Text = Format(CCur(feAdelanto.Text), "Standard")

vError:

End Sub

'Private Sub dtpDocFechaPago_Change()
'If vPaso Then Exit Sub
'
'dtpDocFechaPago.ToolTipText = "Días Pendientes para su pago: " & DateDiff("d", vFecha, dtpDocFechaPago.Value)
'txtPagoDias.Text = DateDiff("d", vFecha, dtpDocFechaPago.Value)
'
'End Sub
'
'Private Sub dtpDocFechaPago_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
'End Sub
'
'Private Sub dtpDocFechaVence_Change()
'Call sbCambioFechas
'End Sub

'Private Sub dtpDocFechaVence_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpDocFechaPago.SetFocus
'End Sub
'


Private Sub feFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbofacturasDivisas.SetFocus

End Sub

Private Sub feFacturaImporte_Change()
On Error GoTo vError

feFacturaMonto.Text = Format(CCur(feFacturaImporte.Text) * CCur(feTipoCambio.Text), "Standard")

vError:

End Sub

Private Sub feFacturaImporte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then feTipoCambio.SetFocus
End Sub

Private Sub feFacturaImporte_LostFocus()
On Error GoTo vError

feFacturaImporte.Text = Format(CCur(feFacturaImporte.Text), "Standard")

vError:

End Sub

Private Sub feTipoCambio_Change()
On Error GoTo vError

feFacturaMonto.Text = Format(CCur(feFacturaImporte.Text) * CCur(feTipoCambio.Text), "Standard")

vError:

End Sub

Private Sub feTipoCambio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFacturaEmite.SetFocus

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtOperacion.Text = "" Then txtOperacion.Text = "0"
If FlatScrollBar.Tag = "" Then FlatScrollBar.Tag = 0

strSQL = "select Top 1 Operacion from CxC_Cuentas"

If FlatScrollBar.Value > CLng(FlatScrollBar.Tag) Then
   strSQL = strSQL & " where Operacion > " & txtOperacion & " order by Operacion asc"
Else
   strSQL = strSQL & " where Operacion < " & txtOperacion & " order by Operacion desc"
End If

FlatScrollBar.Tag = FlatScrollBar.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtOperacion.Text = rs!Operacion
  Call sbConsulta
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbContratoDetalle(pContrato As String, pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


mCntPagadorAbierto = False

strSQL = "select Cnt.Cod_Contrato, Cnt.Descripcion, Cnt.PAGADORES_ABIERTO" _
       & ", isnull(Per.Tasa_Corriente, Cnt.Tasa_Corriente) as 'Tasa_Corriente'" _
       & ", ISNULL(Per.Tasa_Mora,Cnt.Tasa_Mora) as 'Tasa_Mora', isnull(Per.Plazo,Cnt.Plazo) as 'Plazo'" _
       & " from CxC_Contratos Cnt left join CxC_Personas_Contratos Per on  Cnt.Cod_Contrato = Per.cod_contrato" _
       & " and Per.Activo = 1 and Per.Cedula = '" & pCedula & "'" _
       & " Where Cnt.cod_Contrato = '" & pContrato & "'" _
       & "   and (Per.Cedula is not null or Cnt.Suscripcion_Abierta = 1)"
       
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtContratoDesc.Text = rs!Descripcion & ""
   txtContratoCod.Text = rs!COD_CONTRATO & ""
   
   txtTasa.Text = CStr(rs!Tasa_Corriente)
   txtTasaMora.Text = CStr(rs!Tasa_Mora)
'   txtPagoDias.Tag = rs!Plazo
    
   txtPlazo.Text = CLng(rs!Plazo / 30)
   If rs!PAGADORES_ABIERTO = 1 Then
       mCntPagadorAbierto = True
   Else
       mCntPagadorAbierto = False
   End If
    
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBarAutorizado_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarPagador.Tag = "" Then FlatScrollBarPagador.Tag = 0

strSQL = "select Top 1 Per.Cedula,Per.nombre" _
       & " from CxC_Personas Per" _
       & "  inner join CXC_PERSONAS_AUTORIZADOS Pa on Per.Cedula = Pa.Cedula_Autorizado" _
       & " Where Pa.cedula = '" & txtCedula.Text & "'" _

If FlatScrollBarAutorizado.Value > CLng(FlatScrollBarPagador.Tag) Then
   strSQL = strSQL & " and Pa.Cedula_Autorizado > '" & txtAutorizadoCed.Text & "' order by Pa.Cedula_Autorizado asc"
Else
   strSQL = strSQL & " and Pa.Cedula_Autorizado < '" & txtAutorizadoCed.Text & "' order by Pa.Cedula_Autorizado desc"
End If

FlatScrollBarAutorizado.Tag = FlatScrollBarAutorizado.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtAutorizadoCed.Text = rs!Cedula
  txtAutorizadoNom.Text = rs!Nombre
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBarCnt_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarCnt.Tag = "" Then FlatScrollBarCnt.Tag = 0

strSQL = "select Top 1 Cnt.Cod_Contrato,'" & txtCedula.Text & "' as 'Cedula'" _
       & " from CxC_Conceptos_Contratos Cnt" _
       & "      inner join CxC_Contratos Cn on Cnt.Cod_Contrato = Cn.cod_Contrato" _
       & "       left join CxC_Personas_Contratos Pc on Cnt.cod_Contrato = Pc.cod_Contrato" _
       & " and Cnt.Cod_Concepto = '" & txtConceptoCod.Text & "' and Pc.Cedula = '" & txtCedula.Text & "'" _
       & " Where Cn.Activo = 1 and Cnt.Cod_Concepto = '" & txtConceptoCod.Text & "'" _
       & "   and (Pc.Cedula is not null or Cn.Suscripcion_Abierta = 1)"

If FlatScrollBarCnt.Value > CLng(FlatScrollBarCnt.Tag) Then
   strSQL = strSQL & " and Cn.cod_contrato > '" & txtContratoCod.Text & "' order by Cn.cod_contrato asc"
Else
   strSQL = strSQL & " and Cn.cod_contrato < '" & txtContratoCod.Text & "' order by Cn.cod_contrato desc"
End If

FlatScrollBarCnt.Tag = FlatScrollBarCnt.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  Call sbContratoDetalle(rs!COD_CONTRATO, rs!Cedula)
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub FlatScrollBarConcepto_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarConcepto.Tag = "" Then FlatScrollBarConcepto.Tag = 0

strSQL = "Select Top 1 cod_Concepto,Descripcion  from CxC_Conceptos " _
       & " where Activo = 1"

If FlatScrollBarConcepto.Value > CLng(FlatScrollBarConcepto.Tag) Then
   strSQL = strSQL & " and cod_Concepto > '" & txtConceptoCod.Text & "' order by cod_Concepto asc"
Else
   strSQL = strSQL & " and cod_Concepto < '" & txtConceptoCod.Text & "' order by cod_Concepto desc"
End If

FlatScrollBarConcepto.Tag = FlatScrollBarConcepto.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtConceptoCod.Text = rs!cod_Concepto
  txtConceptoDesc.Text = rs!Descripcion
  txtConceptoCod_LostFocus
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub FlatScrollBarPagador_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarPagador.Tag = "" Then FlatScrollBarPagador.Tag = 0


       
If mCntPagadorAbierto Then
      strSQL = "select Top 1 Cp.Cedula,Cp.Nombre from CxC_Personas Cp" _
             & " Where Cp.Rol_Pagador = 1"
Else
    strSQL = "select Top 1 Cp.Cedula,Per.nombre" _
           & " from CxC_Contratos_Pagadores Cp inner join  CxC_Contratos Cn on Cp.Cod_Contrato = Cn.Cod_Contrato" _
           & " inner join CxC_Personas Per on Cp.cedula = Per.cedula" _
           & "  left join CxC_Personas_Contratos_Pagadores PcP on Cp.Cod_Contrato = PcP.cod_Contrato" _
           & " and Cp.Cedula = PcP.cedula_Pagador and PcP.cedula = '" & txtCedula.Text & "'" _
           & " Where Cn.Cod_Contrato = '" & txtContratoCod.Text & "'" _
           & " and (PcP.cedula is not null or Cn.Pagadores_Abierto = 1)"
End If

If FlatScrollBarPagador.Value > CLng(FlatScrollBarPagador.Tag) Then
   strSQL = strSQL & " and Cp.Cedula > '" & txtPagadorCed.Text & "' order by Cp.Cedula asc"
Else
   strSQL = strSQL & " and Cp.Cedula < '" & txtPagadorCed.Text & "' order by Cp.Cedula desc"
End If

FlatScrollBarPagador.Tag = FlatScrollBarPagador.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtPagadorCed.Text = rs!Cedula
  txtPagadorNom.Text = rs!Nombre
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 31
End Sub

Private Sub Form_Load()
 
 vModulo = 31
 
 vFecha = fxFechaServidor
 
 Call sbToolBarIconos(tlbPrincipal, False)
 
 mAdelantoComisionDias = fxCxC_Parametro("16")
With lswFacturas.ColumnHeaders
    .Clear
    .Add , , "No. Factura", 1500
    .Add , , "Estado", 1200
    .Add , , "Divisa", 900, vbCenter
    .Add , , "Importe", 1300, vbRightJustify
    .Add , , "T.C.", 1100, vbRightJustify
    .Add , , "Monto", 1300, vbRightJustify
    .Add , , "Fec.Emite", 1200
    .Add , , "Fec.Pago", 1200
    .Add , , "Adelanta?", 1200, vbCenter
    .Add , , "Ade.Tipo", 1200
    .Add , , "Ade.Monto", 1200, vbRightJustify
    .Add , , "Pendiente", 1200, vbRightJustify
    .Add , , "Liberado", 1200, vbRightJustify
    .Add , , "Op.Origen", 1200, vbCenter
End With
 
With lswAdelantadas.ColumnHeaders
    .Clear
    .Add , , "No. Factura", 1400
    .Add , , "No. Operación", 1200
    .Add , , "Divisa", 900, vbCenter
    .Add , , "Importe", 1300, vbRightJustify
    .Add , , "T.C.", 1100, vbRightJustify
    .Add , , "Monto", 1300, vbRightJustify
    .Add , , "Fec.Emite", 1200
    .Add , , "Fec.Pago", 1200
    .Add , , "Liberado", 1200, vbRightJustify
    .Add , , "Pendiente", 1200, vbRightJustify
End With
 
 Call sbCargaCombos
 Call sbLimpiaDatos

 With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With

Call Formularios(Me)
Call RefrescaTags(Me)

'Activa Replica de Seguridad a Componentes Asociados la Accion de Editar:
btnFacturaRegistra.Enabled = btnFacturas.Enabled
btnFacturaElimina.Enabled = btnFacturas.Enabled
lswFacturas.Enabled = btnFacturas.Enabled
lswAdelantadas.Enabled = btnFacturas.Enabled

End Sub

Private Sub sbLimpiaDatos()
Dim pLeft As Long, pTop As Long

 pLeft = 240
 pTop = gbCalculo.top

 mAutoriza = "P"
 mMonto = 0
 mConcepto = ""
 mCntPagadorAbierto = False
 
 dtpInicioFecha.Value = Now
 
lblAutorizaEstado.Caption = "En Trámite"
Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(11).Picture
ImgAutorizacion.ToolTipText = "En Proceso: Consulta/Nuevo"
 
 tlbAux.Buttons.Item(1).Enabled = False
 tlbAux.Buttons.Item(3).Enabled = False
 tlbAux.Buttons.Item(5).Enabled = False

 txtCedula = ""
 txtNombre = ""
 lblNombre.Caption = txtNombre.Text
 
 txtConceptoCod.Text = ""
 txtConceptoDesc.Text = ""
   
 txtMonto = "0"
 txtPlazo = "1"
 txtTasa = "0"
 txtCuota = "0"
  
 txtTasaMora.Text = "0"
 
 txtDocNum.Text = ""
 txtContratoCod.Text = ""
 txtContratoDesc.Text = ""
 
 txtPagadorCed.Text = ""
 txtPagadorNom.Text = ""
 
 txtAutorizadoCed.Text = ""
 txtAutorizadoNom.Text = ""
 
 cboCuenta.Clear
 
 txtNotas = ""
 
 cboEstado.Clear
 cboEstado.AddItem "Recibida"
 cboEstado.Text = "Recibida"

 ssTab.Item(0).Selected = True
 ssTab.Item(1).Enabled = False
 ssTab.Item(2).Enabled = False
 
 txtFacturasNo.Text = "0"
 txtAdelantoPorcentaje.Text = "0"
 txtAdelantoTotal.Text = "0"
 txtFacturasPendiente.Text = "0"
 txtAdelantoComision.Text = "0"
 txtAdelantoDias.Text = mAdelantoComisionDias
 chkAdelantoComisionApl.Value = xtpChecked
 
 StatusBarX.Panels(1).Text = ""
 StatusBarX.Panels(2).Text = ""
 StatusBarX.Panels(3).Text = ""
 StatusBarX.Panels(4).Text = ""
 StatusBarX.Panels(5).Text = ""

 chkTesoreria.Value = vbUnchecked
  
  
 gbFactoreo.Left = pLeft
 gbFactoreo.top = pTop
 gbFactoreo.Visible = False
  
 
 feArchivo.Text = ""
 
 feDesembolsoTotal.Text = "0.00"
 feDesembolsoRealizados.Text = "0.00"
 feDesembolsoTransito.Text = "0.00"
 feDesembolsoPendiente.Text = "0.00"
 
 
 Call sbCambioFechas
 
 chkCtaApl.Value = xtpUnchecked
 Call chkCtaApl_Click

End Sub

Private Sub sbCargaCombos()
Dim strSQL As String

vPaso = True

'Carga Divisas
strSQL = "select rtrim(cod_Divisa) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from CntX_Divisas where cod_Contabilidad = " & GLOBALES.gEnlace
Call sbCbo_Llena_New(cbofacturasDivisas, strSQL, False, True)

cboAdelantoTipo.Clear
cboAdelantoTipo.AddItem "Porcentual"
cboAdelantoTipo.AddItem "Monto"
cboAdelantoTipo.Text = "Porcentual"

'Oficinas
strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as 'ItmX' from sif_oficinas where estado = 1"
Call sbCbo_Llena_New(cboOficina, strSQL, False, True)


'Bancos
strSQL = "exec spCxC_Bancos_Autorizados"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("ND")
cboTipoDocumento.Text = fxTipoDocumento("TE")

cboCuenta.Clear

vPaso = False

End Sub




Private Sub sbConsulta()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vFecha As Date, iMes As Integer, lngAnio As Long
Dim i As Integer, vTemp As String

On Error GoTo vError

vPaso = True

strSQL = "select * from vCxC_Cuentas_Consulta" _
       & " where Operacion = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
 ssTab.Item(1).Enabled = True
 ssTab.Item(2).Enabled = True
 
 vFecha = rs!FechaServer

' Call sbCargaCombos

 mAutoriza = rs!Autoriza_Estado
 mMonto = rs!Monto
 mConcepto = Trim(rs!cod_Concepto)
 mCntPagadorAbierto = IIf((rs!PAGADORES_ABIERTO = 1), True, False)
 
 txtCedula.Text = rs!Cedula
 txtNombre.Text = rs!Nombre
 lblNombre.Caption = txtNombre.Text
 
 txtConceptoCod.Text = rs!cod_Concepto
 txtConceptoDesc.Text = rs!ConceptoDesc

 txtPagadorCed.Text = rs!cedula_pagador & ""
 txtPagadorNom.Text = rs!PagadorNom & ""
 
 txtMonto.Text = Format(IIf(IsNull(rs!Monto), 0, rs!Monto), "Standard")
' txtPagoDias.Text = IIf(IsNull(rs!Dias_Plazo), 0, rs!Dias_Plazo)
 txtPlazo.Text = IIf(IsNull(rs!Plazo), 1, rs!Plazo)
 
' If rs!tipo_Plazo = "D" Then
'   'Dias
'    txtPlazo.Text = 1
'    txtPagoDias.Text = IIf(IsNull(rs!Dias_Plazo), 0, rs!Dias_Plazo)
'  Else
'   'Meses
'    txtPlazo.Text = rs!Dias_Plazo
'    txtPagoDias.Text = IIf(IsNull(rs!Dias_Plazo), 0, rs!Dias_Plazo * 30)
' End If
 
 txtTasa.Text = Format(IIf(IsNull(rs!Tasa_Corriente), 0, rs!Tasa_Corriente), "Standard")
 txtTasaMora.Text = Format(CStr(IIf(IsNull(rs!Tasa_Mora), 0, rs!Tasa_Mora)), "Standard")
 
 txtCuota.Text = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
 
 
 txtDocNum.Text = Trim(rs!Num_Documento)
' dtpDocFechaEmite.Value = rs!Fecha_Emision
' dtpDocFechaVence.Value = rs!Fecha_Vencimiento
' dtpDocFechaPago.Value = rs!Fecha_Pago
 
 txtContratoCod.Text = rs!COD_CONTRATO & ""
 txtContratoDesc.Text = rs!ContratoDesc & ""
 

     
 txtAutorizadoCed.Text = rs!Cedula_Autorizado & ""
 txtAutorizadoNom.Text = rs!AutorizadoNom & ""
 
     
 
 Call sbCboAsignaDato(cboOficina, Trim(rs!OficinaX), True, Trim(rs!COD_OFICINA))
 Call sbCboAsignaDato(cboBanco, rs!BancoDesc, True, rs!Emitir_Banco)
 cboTipoDocumento.Text = fxTipoDocumento(IIf(IsNull(rs!Emitir_Tipo), "OT", rs!Emitir_Tipo))
 
 If Not IsNull(rs!Emitir_Cuenta) Then
    Call sbCboAsignaDato(cboCuenta, rs!CuentaDesc, True, rs!Emitir_Cuenta)
 End If
 
 txtNotas = IIf(IsNull(rs!Notas), "", rs!Notas)


 tlbAux.Buttons.Item(1).Enabled = False
 tlbAux.Buttons.Item(3).Enabled = False
 tlbAux.Buttons.Item(5).Enabled = False

 cboEstado.Clear
 Select Case rs!Estado
   Case "R"
      cboEstado.AddItem "Recibida"
      cboEstado.AddItem "Pendiente"
      cboEstado.Text = "Recibida"
      
      Select Case rs!autorizado
       Case "P", "R" 'Pendiente
          tlbAux.Buttons.Item(1).Enabled = True
       Case "A" 'Aprobado
          tlbAux.Buttons.Item(1).Enabled = True
          tlbAux.Buttons.Item(3).Enabled = True
       Case "D" 'Denegado
        'Nada
      End Select
      
      imgId_Cambio.Visible = False
      
   Case "P"
      cboEstado.AddItem "Recibida"
      cboEstado.AddItem "Pendiente"
      cboEstado.Text = "Pendiente"
      imgId_Cambio.Visible = False
   
   Case "A"
      cboEstado.AddItem "Activa"
      cboEstado.Text = "Activa"
      tlbAux.Buttons.Item(5).Enabled = True
      
      imgId_Cambio.Visible = True
   
   Case "C"
      cboEstado.AddItem "Cancelada"
      cboEstado.Text = "Cancelada"
      
      imgId_Cambio.Visible = False
   
   
   Case "N"
      cboEstado.AddItem "Anulada"
      cboEstado.Text = "Anulada"
      
      imgId_Cambio.Visible = False
  
  End Select


 lblRecibe.Caption = IIf(IsNull(rs!REGISTRO_USUARIO), "", rs!REGISTRO_USUARIO)
 lblFechaRec.Caption = rs!REGISTRO_FECHA & ""
 lblFormaliza.Caption = IIf(IsNull(rs!ACTIVA_USUARIO), "", rs!ACTIVA_USUARIO)
 lblFechaFor.Caption = IIf(IsNull(rs!Activa_Fecha), rs!Activa_Fecha & "", rs!Activa_Fecha)
 
 lblTesoreria.Caption = IIf(IsNull(rs!tesoreria_usuario), "", rs!tesoreria_usuario)
 lblFechaTes.Caption = rs!tesoreria_fecha & ""
 
 lblAutorizada.Caption = IIf(IsNull(rs!Autoriza_Usuario), "", rs!Autoriza_Usuario)
 lblFechaAuto.Caption = rs!Autoriza_Fecha & ""
 txtAutorizaNota.Text = rs!Autoriza_Notas & ""
 
 Select Case rs!Autoriza_Estado
   Case "A" 'Aprobado
        lblAutorizaEstado.Caption = "Aprobada"
        Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(5).Picture
        ImgAutorizacion.ToolTipText = "Autorización: Aprobada"
        
        fraOperacion.Enabled = False
   
   Case "D" 'Denegado
        lblAutorizaEstado.Caption = "Denegada"
        Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(6).Picture
        ImgAutorizacion.ToolTipText = "Autorización: Denegada"
        fraOperacion.Enabled = False
   
   Case "P" 'Pendiente
        lblAutorizaEstado.Caption = "Pendiente"
        Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(7).Picture
        ImgAutorizacion.ToolTipText = "Autorización: Pendiente"
        fraOperacion.Enabled = True
   Case Else
        lblAutorizaEstado.Caption = "Pendiente"
        Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(7).Picture
        ImgAutorizacion.ToolTipText = "Autorización: Pendiente"
        fraOperacion.Enabled = True
 
 End Select

 StatusBarX.Panels(1).Text = rs!REGISTRO_USUARIO
 StatusBarX.Panels(2).Text = rs!REGISTRO_FECHA
 StatusBarX.Panels(3).Text = rs!Autoriza_Usuario & ""
 StatusBarX.Panels(4).Text = rs!Autoriza_Fecha & ""
 StatusBarX.Panels(5).Text = rs!tesoreria_usuario & ""
 StatusBarX.Panels(6).Text = rs!tesoreria_fecha & ""
 
 
 feDesembolsoTotal.Text = Format(rs!Desembolso_Monto, "Standard")
 feDesembolsoRealizados.Text = Format(rs!Desembolso_Realizado, "Standard")
 feDesembolsoTransito.Text = Format(rs!DESEMBOLSO_PENDIENTE, "Standard")
 feDesembolsoPendiente.Text = Format(rs!Desembolso_Monto - rs!Desembolso_Realizado, "Standard")
 

 With tlbPrincipal.Buttons
   .Item(1).Enabled = True
   .Item(2).Enabled = True
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
 
 'Bloquea y/o activa campos dependiendo de la configuración del concepto
 Call txtConceptoCod_LostFocus
 
 
 txtFacturasNo.Text = rs!facturas
 txtFacturasPendiente.Text = Format(rs!Monto - rs!Adelanto_Monto, "Standard")
 txtAdelantoTotal.Text = Format(rs!Adelanto_Monto, "Standard")
 txtAdelantoPorcentaje.Text = Format(rs!ADELANTO_PORCENTAJE, "Standard")
 txtAdelantoComision.Text = Format(rs!ADELANTO_COMISION, "Standard")
 txtAdelantoDias.Text = CStr(rs!ADELANTO_COMISION_DIAS)
 chkAdelantoComisionApl.Value = rs!ADELANTO_COMISION_APL
 
 If rs!Autoriza_Estado = "A" Then
    txtAdelantoPorcentaje.Enabled = False
 Else
    txtAdelantoPorcentaje.Enabled = True
 End If
 
 chkCtaApl.Value = rs!CUOTAS_APL
 
 If Not IsNull(rs!Fecha_Inicio) Then
    dtpInicioFecha.Value = rs!Fecha_Inicio
 End If
 
Else
 MsgBox "No existe este número de Operación, verifique!", vbCritical
End If
rs.Close

vPaso = False


' Codigo Corrige Apagado Inecesario del Lsw, porque por error el sistema de seguridad lo toma como de el
i = IIf(lsw.Enabled, 1, 0)
Call RefrescaTags(Me)

lsw.Enabled = IIf((i = 1), True, False)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

  
End Sub




Private Sub imgId_Cambio_Click()

If Not IsNumeric(txtOperacion.Text) Then Exit Sub

GLOBALES.gTag = txtOperacion.Text
Call sbFormsCall("frmCxC_Cuentas_Correcciones", vbModal, , , False, Me)

Call sbConsulta


End Sub


Private Sub lswAdelantadas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswAdelantadas.SortKey = ColumnHeader.Index - 1
  If lswAdelantadas.SortOrder = 0 Then lswAdelantadas.SortOrder = 1 Else lswAdelantadas.SortOrder = 0
  lswAdelantadas.Sorted = True
End Sub

Private Sub lswAdelantadas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curTotal As Currency, curAdelanto As Currency, iContador As Long
Dim i As Integer

'Revisa Toda la Lista para evitar errores de cálculo

curTotal = 0
curAdelanto = 0
iContador = 0

With lswAdelantadas.ListItems

    For i = 1 To .Count
     If .Item(i).Checked Then
            curTotal = curTotal + CCur(.Item(i).SubItems(5))
            curAdelanto = curAdelanto + CCur(.Item(i).SubItems(8))
            iContador = iContador + 1
     End If
    Next i
    txtFacturas_Adelanto_Total_Sel.Text = Format(curTotal, "Standard")
    txtFacturas_Adelanto_Total_Sel.ToolTipText = "Adelanto: " & Format(curAdelanto, "Standard")
    txtFacturas_Adelanto_Casos_Sel.Text = Format(iContador, "###,##0")

End With

End Sub


Private Sub lswFacturas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswFacturas.SortKey = ColumnHeader.Index - 1
  If lswFacturas.SortOrder = 0 Then lswFacturas.SortOrder = 1 Else lswFacturas.SortOrder = 0
  lswFacturas.Sorted = True
End Sub

Private Sub lswFacturas_DblClick()
If vPaso Then Exit Sub
If lswFacturas.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

With lswFacturas.SelectedItem

    feFactura.Text = .Text
    feFactura.Tag = .SubItems(13)
    cbofacturasDivisas.Text = Trim(.SubItems(2))
    feFacturaImporte.Text = .SubItems(3)
    feTipoCambio.Text = .SubItems(4)
    feFacturaMonto.Text = .SubItems(5)
    
    dtpFacturaEmite.Value = .SubItems(6)
    dtpFacturaPago.Value = .SubItems(7)
    
    If .SubItems(8) = "Sí" Then
       chkAdelanta.Value = xtpChecked
    Else
       chkAdelanta.Value = xtpUnchecked
    End If
    
    cboAdelantoTipo.Text = .SubItems(9)
    feAdelanto.Text = .SubItems(10)
   
    'Puede que el estado no exista en el combo
    cboFacturaEstado.Text = Trim(.SubItems(1))
   
End With

vError:


End Sub

Private Sub lswOpciones_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xItem As String, i As Integer, curMonto As Currency

Dim itmX As ListViewItem

xItem = lswOpciones.SelectedItem.Key
With lswOpciones.ListItems
 For i = 1 To .Count
   If .Item(i).Key = xItem Then
      .Item(i).SmallIcon = 9
   Else
      .Item(i).SmallIcon = 8
   End If
 Next i
End With

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

Select Case xItem
 Case "ING"
     lsw.ColumnHeaders.Add , , "Codigo", 900
     lsw.ColumnHeaders.Add , , "Descripción", 2900
     lsw.ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
     lsw.ColumnHeaders.Add , , "Tipo", 1200
     lsw.ColumnHeaders.Add , , "Valor", 1200, vbRightJustify

     curMonto = 0
     strSQL = "select R.cod_cargo,R.descripcion,A.tipo,A.valor,A.monto" _
            & " from CxC_Cargos R inner join CxC_Cuentas_Ingresos A" _
            & " on R.cod_cargo = A.cod_cargo and A.Operacion = " & txtOperacion.Text
     Call OpenRecordSet(rs, strSQL, 0)
     Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , rs!COD_CARGO)
           itmX.SubItems(1) = rs!Descripcion
           itmX.SubItems(2) = Format(rs!Monto, "Standard")
           itmX.SubItems(3) = IIf((rs!Tipo = "P"), "Porcentual", "Monto")
           itmX.SubItems(4) = Format(rs!Valor, "Standard")
           curMonto = curMonto + rs!Monto
       rs.MoveNext
     Loop
     rs.Close

     Set itmX = lsw.ListItems.Add(, , "")
         itmX.SubItems(2) = "________"
     Set itmX = lsw.ListItems.Add(, , "")
         itmX.SubItems(2) = Format(curMonto, "Standard")

 Case "CRD"
   lsw.ColumnHeaders.Add , , "Operación", 980
   lsw.ColumnHeaders.Add , , "Línea", 980
   lsw.ColumnHeaders.Add , , "Monto", 1280, vbRightJustify
   lsw.ColumnHeaders.Add , , "Descripción", 3980

   curMonto = 0
   strSQL = "select Reb.*,Cat.descripcion,Cat.codigo" _
          & " from CxC_Cuentas_Rebajos_Crd Reb inner join Reg_Creditos Crd on Reb.id_solicitud = Crd.id_Solicitud" _
          & " inner join catalogo Cat on Crd.codigo = Cat.codigo" _
          & " Where Reb.Operacion = " & txtOperacion.Text
   Call OpenRecordSet(rs, strSQL, 0)
   Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!ID_SOLICITUD)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = Format(rs!Monto, "Standard")
        itmX.SubItems(3) = rs!Descripcion
        curMonto = curMonto + rs!Monto
    rs.MoveNext
   Loop
   rs.Close
   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(2) = "_______"
   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(2) = Format(curMonto, "Standard")


 Case "CXC"
   lsw.ColumnHeaders.Add , , "Operación", 980
   lsw.ColumnHeaders.Add , , "Concepto", 980
   lsw.ColumnHeaders.Add , , "Monto", 1280, vbRightJustify
   lsw.ColumnHeaders.Add , , "Descripción", 3980

   curMonto = 0
   strSQL = "select R.*,Cta.cod_concepto,C.descripcion as 'ConceptoDesc'" _
          & " from CxC_Cuentas_Rebajos R inner join CxC_Cuentas Cta on R.Operacion_Aplicada = Cta.Operacion" _
          & " inner join CxC_Conceptos C on Cta.cod_concepto = C.cod_Concepto" _
          & " Where R.Operacion = " & txtOperacion.Text
   Call OpenRecordSet(rs, strSQL, 0)
   Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!Operacion_Aplicada)
        itmX.SubItems(1) = rs!cod_Concepto
        itmX.SubItems(2) = Format(rs!Monto, "Standard")
        itmX.SubItems(3) = rs!ConceptoDesc
        curMonto = curMonto + rs!Monto
    rs.MoveNext
   Loop
   rs.Close
   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(2) = "_______"
   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(2) = Format(curMonto, "Standard")


 Case "CAR"
     lsw.ColumnHeaders.Add , , "Codigo", 900
     lsw.ColumnHeaders.Add , , "Descripción", 2900
     lsw.ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
     lsw.ColumnHeaders.Add , , "Tipo", 1200
     lsw.ColumnHeaders.Add , , "Valor", 1200, vbRightJustify

     curMonto = 0
     strSQL = "select R.cod_cargo,R.descripcion,A.tipo,A.valor,A.monto" _
            & " from CxC_Cargos R inner join CxC_Cuentas_Rebajos_Cargos A" _
            & " on R.cod_cargo = A.cod_cargo and A.Operacion = " & txtOperacion.Text
     Call OpenRecordSet(rs, strSQL, 0)
     Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , rs!COD_CARGO)
           itmX.SubItems(1) = rs!Descripcion
           itmX.SubItems(2) = Format(rs!Monto, "Standard")
           itmX.SubItems(3) = IIf((rs!Tipo = "P"), "Porcentual", "Monto")
           itmX.SubItems(4) = Format(rs!Valor, "Standard")
           curMonto = curMonto + rs!Monto
       rs.MoveNext
     Loop
     rs.Close

     Set itmX = lsw.ListItems.Add(, , "")
         itmX.SubItems(2) = "________"
     Set itmX = lsw.ListItems.Add(, , "")
         itmX.SubItems(2) = Format(curMonto, "Standard")

 Case "RSM"
   Call sbResumenOperacion

End Select


End Sub

Private Sub sbResumenOperacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMonto As Currency, curMontoTemp As Currency
Dim itmX As ListViewItem

'Inicializa
lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

lsw.ColumnHeaders.Add , , "", 2480
lsw.ColumnHeaders.Add , , "", 1360, vbRightJustify
lsw.ColumnHeaders.Add , , "", 1520
   
   Set itmX = lsw.ListItems.Add(, , "-> Monto Aprobado")
       itmX.SubItems(1) = Format(CCur(txtMonto.Text), "Standard")
       curMonto = CCur(txtMonto.Text)
       itmX.Bold = True

strSQL = "Select isnull(dbo.fxCxC_CuentaRebajos(" & txtOperacion.Text & ",'CRD'),0) as 'Crd'" _
       & ", isnull(dbo.fxCxC_CuentaRebajos(" & txtOperacion.Text & ",'CxC'),0) as 'CxC'" _
       & ", isnull(dbo.fxCxC_CuentaRebajos(" & txtOperacion.Text & ",'CAR'),0) as 'Car'" _
       & ", isnull(dbo.fxCxC_CuentaRebajos(" & txtOperacion.Text & ",'ADL'),0) as 'Adl'" _
       & ", isnull(dbo.fxCxC_CuentaIngresos(" & txtOperacion.Text & "),0) as 'Ing'"
Call OpenRecordSet(rs, strSQL)

   Set itmX = lsw.ListItems.Add(, , "(+) Otros Ingresos")
       itmX.SubItems(1) = Format(rs!Ing, "Standard")
       curMonto = curMonto - CCur(itmX.SubItems(1))
   
   Set itmX = lsw.ListItems.Add(, , "(-) Abonos a Créditos")
       itmX.SubItems(1) = Format(rs!Crd, "Standard")
       itmX.ForeColor = vbRed
       curMonto = curMonto - CCur(itmX.SubItems(1))

   Set itmX = lsw.ListItems.Add(, , "(-) Abonos a CxC Pendientes")
       itmX.SubItems(1) = Format(rs!CxC, "Standard")
       itmX.ForeColor = vbRed
       curMonto = curMonto - CCur(itmX.SubItems(1))

   Set itmX = lsw.ListItems.Add(, , "(-) Cargos Registrados")
       itmX.SubItems(1) = Format(rs!Car, "Standard")
       itmX.ForeColor = vbRed
       curMonto = curMonto - CCur(itmX.SubItems(1))

   Set itmX = lsw.ListItems.Add(, , "(-) Adelantos")
       itmX.SubItems(1) = Format(rs!Adl, "Standard")
       itmX.ForeColor = vbRed
       curMonto = curMonto - CCur(itmX.SubItems(1))

  curMonto = CCur(txtMonto.Text) + rs!Ing - (rs!Crd + rs!CxC + rs!Car + rs!Adl)

   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(1) = "___________"

   Set itmX = lsw.ListItems.Add(, , "Monto a Desembolsar")
       itmX.SubItems(1) = Format(curMonto, "Standard")
       itmX.Bold = True

rs.Close

End Sub

Private Sub lswOpciones_DblClick()
Dim strSQL As String, rs As New ADODB.Recordset

'Si la Operacion esta formalizada o Anulada entonces no ingresar a mantenimiento
If Not lsw.Enabled Then Exit Sub


'No permite ingresar a mantenimiento de las operaciones si no se encuentra en tramite.
Select Case Mid(cboEstado.Text, 1, 1)
   Case "R", "P"
     'Nada
   Case Else
      Exit Sub
End Select

If Not tlbPrincipal.Buttons.Item(1).Enabled Then
  MsgBox "Su usuario no no está autorizado!", vbInformation
  Exit Sub
End If

GLOBALES.gTag = txtOperacion.Text

Select Case lswOpciones.SelectedItem.Key
 Case "CRD"
      Call sbFormsCall("frmCxC_CuentasSGTRebajoCRD", 1, , , False, Me)

 Case "CXC"
      Call sbFormsCall("frmCxC_CuentasSGTRebajosInternos", 1, , , False, Me)

 Case "CAR"
   Operacion.Ventana = "C"
   Call sbFormsCall("frmCxC_CuentasSGTCargos", 1, , , False, Me)

 Case "ING"
   Operacion.Ventana = "C"
   Call sbFormsCall("frmCxC_CuentasSGTIngresos", 1, , , False, Me)
 

 Case "RSM" 'Nada
End Select

'Refresca datos en Pantalla
Call lswOpciones_Click

End Sub



Private Sub sbFacturas_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotal As Currency, curAdelanto As Currency

On Error GoTo vError

tcFacturas.Item(0).Selected = True

With lswFacturas
    .ListItems.Clear
    
    curTotal = 0
    curAdelanto = 0
    strSQL = "exec spCxC_Operacion_Facturas " & txtOperacion.Text & ",0"
    Call OpenRecordSet(rs, strSQL)
    
    Do While Not rs.EOF
      Set itmX = .ListItems.Add(, , rs!cod_Factura)
          itmX.SubItems(1) = Trim(rs!ESTADO_DESC)
          itmX.SubItems(2) = Trim(rs!Divisa_Desc)
          itmX.SubItems(3) = Format(rs!Importe, "Standard")
          itmX.SubItems(4) = rs!TIPO_CAMBIO
          itmX.SubItems(5) = Format(rs!Monto, "Standard")
          itmX.SubItems(6) = Format(rs!Fecha_Emision, "yyyy/mm/dd")
          itmX.SubItems(7) = Format(rs!Fecha_Pago, "yyyy/mm/dd")
          itmX.SubItems(8) = IIf((rs!Adelanto_Indica = 1), "Sí", "No")
          itmX.SubItems(9) = IIf((rs!Adelanto_Tipo = "P"), "Porcentual", "Monto")
          itmX.SubItems(10) = Format(rs!Adelanto_Monto, "Standard")
          itmX.SubItems(11) = Format(rs!Pendiente, "Standard")
          itmX.SubItems(12) = Format(rs!LIBERADO, "Standard")
          itmX.SubItems(13) = IIf(IsNull(rs!Operacion_Origen), 0, rs!Operacion_Origen)
       
          dtpFacturaPago.Value = rs!Fecha_Pago
          curTotal = curTotal + rs!Monto
          curAdelanto = curAdelanto + rs!Adelanto_Monto
       rs.MoveNext
    Loop
    rs.Close

End With

feFactura.Text = ""
feFactura.Tag = ""
feFacturaImporte.Text = 0
dtpFacturaEmite.Value = vFecha
If lswFacturas.ListItems.Count = 0 Then
    dtpFacturaPago.Value = vFecha
End If
feAdelanto.Text = 0
cboAdelantoTipo.Text = "Porcentual"

txtFacturas_Total.Text = Format(curTotal, "Standard")
txtFacturas_Total.ToolTipText = "Adelanto: " & Format(curAdelanto, "Standard")

txtFacturas_Casos.Text = Format(lswFacturas.ListItems.Count, "###,##0")

Call cbofacturasDivisas_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbFacturas_Adelantadas_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotal As Currency, curAdelanto As Currency

On Error GoTo vError

tcFacturas.Item(1).Selected = True

curTotal = 0
curAdelanto = 0

With lswAdelantadas
    .ListItems.Clear
    
    strSQL = "exec spCxC_Facturas_Adelantadas_Pendientes '" & txtCedula.Text & "','" & txtPagadorCed.Text & "'"
    Call OpenRecordSet(rs, strSQL)
        
        
    vPaso = True
    Do While Not rs.EOF
      Set itmX = .ListItems.Add(, , rs!cod_Factura)
          itmX.SubItems(1) = Trim(rs!Operacion)
          itmX.SubItems(2) = Trim(rs!cod_Divisa)
          itmX.SubItems(3) = Format(rs!Importe, "Standard")
          itmX.SubItems(4) = rs!TIPO_CAMBIO
          itmX.SubItems(5) = Format(rs!Monto, "Standard")
          itmX.SubItems(6) = Format(rs!Fecha_Emision, "yyyy/mm/dd")
          itmX.SubItems(7) = Format(rs!Fecha_Pago, "yyyy/mm/dd")
          itmX.SubItems(8) = Format(rs!Adelanto_Monto, "Standard")
          itmX.SubItems(9) = Format(rs!Pendiente, "Standard")
       
       curTotal = curTotal + rs!Monto
       curAdelanto = curAdelanto + rs!Adelanto_Monto
       
       rs.MoveNext
    Loop
    rs.Close
    
    vPaso = False
    
End With

txtFacturas_Adelanto_Total.Text = Format(curTotal, "Standard")
txtFacturas_Adelanto_Total.ToolTipText = "Adelanto: " & Format(curAdelanto, "Standard")
txtFacturas_Adelanto_Casos.Text = Format(lswAdelantadas.ListItems.Count, "###,##0")

txtFacturas_Adelanto_Total_Sel.Text = Format(0, "Standard")
txtFacturas_Adelanto_Total_Sel.ToolTipText = "Adelanto: " & Format(0, "Standard")
txtFacturas_Adelanto_Casos_Sel.Text = Format(0, "###,##0")



Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswOpciones_ItemClick(ByVal Item As MSComctlLib.ListItem)
Call lswOpciones_Click
End Sub

Private Sub ssTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

On Error GoTo vError

Select Case Item.Index
  Case 1 'Facturas
  
    'Carga Estados de la Factura: Registro
    strSQL = "select rtrim(Fe.Factura_Estado) as 'IdX', rtrim(Fe.Descripcion) as 'ItmX'" _
           & " from cxc_facturas_Estados Fe inner join CXC_CONCEPTOS_FACTURA_ESTADO Fa on Fe.Factura_Estado = Fa.Factura_Estado" _
           & " and Fa.cod_concepto = '" & txtConceptoCod.Text & "'" _
           & " where Fe.Proceso in('Registro','Confirmación') and Fe.Activo = 1"
    Call sbCbo_Llena_New(cboFacturaEstado, strSQL, False, True)

    Call sbFacturas_Load
    
  Case 2 'Activacion
    Call sbInitOpciones
  Case 3 'Historial
'    ssTabHistorial.Tab = 0
'    Call sbHistorial
  Case Else
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcFacturas_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Facturas Registradas
    Call sbFacturas_Load
  Case 1 'Facturas Adelantadas Pendientes
    Call sbFacturas_Adelantadas_Load
End Select

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbCargaCombos

End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

Select Case Button.Key
  Case "Autorizacion"
     GLOBALES.gTag = txtOperacion.Text
     Call sbFormsCall("frmCxC_CuentasSGTAutorizacion", 1, , , False, Me)
     
  Case "Activar"
        If fxActivacionVerifica Then
           
            i = MsgBox("Esta seguro que desea >> Activar << esta Operación", vbYesNo)
            If i = vbYes Then
                Call sbActivar
            End If
            
        Else 'Falla Verificacion de Formalizacion
         MsgBox vMensaje, vbCritical
        End If
  
  Case "Anular"

        If fxAnulacionVerifica Then
           i = MsgBox("Esta seguro que desea >> Anular << esta Operación", vbYesNo)
           If i = vbYes Then
              Call sbAnular
           End If
        Else
          MsgBox vMensaje, vbCritical
        End If
End Select

Call sbConsulta


End Sub

Private Sub txtAutorizadoCed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAutorizadoNom.SetFocus


If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "PcP.CEDULA_AUTORIZADO"
   gBusquedas.Orden = "PcP.CEDULA_AUTORIZADO"
   gBusquedas.Consulta = "select PcP.CEDULA_AUTORIZADO,Per.nombre from CXC_PERSONAS_AUTORIZADOS PcP" _
                      & " inner join CxC_Personas Per on PcP.CEDULA_AUTORIZADO = Per.cedula"
   gBusquedas.Filtro = " and PcP.cedula = '" & txtCedula.Text & "'"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtAutorizadoCed.Text = gBusquedas.Resultado
      txtAutorizadoNom.Text = gBusquedas.Resultado2
   End If
End If

End Sub


Private Sub txtAutorizadoNom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn And txtMonto.Enabled Then txtMonto.SetFocus
End Sub

Private Sub txtConceptoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConceptoDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_Concepto"
   gBusquedas.Orden = "cod_Concepto"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select cod_Concepto as 'Concepto',Descripcion  from CxC_Conceptos"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtConceptoCod.Text = gBusquedas.Resultado
      txtConceptoDesc.Text = gBusquedas.Resultado2
      Call txtConceptoCod_LostFocus
   End If
End If

End Sub

Private Sub txtConceptoCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

txtMonto.Locked = False
gbFactoreo.Visible = False

strSQL = "select C.*,dbo.MyGetdate() as 'FechaServer' " _
       & ", isnull(C.PAGADOR_DEFAULT,'') as 'PagadorId', isnull(P.Nombre,'') as 'PagadorDesc'" _
       & " from CxC_Conceptos C left join CxC_Personas P on C.PAGADOR_DEFAULT = P.cedula" _
       & " where C.cod_Concepto = '" & txtConceptoCod.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtConceptoDesc.Text = rs!Descripcion
    
   If rs!Requiere_Contrato = 1 Then
      txtContratoCod.Enabled = True
      txtContratoDesc.Enabled = True
   Else
      txtContratoCod.Enabled = False
      txtContratoDesc.Enabled = False
   End If
   
'   dtpDocFechaEmite.Value = rs!FechaServer
'   dtpDocFechaPago.Value = rs!FechaServer
'   dtpDocFechaVence.Value = rs!FechaServer
   
'   If rs!Requiere_Documento = 1 Then
'        dtpDocFechaEmite.Enabled = True
'        dtpDocFechaPago.Enabled = True
'        dtpDocFechaVence.Enabled = True
'   Else
'        dtpDocFechaEmite.Enabled = False
'        dtpDocFechaPago.Enabled = False
'        dtpDocFechaVence.Enabled = False
'   End If
   
   If rs!Proceso_Descuento = 1 Then
   
      txtPlazo.Enabled = False
   
'      dtpDocFechaEmite.Enabled = True
'      dtpDocFechaPago.Enabled = True
'      dtpDocFechaVence.Enabled = True
'
      txtContratoCod.Enabled = True
      txtContratoDesc.Enabled = True
      
      txtPagadorCed.Enabled = True
      txtPagadorNom.Enabled = True
      
      If Len(txtPagadorCed.Text) = 0 Then
            txtPagadorCed.Text = rs!PagadorId
            txtPagadorNom.Text = rs!PagadorDesc
      End If
      
      txtAutorizadoCed.Enabled = True
      txtAutorizadoNom.Enabled = True
      
      txtMonto.Locked = True
      gbFactoreo.Visible = True
      
      btnActualizaPorcentaje.Enabled = True
      btnFacturas.Enabled = True
      
   Else
      
      txtPlazo.Enabled = True
      
      txtContratoCod.Enabled = False
      txtContratoDesc.Enabled = False
      
      txtPagadorCed.Enabled = False
      txtPagadorNom.Enabled = False
      
      txtAutorizadoCed.Enabled = False
      txtAutorizadoNom.Enabled = False
   End If
    
   If mAdelantoPermite Then
      txtAdelantoPorcentaje.Text = mAdelantoPorc * 100
      
     If mAdelantoModifica Then
        txtAdelantoPorcentaje.Locked = False
     Else
        txtAdelantoPorcentaje.Locked = True
     End If
   
     
     txtAdelantoDias.Text = mAdelantoComisionDias
     If mAdelantoComisionApl Then
        chkAdelantoComisionApl.Value = xtpChecked
        txtAdelantoComision.Text = mAdelantoComision * 100
     Else
        chkAdelantoComisionApl.Value = xtpUnchecked
        txtAdelantoComision.Text = 0
     End If
   
   Else
     txtAdelantoPorcentaje.Text = "0"
     txtAdelantoPorcentaje.Locked = True
   End If
    
   chkTesoreria.Value = rs!Genera_Desembolso
    
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtConceptoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtContratoCod.Enabled Then
'    txtContratoCod.SetFocus
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Orden = "Descripcion"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select cod_Concepto,Descripcion  from CxC_Conceptos"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtConceptoCod.Text = gBusquedas.Resultado
      txtConceptoDesc.Text = gBusquedas.Resultado2
      Call txtConceptoCod_LostFocus
   End If
End If
End Sub

Private Sub txtContratoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContratoDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cnt.cod_Contrato"
   gBusquedas.Orden = "Cnt.cod_Contrato"
   gBusquedas.Consulta = "Select Cnt.cod_Contrato,Cnt.Descripcion" _
                       & " from CxC_Personas_Contratos Con inner join CxC_Contratos Cnt on Con.Cod_Contrato = Cnt.cod_contrato"
   gBusquedas.Filtro = " and Con.cedula = '" & txtCedula.Text & "' and Con.cod_contrato in(select cod_contrato" _
                     & " from CxC_Conceptos_Contratos where cod_concepto = '" & txtConceptoCod.Text & "') and Con.Activo = 1"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtContratoCod.Text = gBusquedas.Resultado
      txtContratoDesc.Text = gBusquedas.Resultado2
      Call sbContratoDetalle(txtContratoCod.Text, txtCedula.Text)
   End If
End If

End Sub


Private Sub txtContratoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And (txtPagadorCed.Enabled) Then txtPagadorCed.SetFocus



If KeyCode = vbKeyF4 Then

   gBusquedas.Columna = "Cnt.Descripcion"
   gBusquedas.Orden = "Cnt.Descripcion"
   gBusquedas.Consulta = "Select Cnt.cod_Contrato,Cnt.Descripcion" _
                       & " from CxC_Personas_Contratos Con inner join CxC_Contratos Cnt on Con.Cod_Contrato = Cnt.cod_contrato"
   gBusquedas.Filtro = " and Con.cedula = '" & txtCedula.Text & "' and Con.cod_contrato in(select cod_contrato" _
                     & " from CxC_Conceptos_Contratos where cod_concepto = '" & txtConceptoCod.Text & "') and Con.Activo = 1"
  
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtContratoCod.Text = gBusquedas.Resultado
      txtContratoDesc.Text = gBusquedas.Resultado2
      Call sbContratoDetalle(txtContratoCod.Text, txtCedula.Text)
   End If
End If

End Sub

Private Sub cboCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtCuota_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocNum.SetFocus

End Sub


Private Sub txtDocNum_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And (dtpDocFechaEmite.Enabled) Then dtpDocFechaEmite.SetFocus
 
End Sub



Private Sub txtFacturaFiltro_Change(Index As Integer)

On Error GoTo vError

If Index = 0 Then
  Call FindListItem(txtFacturaFiltro.Item(Index).Text, lswFacturas, 0)
Else
  Call FindListItem(txtFacturaFiltro.Item(Index).Text, lswAdelantadas, 0)
End If

vError:

End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

vError:

End Sub


Private Sub sbCargosAdicionales(vOperacion As Long, vCodigo As String, vMonto As Currency)
Dim strSQL As String

'strSQL = "exec spCRDOperacionCargosAdd " & vOperacion & ",'" & vCodigo & "'," & vMonto & ""
'Call ConectionExecute(strSQL)

End Sub


Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtPlazo.Enabled Then
     txtPlazo.SetFocus
  Else
     txtTasa.SetFocus
  End If
End If
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto.Text = Format(txtMonto.Text, "Standard")

vError:

End Sub

Private Sub sbInitOpciones()
Dim itmX As ListItem
Dim i As Integer


lswOpciones.ListItems.Clear
lswOpciones.ColumnHeaders.Clear
lswOpciones.ColumnHeaders.Add , , , 3000

'For i = 1 To ImageList1.ListImages.Count
' lswOpciones.Icons.LoadBitmap ImageList1.ListImages(i).ExtractIcon.Handle, i, xtpImageNormal
'Next i

Set itmX = lswOpciones.ListItems.Add(1, "ING", "Registro de Ingesos", , 8)
Set itmX = lswOpciones.ListItems.Add(2, "CRD", "Abonos a Créditos", , 8)
Set itmX = lswOpciones.ListItems.Add(3, "CXC", "Abonos a CxC Pendientes", , 8)
Set itmX = lswOpciones.ListItems.Add(4, "CAR", "Registro de Cargos", , 8)
Set itmX = lswOpciones.ListItems.Add(5, "RSM", "Resumen", , 8)


lsw.ListItems.Clear


End Sub


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pPagadorCed As String, pAutorizadorCed As String, pContratoCod As String, pOperacion As Long
Dim pRebajos As Currency, pPlazoDias As Long, pDesembolsoPendiente As Currency
Dim pCuenta As String, pFechaInicio As String

On Error GoTo vError

'Cuenta Bancaria
If cboCuenta.ListCount > 0 Then
   If cboCuenta.Text <> "" Then
      pCuenta = cboCuenta.ItemData(cboCuenta.ListIndex)
   Else
      pCuenta = ""
   End If
Else
   pCuenta = ""
End If


'Datos de Requerimientos de Pagador y Contrato
If txtPagadorCed.Enabled Then
   pPagadorCed = "'" & txtPagadorCed.Text & "'"
Else
   pPagadorCed = "Null"
End If

If txtAutorizadoCed.Enabled Then
   pAutorizadorCed = "'" & txtAutorizadoCed.Text & "'"
Else
   pAutorizadorCed = "Null"
End If


If chkCtaApl.Value = xtpChecked Then
    pFechaInicio = "'" & Format(dtpInicioFecha.Value, "yyyy-mm-dd") & "'"
Else
    pFechaInicio = "Null"
End If


If txtContratoCod.Enabled Then
   pContratoCod = "'" & txtContratoCod.Text & "'"
Else
   pContratoCod = "Null"
End If
       
If Not vEdita Then
    strSQL = "select isnull(max(Operacion),0) + 1 as 'Operacion' from CxC_Cuentas"
    Call OpenRecordSet(rs, strSQL)
      pOperacion = rs!Operacion
    rs.Close
Else
   'Actualiza Registro de Cargos por Variación en el Monto o Tipo de Concepto
   If mMonto <> CCur(txtMonto.Text) Or mConcepto <> Trim(txtConceptoCod.Text) Then
      strSQL = "exec spCxC_CuentaCargosActualiza " & txtOperacion.Text & "," & CCur(txtMonto.Text)
      Call ConectionExecute(strSQL)
   End If
   
   'Extrae Rebajos Totales de la Operación
   strSQL = "Select isnull(dbo.fxCxC_CuentaRebajos(" & txtOperacion.Text & ",'TOT'),0) as 'Rebajos'"
   Call OpenRecordSet(rs, strSQL)
      pRebajos = rs!Rebajos
   rs.Close
End If

pPlazoDias = CLng(txtPlazo.Text) * 30
pDesembolsoPendiente = CCur(txtMonto.Text)


Select Case Mid(cboEstado.Text, 1, 1)
  Case "R", "P" 'Recepción
    If Not vEdita Then
       strSQL = "insert CxC_Cuentas(OPERACION,CEDULA,CEDULA_PAGADOR,COD_CONCEPTO,COD_OFICINA,NOTAS,MONTO,SALDO,REBAJOS_TOTAL" _
              & ",EMITIR_TIPO,EMITIR_BANCO,EMITIR_CUENTA,DESEMBOLSO_MONTO,TIPO_PLAZO,TASA_CORRIENTE,TASA_MORA,CUOTA,DIAS_PLAZO, PLAZO,AMORTIZA,INTERESC" _
              & ",ESTADO,NUM_DOCUMENTO,COD_CONTRATO,REGISTRO_FECHA,REGISTRO_USUARIO,FECHA_ULTMOV,AUTORIZA_ESTADO" _
              & ", ADELANTO_MONTO, ADELANTO_PORCENTAJE, DESEMBOLSO_REALIZADO, DESEMBOLSO_PENDIENTE, CEDULA_AUTORIZADO" _
              & ", ADELANTO_COMISION_APL, ADELANTO_COMISION, ADELANTO_COMISION_DIAS, FREQ_PAGO, FECHA_INICIO)" _
              & " VALUES(" & pOperacion & ",'" & txtCedula.Text & "'," & pPagadorCed & ",'" & txtConceptoCod.Text & "','" & GLOBALES.gOficinaTitular _
              & "','" & txtNotas.Text & "'," & CCur(txtMonto.Text) & "," & CCur(txtMonto.Text) & ",0,'" & fxTipoDocumento(cboTipoDocumento.Text) _
              & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & pCuenta & "'," & CCur(txtMonto.Text) & ",'M'," & CCur(txtTasa.Text) _
              & "," & CCur(txtTasaMora.Text) & "," & CCur(txtCuota.Text) & "," & pPlazoDias & "," & txtPlazo.Text & ",0,0,'R','" & txtDocNum.Text _
              & "'," & pContratoCod & ",dbo.MyGetdate(),'" & glogon.Usuario & "',dbo.MyGetdate(),'P'" _
              & "," & CCur(txtAdelantoTotal.Text) & "," & CCur(txtAdelantoPorcentaje.Text) & ",0,0," & pAutorizadorCed _
              & "," & chkAdelantoComisionApl.Value & "," & CCur(txtAdelantoComision.Text) & "," & CLng(txtAdelantoDias.Text) _
              & ", " & IIf(chkCtaApl.Value, 30, 0) & ", " & pFechaInicio & ")"
              
       Call ConectionExecute(strSQL)
              
      strSQL = "exec spCxC_CuentaCargosActualiza " & pOperacion & "," & CCur(txtMonto.Text)
      Call ConectionExecute(strSQL)
      
      txtOperacion.Text = pOperacion
              
              
              
    Else
     If mAutoriza = "P" Then
       strSQL = "update CxC_Cuentas set cedula_pagador = " & pPagadorCed & ",CEDULA_AUTORIZADO = " & pAutorizadorCed _
              & ",cod_concepto = '" & txtConceptoCod.Text & "',cod_oficina = '" _
              & cboOficina.ItemData(cboOficina.ListIndex) & "',notas = '" & txtNotas.Text & "',Monto = " & CCur(txtMonto.Text) & ", Saldo = " & CCur(txtMonto.Text) _
              & ",Rebajos_Total = " & pRebajos & ", emitir_tipo = '" & fxTipoDocumento(cboTipoDocumento.Text) & "', Emitir_Banco = " _
              & cboBanco.ItemData(cboBanco.ListIndex) & ",Emitir_Cuenta = '" & pCuenta & "', Tasa_Corriente = " & CCur(txtTasa.Text) _
              & ", Tasa_Mora = " & CCur(txtTasaMora.Text) & ", Cuota =  " & CCur(txtCuota.Text) & ",cod_contrato = " & pContratoCod _
              & ", Estado = '" & Mid(cboEstado.Text, 1, 1) & "',num_documento = '" & txtDocNum.Text _
              & "',desembolso_Monto = " & CCur(txtMonto.Text) - pRebajos & ", Dias_Plazo = " & pPlazoDias & ", Plazo = " & txtPlazo.Text _
              & ", ADELANTO_MONTO = " & CCur(txtAdelantoTotal.Text) & ",ADELANTO_PORCENTAJE = " & CCur(txtAdelantoPorcentaje.Text) _
              & ", ADELANTO_COMISION_APL = " & chkAdelantoComisionApl.Value & ", ADELANTO_COMISION = " & CCur(txtAdelantoComision.Text) _
              & ", ADELANTO_COMISION_DIAS = " & CLng(txtAdelantoDias.Text) _
              & ", FREQ_PAGO = " & IIf(chkCtaApl.Value, 30, 0) & ", FECHA_INICIO = " & pFechaInicio _
              & " where Operacion = " & txtOperacion.Text
        Call ConectionExecute(strSQL)
      
      
        If CLng(txtFacturasNo.Text) > 0 Then
            'Para Factoreo: Este procedimiento Lleva Implicito el de Actualización de Cargos
            strSQL = "exec spCxC_Operacion_Facturas_Actualiza " & pOperacion & ",0,'" & glogon.Usuario & "'"
            Call ConectionExecute(strSQL)
        Else
            strSQL = "exec spCxC_CuentaCargosActualiza " & pOperacion & "," & CCur(txtMonto.Text)
            Call ConectionExecute(strSQL)
        End If
      
      Else
       strSQL = "update CxC_Cuentas set notas = '" & txtNotas.Text & "', emitir_tipo = '" & fxTipoDocumento(cboTipoDocumento.Text) & "', Emitir_Banco = " _
              & cboBanco.ItemData(cboBanco.ListIndex) & ",Emitir_Cuenta = '" & pCuenta _
              & "', cedula_pagador = " & pPagadorCed & " where Operacion = " & txtOperacion.Text
           Call ConectionExecute(strSQL)
  End If
    
    
    End If
  
  Case "A" 'Activa
      If gbFactoreo.Visible Then
        strSQL = "update CxC_Cuentas set notas = '" & txtNotas.Text & "', emitir_tipo = '" & fxTipoDocumento(cboTipoDocumento.Text) & "', Emitir_Banco = " _
               & cboBanco.ItemData(cboBanco.ListIndex) & ",Emitir_Cuenta = '" & pCuenta _
               & "', Desembolso_Pendiente = ADELANTO_MONTO, cedula_pagador = " & pPagadorCed _
               & " where Operacion = " & txtOperacion.Text & " and Tesoreria_Fecha is null"
      Else
        strSQL = "update CxC_Cuentas set notas = '" & txtNotas.Text & "', emitir_tipo = '" & fxTipoDocumento(cboTipoDocumento.Text) & "', Emitir_Banco = " _
               & cboBanco.ItemData(cboBanco.ListIndex) & ",Emitir_Cuenta = '" & pCuenta _
               & "', Desembolso_Pendiente = Desembolso_Monto, cedula_pagador = " & pPagadorCed _
               & " where Operacion = " & txtOperacion.Text & " and Tesoreria_Fecha is null"
      End If
      Call ConectionExecute(strSQL)
      

      
      
End Select

Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key
 Case "nuevo"
  txtOperacion.Text = ""
  txtOperacion.Enabled = False
  
  Call sbLimpiaDatos
  tlbPrincipal.Buttons(1).Enabled = False
  tlbPrincipal.Buttons(2).Enabled = False
  tlbPrincipal.Buttons(3).Enabled = True
  tlbPrincipal.Buttons(4).Enabled = True
  fraOperacion.Enabled = True
  txtNotas.Locked = False
  
  vEdita = False
  
  
  txtCedula.SetFocus
'  Call sbCargaCombos
  
  
  
 Case "editar"
  If CLng(txtOperacion.Text) > 0 Then  'And Operacion.Estado = "A" Then
      vEdita = True
'      Call Edicion(1)
    
      'Si el Estado Esta en Recepcion o Resolucion puede Cambiar Todos Los Datos
      'Si Esta en Formalización Solo puede Cambiar la Salida
      tlbPrincipal.Buttons(1).Enabled = False
      tlbPrincipal.Buttons(2).Enabled = False
      tlbPrincipal.Buttons(3).Enabled = True
      tlbPrincipal.Buttons(4).Enabled = True
      
      txtOperacion.Enabled = False
      fraOperacion.Enabled = True
      txtNotas.Locked = False
      txtCedula.SetFocus

  End If
 
 Case "guardar"
  
  If fxVerificaRecepcion Then
    Call sbGuardar
    Call sbConsulta
    txtOperacion.Enabled = True
    
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    
    fraOperacion.Enabled = False
    txtNotas.Locked = True
  End If
 
 Case "deshacer"
    txtOperacion.Enabled = True
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    fraOperacion.Enabled = False
    If txtOperacion <> "" Then Call sbConsulta
    txtOperacion.SetFocus
 
 Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
 
 Case "cerrar"
    Unload Me

End Select


End Sub



Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

If Not IsNumeric(txtOperacion.Text) Then Exit Sub

Me.MousePointer = vbHourglass

Select Case ButtonMenu.Key
  Case "Boleta"
      Call sbReporte_Boleta
  Case "Nota"
      Call sbImprimeRecibo(txtOperacion.Text, "CxC_FRM")
  Case "Cesion"
      Call sbReporte_Cesion
        
End Select

Me.MousePointer = vbDefault

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre, Categoria from vCxC_Personas_Filtro"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If


End Sub

Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCedula.Text = "" Then Exit Sub

Me.MousePointer = vbHourglass


Call cboBanco_Click

strSQL = "select Nombre, Adelanto_Permite,Adelanto_Porcentaje, Adelanto_Modifica, Credito_Limite, Credito_Cerrado" _
       & ", ADELANTO_COMISION, ADELANTO_COMISION_APL" _
       & " from cxc_personas where cedula = '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)

 mAdelantoPermite = IIf((rs!ADELANTO_PERMITE = 1), True, False)
 mAdelantoModifica = IIf((rs!ADELANTO_MODIFICA = 1), True, False)
 mAdelantoPorc = CCur(rs!ADELANTO_PORCENTAJE) / 100
 mAdelantoComision = CCur(rs!ADELANTO_COMISION) / 100
 mAdelantoComisionApl = IIf((rs!ADELANTO_COMISION_APL = 1), True, False)
 
 mCreditoLimite = rs!Credito_Limite
 mCreditoCerrado = IIf((rs!Credito_Cerrado = 1), True, False)
 
txtNombre = rs!Nombre
lblNombre.Caption = txtNombre.Text
 
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtPagadorCed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtPagadorNom.Enabled Then txtPagadorNom.SetFocus


If KeyCode = vbKeyF4 Then
   
  If mCntPagadorAbierto Then
        gBusquedas.Columna = "Cedula"
        gBusquedas.Orden = "Cedula"
        gBusquedas.Consulta = "select Cedula,Nombre" _
                           & " from CxC_Personas"
        gBusquedas.Filtro = " and Rol_Pagador = 1"
  Else
        gBusquedas.Columna = "PcP.Cedula_Pagador"
        gBusquedas.Orden = "PcP.Cedula_Pagador"
        gBusquedas.Consulta = "select PcP.Cedula_Pagador,Per.nombre" _
                           & " from CxC_Personas_Contratos_Pagadores PcP" _
                           & " inner join CxC_Personas Per on PcP.cedula_pagador = Per.cedula"
        gBusquedas.Filtro = " and PcP.cod_contrato = '" & txtContratoCod.Text & "' and PcP.cedula = '" & txtCedula.Text & "'"
   End If
   
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtPagadorCed.Text = gBusquedas.Resultado
      txtPagadorNom.Text = gBusquedas.Resultado2
   End If
End If

End Sub
Private Sub txtPagadorNom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtAutorizadoCed.Enabled Then txtAutorizadoCed.SetFocus
End Sub

Private Sub txtTasa_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
  txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa))
End If
vError:
End Sub

Private Sub txtMonto_Change()
On Error GoTo vError

If CCur(IIf((txtTasa = ""), 0, txtTasa)) > 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
 txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa))
End If
vError:

End Sub

Private Sub txtTasa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoDocumento.SetFocus
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
  On Error Resume Next
    If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
        And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
      txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa))
    End If
End If
End Sub

Private Sub txtTasa_LostFocus()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
  txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa))
End If
vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConceptoCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from CxC_Personas"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If

End Sub



Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion = xOpTemp
 Call sbConsulta
End Sub


Private Sub txtOperacion_Change()
 Call sbLimpiaDatos
  With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbConsulta
'If KeyCode = vbKeyF4 Then Call sbBusqueda(0)
End Sub

Private Sub txtPlazo_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
  txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa))
End If

vError:
End Sub


Private Sub txtPlazo_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtTasa.SetFocus
End Sub



