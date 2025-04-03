VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCxPControlReprogramacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprogramación de Pagos de Facturas"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   9705
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   9495
      _Version        =   1441793
      _ExtentX        =   16748
      _ExtentY        =   8281
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
      Item(0).Caption =   "Paso 1"
      Item(0).ControlCount=   25
      Item(0).Control(0)=   "txtCodigo"
      Item(0).Control(1)=   "txtNombre"
      Item(0).Control(2)=   "txtFactura"
      Item(0).Control(3)=   "txtMontoAjuste"
      Item(0).Control(4)=   "Label1(5)"
      Item(0).Control(5)=   "Label2(15)"
      Item(0).Control(6)=   "lblVencimiento"
      Item(0).Control(7)=   "Label2(3)"
      Item(0).Control(8)=   "Label2(4)"
      Item(0).Control(9)=   "lblSaldoDivReal"
      Item(0).Control(10)=   "Label2(8)"
      Item(0).Control(11)=   "lblFacturaSaldo"
      Item(0).Control(12)=   "Label2(6)"
      Item(0).Control(13)=   "lblFacturaMonto"
      Item(0).Control(14)=   "Label2(14)"
      Item(0).Control(15)=   "Label2(11)"
      Item(0).Control(16)=   "Label2(12)"
      Item(0).Control(17)=   "Label2(13)"
      Item(0).Control(18)=   "lblImporteReal"
      Item(0).Control(19)=   "lblDivisa"
      Item(0).Control(20)=   "lblTipoCambio"
      Item(0).Control(21)=   "Label1(0)"
      Item(0).Control(22)=   "Label1(1)"
      Item(0).Control(23)=   "btnAjustar"
      Item(0).Control(24)=   "ShortcutCaption1"
      Item(1).Caption =   "Paso 2"
      Item(1).ControlCount=   8
      Item(1).Control(0)=   "Frame1"
      Item(1).Control(1)=   "vGridCargos"
      Item(1).Control(2)=   "txtSaldo"
      Item(1).Control(3)=   "txtPagosTrans"
      Item(1).Control(4)=   "chkAplCarPerPorc"
      Item(1).Control(5)=   "chkAplCarPerMonto"
      Item(1).Control(6)=   "Label1(3)"
      Item(1).Control(7)=   "Label1(2)"
      Item(2).Caption =   "Re Programación"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "vGrid"
      Item(2).Control(1)=   "cmdAplicar"
      Item(2).Control(2)=   "txtPendiente"
      Item(2).Control(3)=   "Label1(4)"
      Begin XtremeSuiteControls.PushButton btnAjustar 
         Height          =   315
         Left            =   8400
         TabIndex        =   47
         Top             =   1680
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Ajustar!"
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
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Pagos"
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
         Height          =   2412
         Left            =   -69760
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   3855
         Begin XtremeSuiteControls.DateTimePicker dtpProgPrimerPago 
            Height          =   372
            Left            =   1800
            TabIndex        =   28
            Top             =   1080
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   656
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
         Begin XtremeSuiteControls.FlatEdit txtProgNPagos 
            Height          =   312
            Left            =   1800
            TabIndex        =   29
            Top             =   360
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtProgFrecuencia 
            Height          =   312
            Left            =   1800
            TabIndex        =   30
            Top             =   720
            Width           =   1332
            _Version        =   1441793
            _ExtentX        =   2350
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkProgImpPrimPago 
            Height          =   372
            Left            =   120
            TabIndex        =   31
            Top             =   1560
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cancelar el I.V. en el 1 er. Pago"
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
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkProCargos 
            Height          =   372
            Left            =   120
            TabIndex        =   32
            Top             =   1920
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cargos entre el No. de Pagos"
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
            Appearance      =   16
            Value           =   1
            Alignment       =   1
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Pagos"
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
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha 1er. Pago"
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
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1692
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Frecuencia (Días)"
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
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   1692
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1926
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
         Width           =   6735
         _Version        =   1441793
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFactura 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4466
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMontoAjuste 
         Height          =   315
         Left            =   6360
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGridCargos 
         Height          =   3735
         Left            =   -65440
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   4575
         _Version        =   524288
         _ExtentX        =   8070
         _ExtentY        =   6588
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
         MaxCols         =   484
         ScrollBars      =   2
         SpreadDesigner  =   "frmCxPControlReprogramacion.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   315
         Left            =   -67840
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPagosTrans 
         Height          =   315
         Left            =   -67840
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
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
      Begin XtremeSuiteControls.CheckBox chkAplCarPerPorc 
         Height          =   375
         Left            =   -69640
         TabIndex        =   39
         Top             =   3600
         Visible         =   0   'False
         Width           =   3735
         _Version        =   1441793
         _ExtentX        =   6583
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aplica Cargos Periodicos Porcentuales"
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
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkAplCarPerMonto 
         Height          =   375
         Left            =   -69640
         TabIndex        =   40
         Top             =   3960
         Visible         =   0   'False
         Width           =   3735
         _Version        =   1441793
         _ExtentX        =   6583
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aplica Cargos Periodicos por Monto"
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
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3015
         Left            =   -69880
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   9255
         _Version        =   524288
         _ExtentX        =   16325
         _ExtentY        =   5318
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
         MaxCols         =   484
         SpreadDesigner  =   "frmCxPControlReprogramacion.frx":059E
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   615
         Left            =   -63880
         TabIndex        =   44
         Top             =   3720
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Re-Programar"
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
         Picture         =   "frmCxPControlReprogramacion.frx":0DE8
      End
      Begin XtremeSuiteControls.FlatEdit txtPendiente 
         Height          =   315
         Left            =   -66640
         TabIndex        =   45
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   48
         Top             =   360
         Width           =   9495
         _Version        =   1441793
         _ExtentX        =   16748
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Indique el Proveedor y el Número de Factura (Programada) "
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pendiente de Distribuir"
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
         Left            =   -68920
         TabIndex        =   46
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   2
         Left            =   -69640
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Pagos Transcurridos"
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
         Left            =   -69640
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "No. Factura"
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
         Left            =   360
         TabIndex        =   26
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
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
         Index           =   0
         Left            =   360
         TabIndex        =   25
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblTipoCambio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         Left            =   2040
         TabIndex        =   24
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label lblDivisa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         Left            =   2040
         TabIndex        =   23
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label lblImporteReal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         Left            =   2040
         TabIndex        =   22
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Cambio"
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
         Index           =   13
         Left            =   600
         TabIndex        =   21
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   12
         Left            =   600
         TabIndex        =   20
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Importe Real"
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
         Left            =   600
         TabIndex        =   19
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Divisa Local:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   14
         Left            =   4680
         TabIndex        =   18
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblFacturaMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         Left            =   6360
         TabIndex        =   17
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   6
         Left            =   5160
         TabIndex        =   16
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label lblFacturaSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         Left            =   6360
         TabIndex        =   15
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   14
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblSaldoDivReal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         Left            =   2040
         TabIndex        =   13
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   12
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Divisa Real:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblVencimiento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         Left            =   6360
         TabIndex        =   10
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   15
         Left            =   5160
         TabIndex        =   9
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Corrección de Monto ?"
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
         Left            =   4320
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
   End
   Begin XtremeSuiteControls.PushButton cmdAtras 
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Atras"
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
      Picture         =   "frmCxPControlReprogramacion.frx":14D0
   End
   Begin XtremeSuiteControls.PushButton cmdSiguiente 
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   6120
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Adelante"
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
      Picture         =   "frmCxPControlReprogramacion.frx":1ABD
      TextImageRelation=   4
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Re.Programación de Pagos de Facturas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   852
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   7572
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmCxPControlReprogramacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnAjustar_Click()
Dim strSQL As String, curAjuste As Currency, vMensaje As String

On Error GoTo vError

vMensaje = ""

If CCur(lblFacturaSaldo.Caption) <> CCur(lblFacturaMonto.Caption) Then
   vMensaje = vMensaje & " - No se puede corregir monto de la factura porque la factura ya se le han realizado movimientos!" & vbCrLf
End If


If Abs(CCur(txtMontoAjuste.Text) - CCur(lblFacturaMonto.Caption)) > 10 Then
   vMensaje = vMensaje & " - El Monto de Ajuste de la Factura no puede ser superior a 10!" & vbCrLf
End If

If Len(vMensaje) > 0 Then
   MsgBox vMensaje, vbExclamation, "Verificación de Datos!"
   Exit Sub
End If

Me.MousePointer = vbHourglass

curAjuste = CCur(txtMontoAjuste.Text) - CCur(lblFacturaMonto.Caption)

strSQL = "exec spCxP_AjusteMontoFactura " & txtCodigo.Text & ",'" & txtFactura.Text & "'," & curAjuste
Call ConectionExecute(strSQL)

vMensaje = "Ajuste Monto Factura: " & txtFactura.Text & " [Prov." & txtCodigo.Text & "] Mnt.Ant.:" & lblFacturaMonto.Caption & " -> Mnt.Nv.:" & txtMontoAjuste.Text

Call Bitacora("Modifica", vMensaje)

Me.MousePointer = vbDefault
MsgBox "Ajuste de Monto de Factura Realizado Correctamente!", vbInformation

Call sbInicialConsulta

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curCargos As Currency, i As Integer, y As Integer
Dim vFecha As Date, vMonto As Currency, vPago As Integer
Dim vImporteReal As Currency

On Error GoTo vError

If vGrid.MaxRows <= 0 Then Exit Sub

If Not fxVerificaDatos Then Exit Sub

vGrid.Row = 1
vGrid.col = 1
vPago = vGrid.Text
vPago = vPago

glogon.Conection.BeginTrans

'Limpia Cargos
strSQL = "delete cxp_pagoProvCargos where nPago >= " & vPago _
       & " and cod_proveedor = " & txtCodigo & " and cod_factura ='" & txtFactura & "'"
Call ConectionExecute(strSQL)

'Limpia Pagos
strSQL = "delete cxp_pagoProv where nPago >= " & vPago _
       & " and cod_proveedor = " & txtCodigo & " and cod_factura ='" & txtFactura & "'"
Call ConectionExecute(strSQL)


With vGridCargos
 curCargos = 0
 For i = 1 To .MaxRows
   .col = 3
   .Row = i
   curCargos = curCargos + CCur(.Text)
 Next i
End With


'Generar Pagos

Me.MousePointer = vbHourglass

With vGrid
 For i = 1 To .MaxRows
     .Row = i
     .col = 1
     vPago = .Text
     .col = 5
     vMonto = .Text
     .col = 8
     vFecha = .Text
     .col = 9
'     vImporteReal = .Text
     vImporteReal = vMonto / CCur(lblTipoCambio.Caption)
     
     strSQL = "insert cxp_pagoProv(npago,cod_proveedor,cod_factura,fecha_vencimiento,monto" _
            & ",frecuencia,tipo_transac,apl_cargo_flotante,pago_anticipado,forma_pago" _
            & ",importe_divisa_real,tipo_cambio,cod_divisa) values(" & vPago _
            & "," & txtCodigo & ",'" & txtFactura & "','" & Format(vFecha, "yyyy/mm/dd") & "'," & vMonto _
            & "," & txtProgFrecuencia & "," & IIf((CInt(txtProgNPagos) = 1), 0, 1) _
            & "," & chkAplCarPerMonto.Value & ",0,'CR'," & vImporteReal & "," & CCur(lblTipoCambio.Caption) _
            & ",'" & Trim(lblDivisa.Caption) & "')"
     Call ConectionExecute(strSQL)
 
 Next i
End With


vGrid.Row = 1
vGrid.col = 1
vPago = vGrid.Text
vPago = vPago - 1 'Suma + 1, en el Ciclo

'Procedimiento 1 de la Aplicación de Cargos es Excluyente con el 2
If chkProCargos.Value = vbChecked And curCargos > 0 Then
 With vGridCargos
   For y = 1 To .MaxRows
      .col = 3
      .Row = y
      curCargos = CCur(.Text) / CInt(txtProgNPagos)
      If curCargos > 0 Then
        For i = 1 To CInt(txtProgNPagos)
           .col = 1
           strSQL = "insert cxp_PagoProvCargos(Npago,Cod_factura,cod_proveedor,cod_cargo,monto,registro_fecha,registro_usuario" _
                  & ",cod_divisa,tipo_cambio,tipo_cargo,tipo_proceso)" _
                  & " values(" & i + vPago & ",'" & txtFactura & "'," & txtCodigo _
                  & ",'" & Trim(.Text) & "'," & curCargos & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & lblDivisa.Caption _
                  & "'," & CCur(lblTipoCambio.Caption) & ",'M','D')"
           Call ConectionExecute(strSQL)
        
        Next i
      End If
   Next y
 End With
End If 'Chk y CurCargos > 0


glogon.Conection.CommitTrans

'Revisar si el Pago, se cancela con el cargo asignado. Dejando el Monto del Pago en Cero
'En este caso desactivar de Envio a Tesoreria. Y por ende de la antiguedad de saldos y programacion de pagos

strSQL = "update cxp_pagoProv set Tesoreria = 0,fecha_traslada = dbo.MyGetdate(),user_traslada = '" & glogon.Usuario _
       & "' where cod_proveedor = " & txtCodigo & " and cod_factura = '" & txtFactura _
       & "' and Npago in(select P.npago from cxp_pagoProv P inner join cxp_PagoprovCargos C" _
       & " on P.cod_proveedor = C.cod_proveedor and P.cod_factura = C.cod_factura and P.npago = C.npago" _
       & " where P.cod_proveedor = " & txtCodigo & " and P.cod_factura = '" & txtFactura _
       & "' group by P.npago,P.cod_proveedor,P.cod_factura,P.monto Having P.Monto = isnull(Sum(C.Monto), 0))"
Call ConectionExecute(strSQL)

'TODO: Revisar Afectación de Saldos?

Me.MousePointer = vbDefault
MsgBox "Generación de Cola de Pagos Realizada Satisfactoriamente...", vbInformation

'ssTab.Tab = 0

tcMain.Item(0).Selected = True
Call sbLimpiaDatos

Exit Sub

vError:
  Me.MousePointer = vbDefault
  glogon.Conection.RollbackTrans
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtCodigo_Change()
Call sbInicialLimpia
End Sub

Private Sub txtCodigo_LostFocus()
txtNombre = fxSIFCCodigos("D", txtCodigo, "proveedores")
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtFactura_Change()
Call sbInicialLimpia
End Sub

Private Sub txtMontoAjuste_GotFocus()
On Error GoTo vError

  txtMontoAjuste.Text = CCur(txtMontoAjuste.Text)

vError:
End Sub

Private Sub txtMontoAjuste_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtMontoAjuste.Text = Format(CCur(txtMontoAjuste.Text), "Standard")
End If

vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFactura.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub sbInicialLimpia()

lblDivisa.Caption = ""
lblFacturaMonto.Caption = "0"
lblFacturaSaldo.Caption = "0"
lblImporteReal.Caption = "0"
lblTipoCambio.Caption = "0"
lblSaldoDivReal.Caption = "0"
lblVencimiento.Caption = ""
txtMontoAjuste.Text = "0"

End Sub

Private Sub sbInicialConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbInicialLimpia

strSQL = "select * from vCxP_ProgramacionPago where Cxp_Estado = 'G' and cod_proveedor = " & txtCodigo.Text _
       & " and cod_factura = '" & txtFactura.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  lblDivisa.Caption = rs!cod_Divisa
  lblTipoCambio.Caption = rs!TIPO_CAMBIO
  lblImporteReal.Caption = Format(rs!Importe_divisa_real, "Standard")
  lblFacturaMonto.Caption = Format(rs!Total, "Standard")
  lblVencimiento.Caption = Format(rs!Vence, "dd/mm/yyyy")
  
  txtMontoAjuste.Text = lblFacturaMonto.Caption
  
  rs.Close
    strSQL = "select isnull(sum(Monto),0) as Monto, isnull(sum(IMPORTE_DIVISA_REAL),0) as ImporteReal" _
           & " From CxP_PagoPRov Where Tesoreria Is Null" _
           & " and cod_proveedor = " & txtCodigo & " and Cod_FActura = '" & txtFactura & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      lblFacturaSaldo.Caption = Format(rs!Monto, "Standard")
      lblSaldoDivReal.Caption = Format(rs!ImporteReal, "Standard")
    End If
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbInicialConsulta

On Error GoTo vError

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_factura"
  gBusquedas.Orden = "cod_factura"
  gBusquedas.Consulta = "select cod_factura,total,fecha From vCxP_ProgramacionPago"
  gBusquedas.Filtro = " and CxP_Estado = 'G' and Cod_Proveedor = " & txtCodigo
  frmBusquedas.Show vbModal
  txtFactura.Text = gBusquedas.Resultado
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Load()

vModulo = 30

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle
vGridCargos.AppearanceStyle = fxGridStyle


tcMain.Item(0).Selected = True

Call sbLimpiaDatos

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbReCalculaGrid()
Dim y As Integer, curTotal As Currency
Dim curMonto As Currency, curCargos As Currency

curTotal = 0

With vGrid
   For y = 1 To .MaxRows
      .Row = y
      .col = 4
      curCargos = vGrid.Text
      .col = 5
      curMonto = vGrid.Text
      curTotal = curTotal + curMonto
      .col = 6
      vGrid.Text = CStr(curMonto - curCargos)
  
   Next y
End With

txtPendiente = Format((CCur(txtSaldo) - curTotal), "Standard")

If Abs(txtPendiente) < 1 Then txtPendiente = "0"

End Sub

Private Sub sbReProgramar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curCargos As Currency, i As Integer, y As Integer
Dim curMonto As Currency, curImpVentas As Currency
Dim curSubTotal As Currency

On Error GoTo vError

vGrid.MaxRows = 0

With vGridCargos
 curCargos = 0
 For i = 1 To .MaxRows
   .col = 3
   .Row = i
   curCargos = curCargos + CCur(.Text)
 Next i
End With

  If txtFactura.Tag = "C" Then
    'Factura de Compras
    strSQL = "select CxP_Estado,Total,Imp_ventas from CPR_COMPRAS where cod_factura = '" & txtFactura _
           & "' and cod_proveedor = " & txtCodigo
  Else
    'Factura Directa
    strSQL = "select CxP_Estado,Total, 0 as 'Imp_ventas' from cxp_facturas where cod_factura = '" & txtFactura _
           & "' and cod_proveedor = " & txtCodigo
  End If
  Call OpenRecordSet(rs, strSQL)
    curMonto = txtSaldo
    curImpVentas = rs!imp_ventas
    curSubTotal = curMonto - curImpVentas
  rs.Close


'Generar Pagos
Me.MousePointer = vbHourglass

'Genera Detalle de Pagos
If CInt(txtProgNPagos) >= 1 Then
 'Paga el Impuesto de Ventas en el Primer Pago, y distribuye con base al subTotal
 If chkProgImpPrimPago.Value = vbChecked Then
    For i = 0 To CInt(txtProgNPagos) - 1
     vGrid.MaxRows = vGrid.MaxRows + 1
     vGrid.Row = vGrid.MaxRows
     vGrid.col = 1
     vGrid.Text = CStr(i + CInt(txtPagosTrans) + 1)
     vGrid.col = 2
     vGrid.Text = CStr(txtFactura)
     vGrid.col = 3
     vGrid.Text = CStr(txtCodigo)
     vGrid.col = 4
     vGrid.Text = "0"
     vGrid.col = 5
     vGrid.Text = CStr(curSubTotal / CInt(txtProgNPagos))
     vGrid.col = 6
     vGrid.Text = CStr(curSubTotal / CInt(txtProgNPagos))
     vGrid.col = 7
     vGrid.Text = "0"
     vGrid.col = 8
     vGrid.Text = CStr(DateAdd("d", (i * CInt(txtProgFrecuencia)), dtpProgPrimerPago.Value))
     vGrid.col = 9
     vGrid.Text = CStr(curSubTotal / CCur(lblTipoCambio.Caption) / CInt(txtProgNPagos))
      
    Next i
    
    vGrid.Row = 1
    vGrid.col = 5
    vGrid.Text = CStr(CCur(vGrid.Text) + curImpVentas)
    
 Else
 'Distribuye Imp.ventas en todos los pagos, con base al Total
    For i = 0 To CInt(txtProgNPagos) - 1
     vGrid.MaxRows = vGrid.MaxRows + 1
     vGrid.Row = vGrid.MaxRows
     vGrid.col = 1
     vGrid.Text = CStr(i + CInt(txtPagosTrans) + 1)
     vGrid.col = 2
     vGrid.Text = CStr(txtFactura)
     vGrid.col = 3
     vGrid.Text = CStr(txtCodigo)
     vGrid.col = 4
     vGrid.Text = "0"
     vGrid.col = 5
     vGrid.Text = CStr(curMonto / CInt(txtProgNPagos))
     vGrid.col = 6
     vGrid.Text = CStr(curMonto / CInt(txtProgNPagos))
     vGrid.col = 7
     vGrid.Text = "0"
     vGrid.col = 8
     vGrid.Text = Format(DateAdd("d", (i * CInt(txtProgFrecuencia)), dtpProgPrimerPago.Value), "yyyy/mm/dd")
     vGrid.col = 9
     vGrid.Text = CStr(curMonto / CCur(lblTipoCambio.Caption) / CInt(txtProgNPagos))
    
    Next i
 End If 'Paga Impuesto 1er Pago s/n
 
End If 'Numero de Pagos

'''Procedimiento 1 de la Aplicación de Cargos es Excluyente con el 2
If chkProCargos.Value = vbChecked And curCargos > 0 Then
 With vGrid
   For y = 1 To .MaxRows
      .Row = y
      .col = 4
      vGrid.Text = CStr(curCargos / CInt(txtProgNPagos))
   Next y
 End With
End If 'Chk y CurCargos > 0

Call sbReCalculaGrid

Me.MousePointer = vbDefault


Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaDatos()
Dim strSQL As String, rs As New ADODB.Recordset

Select Case True
  Case tcMain.Item(1).Selected
      txtSaldo = 0
      txtPagosTrans = 0
      
      txtProgNPagos = 1
      txtProgFrecuencia = 15
      dtpProgPrimerPago.Value = fxFechaServidor
      chkProCargos.Value = vbChecked
      chkProgImpPrimPago.Value = vbUnchecked
      chkProgImpPrimPago.Enabled = False
      
      strSQL = "select isnull(max(Npago),0) as 'Pago', isnull(sum(Monto),0) as 'Monto'" _
             & " From CxP_PagoPRov Where Tesoreria Is not Null" _
             & " and cod_proveedor = " & txtCodigo.Text & " and Cod_FActura = '" & txtFactura.Text & "'"
      Call OpenRecordSet(rs, strSQL)
      If Not rs.EOF And Not rs.BOF Then
        txtSaldo.Text = Format(CCur(lblFacturaMonto.Caption) - rs!Monto, "Standard")
        txtPagosTrans.Text = CStr(rs!Pago)
        If rs!Pago = 0 Then chkProgImpPrimPago.Enabled = True
      End If
      
      'Selecciona los Saldos de los Cargos Registrados al Inicio de Forma Directa
      strSQL = "select C.Cod_Cargo,C.descripcion,isnull(Sum(Monto),0) as Monto" _
             & " from CxP_Cargos C left join cxp_PagoProvCargos D on C.cod_Cargo = D.cod_Cargo" _
             & " and D.cod_Proveedor = " & txtCodigo & " and D.cod_Factura = '" & txtFactura _
             & "' and D.NPago > " & rs!Pago _
             & " group by C.Cod_Cargo,C.descripcion"
      rs.Close
      Call sbCargaGrid(vGridCargos, 3, strSQL)
      vGridCargos.MaxRows = vGridCargos.MaxRows - 1
  
  Case tcMain.Item(2).Selected
      Call sbReProgramar

End Select


End Sub

Private Function fxVerificaDatos() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim curCargos As Currency, i As Integer
Dim vMensaje As String


vMensaje = ""

On Error GoTo vError

Select Case True
  Case tcMain.Item(0).Selected
    'Verifica que se haya realizado la consulta inicial
    If CCur(lblFacturaMonto.Caption) = 0 Then
      vMensaje = vMensaje & vbCrLf & " - Consulta la Factura antes de seguir al siguiente paso!"
    End If
    
    'Ahora se permite reprogramar facturas de contado and Forma_Pago = 'CR'
    'Verificar que la factura exista en la programacion
    strSQL = "select cod_factura,Tipo From vCxP_ProgramacionPago" _
           & " where CxP_Estado = 'G' and Cod_Proveedor = " & txtCodigo _
           & " and cod_factura = '" & txtFactura & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs.EOF And rs.BOF Then
      vMensaje = vMensaje & vbCrLf & " - La Factura No Existe en la Programación"
    Else
      txtFactura.Tag = rs!Tipo
    End If
    rs.Close
    
  Case tcMain.Item(1).Selected
    'Verificar Montos
    
    If Not IsNumeric(txtProgNPagos) Then
     vMensaje = vMensaje & " - El # Pagos no es válido" & vbCrLf
    Else
     txtProgNPagos = CInt(txtProgNPagos)
     If CInt(txtProgNPagos) < 1 Then vMensaje = vMensaje & " - El # Pagos no es válido" & vbCrLf
    End If
    
    If Not IsNumeric(txtProgFrecuencia) Then
     vMensaje = vMensaje & " - La frecuencia de pagos no es válida" & vbCrLf
    Else
     txtProgFrecuencia = CInt(txtProgFrecuencia)
    End If
    
    If CCur(txtSaldo) = 0 Then
     vMensaje = vMensaje & " - No Existe Saldo Pendiente a Programar" & vbCrLf
    End If
    
    With vGridCargos
     curCargos = 0
     For i = 1 To .MaxRows
       .col = 3
       .Row = i
       curCargos = curCargos + CCur(.Text)
     Next i
    End With
    If curCargos > CCur(txtSaldo) Then vMensaje = vMensaje & " - Los Cargos Son Mayores que el Monto de la Factura" & vbCrLf
    
  
  Case tcMain.Item(2).Selected
   
   Call sbReCalculaGrid
   If CCur(txtPendiente) > 0 Then vMensaje = vMensaje & " - Los montos distribuidos estan dejando un pendiente (Revisar)" & vbCrLf
   With vGrid
     For i = 1 To .MaxRows
        .Row = i
        .col = 6
        If CCur(vGrid.Text) < 0 Then
          .col = 1
          vMensaje = vMensaje & " - Los Cargos en el Pago " & vGrid.Text & " son mayores que el Monto del Pago ..." & vbCrLf
        End If
     Next i
   End With

End Select

If Len(vMensaje) > 0 Then
  fxVerificaDatos = False
  MsgBox vMensaje, vbExclamation
Else
  fxVerificaDatos = True
End If

Exit Function

vError:
 fxVerificaDatos = False
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Function


Private Sub cmdAtras_Click()
Dim i As Integer

If tcMain.SelectedItem > 0 Then
' ssTab.Tab = ssTab.Tab - 1
 tcMain.Item(tcMain.SelectedItem - 1).Selected = True
 
' For i = 0 To ssTab.Tabs - 1
'   ssTab.TabEnabled(i) = False
' Next i
 
 For i = 0 To tcMain.ItemCount - 1
    tcMain.Item(i).Enabled = False
 Next i
 
 'Preguntar si desea limpiar los datos
 i = MsgBox("Desea Limpiar Los Datos Anteriores...", vbYesNo)
 If i = vbYes Then
   sbLimpiaDatos
   sbCargaDatos
 End If
End If

'ssTab.TabEnabled(ssTab.Tab) = True

tcMain.Item(tcMain.SelectedItem).Enabled = True

End Sub


Private Sub cmdSiguiente_Click()
If tcMain.Item(0).Selected = True Then
 
 If fxVerificaDatos Then
'    ssTab.Tab = ssTab.Tab + 1
    
    tcMain.Item(tcMain.SelectedItem + 1).Selected = True
    
    Call sbLimpiaDatos
    Call sbCargaDatos
 End If

Else
 
 If tcMain.SelectedItem < 2 Then
'   ssTab.Tab = ssTab.Tab + 1
    tcMain.Item(tcMain.SelectedItem + 1).Selected = True
   Call sbLimpiaDatos
   Call sbCargaDatos
 End If

End If

End Sub


Private Sub sbLimpiaDatos()

Select Case True
  Case tcMain.Item(0).Selected  'Paso 1
      
    txtCodigo = ""
    txtNombre = ""
    txtFactura = ""
    
    tcMain.Item(0).Enabled = True
    tcMain.Item(1).Enabled = False
    tcMain.Item(2).Enabled = False

  
  Case tcMain.Item(1).Selected 'Paso 2
    
    tcMain.Item(0).Enabled = False
    tcMain.Item(1).Enabled = True
    tcMain.Item(2).Enabled = False
     
     
  Case tcMain.Item(2).Selected 'Re-Programacion
    vGrid.MaxCols = 9
    
    
    tcMain.Item(0).Enabled = False
    tcMain.Item(1).Enabled = False
    tcMain.Item(2).Enabled = True
  
End Select

End Sub


Private Sub vGrid_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If col = 5 Then
  Call sbReCalculaGrid
End If

End Sub
