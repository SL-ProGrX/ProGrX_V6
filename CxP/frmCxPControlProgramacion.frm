VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxPControlProgramacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Pago : Programación de Pagos"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   11295
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      _Version        =   1441793
      _ExtentX        =   19923
      _ExtentY        =   13996
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
      Item(0).Caption =   "Facturas"
      Item(0).ControlCount=   19
      Item(0).Control(0)=   "cboTipo"
      Item(0).Control(1)=   "txtFactura"
      Item(0).Control(2)=   "txtPrgProveedor"
      Item(0).Control(3)=   "txtProgDivisa"
      Item(0).Control(4)=   "txtProgMonto"
      Item(0).Control(5)=   "txtProgCodProv"
      Item(0).Control(6)=   "Label1(5)"
      Item(0).Control(7)=   "Label1(4)"
      Item(0).Control(8)=   "Label1(0)"
      Item(0).Control(9)=   "Label1(1)"
      Item(0).Control(10)=   "Label1(2)"
      Item(0).Control(11)=   "Label1(3)"
      Item(0).Control(12)=   "chkFacturaSaldo"
      Item(0).Control(13)=   "opt(0)"
      Item(0).Control(14)=   "opt(1)"
      Item(0).Control(15)=   "opt(2)"
      Item(0).Control(16)=   "lswProg"
      Item(0).Control(17)=   "txtLineas"
      Item(0).Control(18)=   "Label3(0)"
      Item(1).Caption =   "Programación"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "lblProgProveedor"
      Item(1).Control(1)=   "lblProgFactura"
      Item(1).Control(2)=   "Label2(3)"
      Item(1).Control(3)=   "Label2(4)"
      Item(1).Control(4)=   "tcPro"
      Begin XtremeSuiteControls.ListView lswProg 
         Height          =   6015
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19494
         _ExtentY        =   10604
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
         HideSelection   =   0   'False
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   315
         Left            =   9600
         TabIndex        =   1
         Top             =   720
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
      Begin XtremeSuiteControls.FlatEdit txtFactura 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   720
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPrgProveedor 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   720
         Width           =   3855
         _Version        =   1441793
         _ExtentX        =   6794
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtProgDivisa 
         Height          =   315
         Left            =   8280
         TabIndex        =   4
         Top             =   720
         Width           =   1335
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
      Begin XtremeSuiteControls.FlatEdit txtProgMonto 
         Height          =   315
         Left            =   6840
         TabIndex        =   5
         Top             =   720
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.FlatEdit txtProgCodProv 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   720
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
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
      Begin XtremeSuiteControls.CheckBox chkFacturaSaldo 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Facturas con Saldo"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Pendientes"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   2
         Left            =   6360
         TabIndex        =   16
         Top             =   1080
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Programadas"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtLineas 
         Height          =   315
         Left            =   10200
         TabIndex        =   18
         Top             =   7560
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
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
         Text            =   "100"
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcPro 
         Height          =   7095
         Left            =   -70000
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   11295
         _Version        =   1441793
         _ExtentX        =   19923
         _ExtentY        =   12515
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
         Item(0).Caption =   "Generación"
         Item(0).ControlCount=   4
         Item(0).Control(0)=   "GroupBox1(0)"
         Item(0).Control(1)=   "gbFactura"
         Item(0).Control(2)=   "GroupBox1(1)"
         Item(0).Control(3)=   "gbAplicar"
         Item(1).Caption =   "Detalle de Pagos"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "vGrid"
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   3255
            Index           =   0
            Left            =   5160
            TabIndex        =   25
            Top             =   2760
            Width           =   5775
            _Version        =   1441793
            _ExtentX        =   10181
            _ExtentY        =   5736
            _StockProps     =   79
            Caption         =   "Cargos aplicables:"
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
            BorderStyle     =   1
            Begin FPSpreadADO.fpSpread vGridCargos 
               Height          =   2412
               Left            =   0
               TabIndex        =   26
               Top             =   360
               Width           =   5652
               _Version        =   524288
               _ExtentX        =   9970
               _ExtentY        =   4255
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
               SpreadDesigner  =   "frmCxPControlProgramacion.frx":0000
               VScrollSpecial  =   -1  'True
               VScrollSpecialType=   2
               AppearanceStyle =   1
            End
            Begin XtremeSuiteControls.FlatEdit txtTotalCargos 
               Height          =   312
               Left            =   3840
               TabIndex        =   27
               Top             =   2880
               Width           =   1572
               _Version        =   1441793
               _ExtentX        =   2773
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
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPagoMin 
               Height          =   312
               Left            =   960
               TabIndex        =   28
               Top             =   2880
               Width           =   1572
               _Version        =   1441793
               _ExtentX        =   2773
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
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label2 
               Caption         =   "Pago Min:"
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
               Left            =   0
               TabIndex        =   30
               Top             =   2880
               Width           =   1332
            End
            Begin VB.Label Label2 
               Caption         =   "Total "
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
               Left            =   3000
               TabIndex        =   29
               Top             =   2880
               Width           =   852
            End
         End
         Begin XtremeSuiteControls.GroupBox gbFactura 
            Height          =   3255
            Left            =   120
            TabIndex        =   31
            Top             =   2760
            Width           =   4575
            _Version        =   1441793
            _ExtentX        =   8064
            _ExtentY        =   5736
            _StockProps     =   79
            Caption         =   "Datos de la Factura:"
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
            BorderStyle     =   1
            Begin VB.Label Label2 
               Caption         =   "(Divisa Local)"
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
               Index           =   14
               Left            =   3360
               TabIndex        =   46
               Top             =   1200
               Width           =   1212
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
               Height          =   252
               Index           =   8
               Left            =   240
               TabIndex        =   45
               Top             =   2880
               Width           =   1212
            End
            Begin VB.Label lblSaldoFact 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   1560
               TabIndex        =   44
               Top             =   2880
               Width           =   1692
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
               Height          =   252
               Index           =   6
               Left            =   240
               TabIndex        =   43
               Top             =   1200
               Width           =   1212
            End
            Begin VB.Label lblMontoFactura 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   1560
               TabIndex        =   42
               Top             =   1200
               Width           =   1692
            End
            Begin VB.Label lblFechaFactura 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   1560
               TabIndex        =   41
               Top             =   360
               Width           =   1692
            End
            Begin VB.Label Label2 
               Caption         =   "Fecha"
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
               Left            =   240
               TabIndex        =   40
               Top             =   360
               Width           =   1212
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
               Height          =   252
               Index           =   11
               Left            =   240
               TabIndex        =   39
               Top             =   1560
               Width           =   1212
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
               Height          =   252
               Index           =   12
               Left            =   240
               TabIndex        =   38
               Top             =   1920
               Width           =   1212
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
               Height          =   252
               Index           =   13
               Left            =   240
               TabIndex        =   37
               Top             =   2280
               Width           =   1212
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
               Height          =   252
               Index           =   15
               Left            =   240
               TabIndex        =   36
               Top             =   720
               Width           =   1212
            End
            Begin VB.Label lblVencimiento 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   1560
               TabIndex        =   35
               Top             =   720
               Width           =   1692
            End
            Begin VB.Label lblImporteReal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   1560
               TabIndex        =   34
               Top             =   1560
               Width           =   1692
            End
            Begin VB.Label lblDivisa 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   1560
               TabIndex        =   33
               Top             =   1920
               Width           =   1692
            End
            Begin VB.Label lblTipoCambio 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   1560
               TabIndex        =   32
               Top             =   2280
               Width           =   1692
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   2175
            Index           =   1
            Left            =   240
            TabIndex        =   47
            Top             =   480
            Width           =   10815
            _Version        =   1441793
            _ExtentX        =   19076
            _ExtentY        =   3836
            _StockProps     =   79
            Caption         =   "Programación de Pagos:"
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
            BorderStyle     =   1
            Begin VB.TextBox txtProgNPagos 
               Alignment       =   2  'Center
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
               Left            =   1920
               TabIndex        =   49
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtProgFrecuencia 
               Alignment       =   2  'Center
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
               Left            =   1920
               TabIndex        =   48
               Top             =   720
               Width           =   1335
            End
            Begin FPSpreadADO.fpSpread vGridVence 
               Height          =   1812
               Left            =   6360
               TabIndex        =   50
               Top             =   360
               Width           =   4092
               _Version        =   524288
               _ExtentX        =   7218
               _ExtentY        =   3196
               _StockProps     =   64
               AllowCellOverflow=   -1  'True
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
               SpreadDesigner  =   "frmCxPControlProgramacion.frx":0596
               VScrollSpecial  =   -1  'True
               VScrollSpecialType=   2
               AppearanceStyle =   1
            End
            Begin XtremeSuiteControls.DateTimePicker dtpProgPrimerPago 
               Height          =   312
               Left            =   1920
               TabIndex        =   51
               Top             =   1080
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
            Begin XtremeSuiteControls.CheckBox chkProgImpPrimPago 
               Height          =   372
               Left            =   240
               TabIndex        =   52
               Top             =   1440
               Width           =   3012
               _Version        =   1441793
               _ExtentX        =   5313
               _ExtentY        =   656
               _StockProps     =   79
               Caption         =   "Cancelar el I.V. en el 1 er. Pago"
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
               Alignment       =   1
            End
            Begin XtremeSuiteControls.CheckBox chkProCargos 
               Height          =   372
               Left            =   240
               TabIndex        =   53
               Top             =   1800
               Width           =   3012
               _Version        =   1441793
               _ExtentX        =   5313
               _ExtentY        =   656
               _StockProps     =   79
               Caption         =   "Cargos entre el No. de Pagos"
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
               Alignment       =   1
            End
            Begin XtremeSuiteControls.FlatEdit txtDistribuido 
               Height          =   312
               Left            =   3960
               TabIndex        =   54
               Top             =   1800
               Width           =   1812
               _Version        =   1441793
               _ExtentX        =   3196
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
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label2 
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
               Height          =   252
               Index           =   0
               Left            =   240
               TabIndex        =   62
               Top             =   360
               Width           =   1332
            End
            Begin VB.Label Label2 
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
               Height          =   252
               Index           =   1
               Left            =   240
               TabIndex        =   61
               Top             =   1080
               Width           =   1572
            End
            Begin VB.Label Label2 
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
               Left            =   240
               TabIndex        =   60
               Top             =   720
               Width           =   1452
            End
            Begin VB.Label lblDiasCredito 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   3960
               TabIndex        =   59
               Top             =   600
               Width           =   1812
            End
            Begin VB.Label lblSaldoProv 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   3960
               TabIndex        =   58
               Top             =   1200
               Width           =   1812
            End
            Begin VB.Label Label2 
               Caption         =   "Saldo Proveedor:"
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
               Left            =   3600
               TabIndex        =   57
               Top             =   960
               Width           =   1332
            End
            Begin VB.Label Label2 
               Caption         =   "Dias de Crédito:"
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
               Left            =   3600
               TabIndex        =   56
               Top             =   360
               Width           =   1572
            End
            Begin VB.Label Label2 
               Caption         =   "Distribuido:"
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
               Left            =   3600
               TabIndex        =   55
               Top             =   1560
               Width           =   1332
            End
         End
         Begin FPSpreadADO.fpSpread vGrid 
            Height          =   6495
            Left            =   -69760
            TabIndex        =   63
            Top             =   480
            Visible         =   0   'False
            Width           =   10815
            _Version        =   524288
            _ExtentX        =   19071
            _ExtentY        =   11451
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
            SpreadDesigner  =   "frmCxPControlProgramacion.frx":0B53
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.GroupBox gbAplicar 
            Height          =   1095
            Left            =   0
            TabIndex        =   64
            Top             =   6120
            Width           =   11295
            _Version        =   1441793
            _ExtentX        =   19923
            _ExtentY        =   1931
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.PushButton cmdProgAplicar 
               Height          =   615
               Left            =   9480
               TabIndex        =   65
               Top             =   240
               Width           =   1575
               _Version        =   1441793
               _ExtentX        =   2773
               _ExtentY        =   1080
               _StockProps     =   79
               Caption         =   "Aplicar"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   16
               Picture         =   "frmCxPControlProgramacion.frx":1453
            End
            Begin XtremeSuiteControls.CheckBox chkAplCarPerPorc 
               Height          =   375
               Left            =   120
               TabIndex        =   66
               Top             =   240
               Width           =   4215
               _Version        =   1441793
               _ExtentX        =   7429
               _ExtentY        =   656
               _StockProps     =   79
               Caption         =   "Aplica Cargos Periodicos Porcentuales"
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
            Begin XtremeSuiteControls.CheckBox chkAplCarPerMonto 
               Height          =   375
               Left            =   120
               TabIndex        =   67
               Top             =   600
               Width           =   4215
               _Version        =   1441793
               _ExtentX        =   7429
               _ExtentY        =   656
               _StockProps     =   79
               Caption         =   "Aplica Cargos Periodicos por Monto"
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
         End
      End
      Begin VB.Label Label2 
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
         Height          =   315
         Index           =   4
         Left            =   -65800
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Height          =   315
         Index           =   3
         Left            =   -69880
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblProgFactura 
         Caption         =   "..."
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
         Height          =   315
         Left            =   -68800
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblProgProveedor 
         Caption         =   "..."
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
         Height          =   315
         Left            =   -64840
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "Líneas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   9240
         TabIndex        =   19
         Top             =   7560
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   6840
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nombre del Proveedor"
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
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Prov. (ID)"
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
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Forma de Pago"
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
         Height          =   255
         Index           =   4
         Left            =   9600
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   8280
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCxPControlProgramacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vInicia As Boolean, vPaso As Boolean
Dim mcurMonto As Currency, mcurImpVentas As Currency
Dim mcurSubTotal As Currency

Private Sub cboTipo_Click()
Call sbProgLlenaLsw
End Sub

Private Sub sbPagos_Calcula_Cambio(col As Long)
Dim curCargos As Currency, i As Integer, y As Integer
Dim curPago As Currency, vFecha As Date
Dim curTotal As Currency, curSubTotal As Currency

On Error GoTo vError

If Len(lblProgFactura.Caption) = 0 Then Exit Sub
If col = 2 Then Exit Sub

curTotal = 0

'Aplica siempre el monto completo
curSubTotal = mcurMonto

txtDistribuido.Text = Format(0, "Standard")
curTotal = 0

curCargos = CCur(txtTotalCargos.Text) / CInt(txtProgNPagos)

vPaso = True


'Cambio de Porcentaje
If col = 1 Then
    With vGridVence
      For i = 1 To CInt(txtProgNPagos)
          .Row = i
          .col = 1
          
          curPago = curSubTotal * CCur(.Value)
          
          .col = 3
          .Text = Format(curPago, "Standard")
      Next i
      
    End With
End If 'Cambio de Porcentaje


'Cambio de Monto
If col = 3 Then
    With vGridVence
      For i = 1 To CInt(txtProgNPagos)
          .Row = i
          .col = 3
          
          curPago = CCur(.Text)
          
          .col = 1
          .Value = curPago / curSubTotal
      Next i
      
    End With
End If 'Cambio de Porcentaje


'Calcula Monto Distribuido
With vGridVence
  For i = 1 To .MaxRows
    .Row = i
    .col = 3
    curTotal = curTotal + CCur(.Text)
  Next i
End With

txtDistribuido.Text = Format(curTotal, "Standard")

vPaso = False
  
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  vPaso = False
 
End Sub



Private Sub sbPagos_Calcula()
Dim curCargos As Currency, i As Integer, y As Integer
Dim curPago As Currency, vFecha As Date
Dim curTotal As Currency, curSubTotal As Currency
Dim vMensaje As String

On Error GoTo vError

vMensaje = ""
vGridVence.MaxRows = 0
vGridVence.MaxCols = 3

If Len(lblProgFactura.Caption) = 0 Then
 vMensaje = vMensaje & " - N se ha indicado un numero de factura válido" & vbCrLf
End If

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

If Len(vMensaje) > 0 Then
    Exit Sub
End If

curTotal = 0
If chkProgImpPrimPago.Value = xtpChecked Then
   curSubTotal = mcurSubTotal
Else
  curSubTotal = mcurMonto
End If

txtDistribuido.Text = Format(0, "Standard")
curTotal = 0

vPaso = True

With vGridVence
  curCargos = CCur(txtTotalCargos.Text) / CInt(txtProgNPagos)
  .MaxRows = 0
  For i = 0 To CInt(txtProgNPagos) - 1
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
      curPago = curSubTotal / CInt(txtProgNPagos)
      vFecha = DateAdd("d", (i * CInt(txtProgFrecuencia)), dtpProgPrimerPago.Value)
      
      If chkProgImpPrimPago.Value = vbChecked And i = 0 Then
        curPago = curPago + mcurImpVentas
      End If
      
      .col = 1
      .Text = CStr((curPago / mcurMonto) * 100)
      .col = 2
      .Text = Format(vFecha, "dd/mm/yyyy")
      .col = 3
      .Text = Format(curPago, "Standard")
         
         
      curTotal = curTotal + curPago
  Next i
  
End With

txtDistribuido.Text = Format(curTotal, "Standard")

vPaso = False
  
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  vPaso = False
 
End Sub


Private Sub chkProCargos_Click()
 Call sbPagos_Calcula
End Sub

Private Sub chkProgImpPrimPago_Click()
 Call sbPagos_Calcula
End Sub

Private Sub cmdProgAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTransac As Boolean, vMensaje As String
Dim curCargos As Currency, i As Integer, y As Integer
Dim curMonto As Currency, curImpVentas As Currency
Dim curSubTotal As Currency
Dim pFecha As Date, pMonto As Currency, pPagoMin As Currency

On Error GoTo vError

vTransac = False
vMensaje = ""

'Verifica Datos
'1. Verifica que la factura este pendiente y que exista
'2. Verifica validez de los parametros


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

With vGridCargos
 curCargos = 0
 For i = 1 To .MaxRows
   .Row = i
   .col = 3
   curCargos = curCargos + CCur(.Text)
 Next i
End With

'Valida el Monto Mínimo de Pagos

pFecha = fxFechaServidor
pPagoMin = curCargos / CInt(txtProgNPagos.Text)
With vGridVence
 For i = 1 To .MaxRows
   .Row = i
   
   .col = 2
   
   If DateDiff("d", pFecha, (Mid(.Value, 5, 4) & "/" & Mid(.Value, 1, 2) & "/" & Mid(.Value, 3, 2))) < 0 Then
     vMensaje = vMensaje & " - El Vencimiento del Pago No." & i & " no puede ser menor al anterior!" & vbCrLf
   End If
   
   pFecha = Mid(.Value, 5, 4) & "/" & Mid(.Value, 1, 2) & "/" & Mid(.Value, 3, 2)
   
   
   .col = 3
       
   If CCur(.Text) < pPagoMin Then
     vMensaje = vMensaje & " - El Pago No." & i & " no es válido. Es menor que el monto minimo para cobro de cargos!" & vbCrLf
   End If
 
 
 Next i
End With




If Len(lblProgFactura.Caption) > 0 Then
  If txtProgNPagos.Tag = "C" Then
    'Factura de Compras
    strSQL = "select CxP_Estado,Total,Imp_ventas from CPR_COMPRAS where cod_factura = '" & lblProgFactura.Caption _
           & "' and cod_proveedor = " & lblProgProveedor.Tag
  Else
    'Factura Directa
    strSQL = "select CxP_Estado,Total, isnull(impuesto_ventas,0) as 'Imp_ventas' from cxp_facturas where cod_factura = '" & lblProgFactura.Caption _
           & "' and cod_proveedor = " & lblProgProveedor.Tag
  End If
    
    Call OpenRecordSet(rs, strSQL)
    If rs.EOF And rs.BOF Then
     vMensaje = vMensaje & " - No Existe la Factura..." & vbCrLf
    Else
     If curCargos > rs!Total Then vMensaje = vMensaje & " - Los Cargos Son Mayores que el Monto de la Factura" & vbCrLf
     If CCur(txtDistribuido.Text) <> rs!Total Then vMensaje = vMensaje & " - No se ha distribuido el monto CORRECTO de la Factura!" & vbCrLf
     
     If rs!CxP_Estado <> "P" Then vMensaje = vMensaje & " - La Factura no se encuentra pendiente..." & vbCrLf
     curMonto = rs!Total
     curImpVentas = rs!imp_ventas
     curSubTotal = curMonto - curImpVentas
    End If
    rs.Close

Else
   vMensaje = vMensaje & " - No se especificó factura..." & vbCrLf
End If

If Len(vMensaje) > 0 Then
 MsgBox vMensaje, vbExclamation, "Operación Cancelada"
 Exit Sub
End If

'Generar Pagos

Me.MousePointer = vbHourglass

glogon.Conection.BeginTrans
vTransac = True

'Reflejar Cargos al Saldo del proveedor
If curCargos > 0 Then
    strSQL = "update cxp_proveedores set saldo = isnull(saldo,0) - " & curCargos _
           & ",SALDO_DIVISA_REAL = isnull(SALDO_DIVISA_REAL,0) - " & curCargos / CCur(lblTipoCambio.Caption) _
           & " where cod_proveedor = " & lblProgProveedor.Tag
    Call ConectionExecute(strSQL)
End If

'Actualiza Estado de la Programacion de la factura
If txtProgNPagos.Tag = "C" Then
    strSQL = "update CPR_COMPRAS set cxp_estado = 'G' where cod_factura = '" & lblProgFactura.Caption _
           & "' and cod_proveedor = " & lblProgProveedor.Tag
Else
    strSQL = "update cxp_facturas set cxp_estado = 'G' where cod_factura = '" & lblProgFactura.Caption _
           & "' and cod_proveedor = " & lblProgProveedor.Tag
End If
Call ConectionExecute(strSQL)


'Genera Detalle de Pagos
'----Nuevo:2018/09/07

strSQL = ""

With vGridVence
    For i = 1 To .MaxRows
       .Row = i
       .col = 2
       pFecha = Mid(.Value, 5, 4) & "/" & Mid(.Value, 1, 2) & "/" & Mid(.Value, 3, 2)
       .col = 3
       pMonto = CCur(.Text)
       
         strSQL = strSQL & Space(10) & "insert cxp_pagoProv(npago,cod_proveedor,cod_factura,fecha_vencimiento,monto" _
                & ",frecuencia,tipo_transac,apl_cargo_flotante,pago_anticipado,forma_pago" _
                & ",importe_divisa_real,tipo_cambio,cod_divisa) values(" & i & "," & lblProgProveedor.Tag _
                & ",'" & lblProgFactura.Caption & "','" & Format(pFecha, "yyyy/mm/dd") & "'," & pMonto _
                & "," & txtProgFrecuencia & "," & IIf((CInt(txtProgNPagos) = 1), 0, 1) _
                & "," & chkAplCarPerMonto.Value & ",0,'CR'," & (pMonto / CCur(lblTipoCambio.Caption)) _
                & "," & CCur(lblTipoCambio.Caption) & ",'" & lblDivisa.Caption & "')"
    Next i
End With

'Registra los Pagos
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
End If


'Procedimiento 1 de la Aplicación de Cargos es Excluyente con el 2
strSQL = ""
If chkProCargos.Value = vbChecked And curCargos > 0 Then
 With vGridCargos
   For y = 1 To .MaxRows
      .Row = y
      .col = 3
      curCargos = CCur(.Text) / CInt(txtProgNPagos)
      If curCargos > 0 Then
        For i = 1 To CInt(txtProgNPagos)
           .col = 1
           strSQL = strSQL & Space(10) & "insert cxp_PagoProvCargos(Npago,Cod_factura,cod_proveedor,cod_cargo,monto,registro_fecha,registro_usuario" _
                  & ",cod_divisa,tipo_cambio,tipo_cargo,tipo_proceso)" _
                  & " values(" & i & ",'" & lblProgFactura.Caption & "'," & lblProgProveedor.Tag _
                  & ",'" & Trim(.Text) & "'," & curCargos & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & lblDivisa.Caption _
                  & "'," & CCur(lblTipoCambio.Caption) & ",'M','D')"
        Next i
      End If
   Next y
 End With
End If 'Chk y CurCargos > 0

'Procesa Todos los Cargos
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
End If

'--------------------------------------------------------------------
'Necesita Cerrar la Transacción, y trabajar a modo implicito
'Procedimiento 2 de la Aplicación de Cargos es Excluyente con el 1

glogon.Conection.CommitTrans
vTransac = False

If chkProCargos.Value = vbUnchecked And curCargos > 0 Then
 With vGridCargos
   For y = 1 To .MaxRows
      .Row = y
      .col = 3
      curCargos = CCur(.Text)
      If curCargos > 0 Then
        For i = 1 To CInt(txtProgNPagos)
           .col = 1
           
           'Saca el Disponible para Aplicarle el Cargo
           strSQL = "select P.Npago,(P.Monto - (isnull(sum(C.monto),0))) as Neto" _
                  & " from cxp_pagoProv P left join cxp_pagoProvCargos C on P.npago = C.npago" _
                  & " and P.cod_factura = C.cod_factura and P.cod_proveedor = C.cod_proveedor" _
                  & " where P.npago = " & i & " and P.cod_factura = '" & lblProgFactura.Caption _
                  & "' and P.cod_proveedor = " & lblProgProveedor.Tag _
                  & " group by P.NPago,P.Monto"
           Call OpenRecordSet(rs, strSQL)
           If rs!Neto > curCargos Then
                strSQL = "insert cxp_PagoProvCargos(Npago,Cod_factura,cod_proveedor,cod_cargo,monto,registro_fecha,registro_usuario" _
                       & ",cod_divisa,tipo_cambio,tipo_cargo,tipo_proceso)" _
                       & " values(" & i & ",'" & lblProgFactura.Caption & "'," & lblProgProveedor.Tag _
                       & ",'" & Trim(.Text) & "'," & curCargos & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & lblDivisa.Caption _
                       & "'," & CCur(lblTipoCambio.Caption) & ",'M','D')"
                Call ConectionExecute(strSQL)
           
                curCargos = 0
           Else
                strSQL = "insert cxp_PagoProvCargos(Npago,Cod_factura,cod_proveedor,cod_cargo,monto,registro_fecha,registro_usuario" _
                       & ",cod_divisa,tipo_cambio,tipo_cargo,tipo_proceso)" _
                       & " values(" & i & ",'" & lblProgFactura.Caption & "'," & lblProgProveedor.Tag _
                       & ",'" & Trim(.Text) & "'," & rs!Neto & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & lblDivisa.Caption _
                       & "'," & CCur(lblTipoCambio.Caption) & ",'M','D')"
                Call ConectionExecute(strSQL)
                
                curCargos = curCargos - rs!Neto
           End If
           rs.Close
           
           If curCargos <= 0 Then Exit For
        
        Next i
      
      End If
   Next y
 End With
End If 'UbChk y CurCargos > 0

'Revisar si el Pago, se cancela con el cargo asignado. Dejando el Monto del Pago en Cero
'En este caso desactivar el Envio a Tesoreria. Y por ende de la antiguedad de saldos y programacion de pagos

strSQL = "update cxp_pagoProv set Tesoreria = 0,fecha_traslada = dbo.MyGetdate(),user_traslada = '" & glogon.Usuario _
       & "' where cod_proveedor = " & lblProgProveedor.Tag & " and cod_factura = '" & lblProgFactura.Caption _
       & "' and Npago in(select P.npago from cxp_pagoProv P inner join cxp_PagoprovCargos C" _
       & " on P.cod_proveedor = C.cod_proveedor and P.cod_factura = C.cod_factura and P.npago = C.npago" _
       & " where P.cod_proveedor = " & lblProgProveedor.Tag & " and P.cod_factura = '" & lblProgFactura.Caption _
       & "' group by P.npago,P.cod_proveedor,P.cod_factura,P.monto Having P.Monto = isnull(Sum(C.Monto), 0))"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Generación de Cola de Pagos Realizada Satisfactoriamente...", vbInformation


tcPro.Item(1).Selected = True
'Call tcPro_SelectedChanged
Exit Sub
vError:
  If vTransac Then glogon.Conection.RollbackTrans
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub dtpProgPrimerPago_Change()
If Not vPaso Then
 Call sbPagos_Calcula
End If
End Sub

Private Sub Form_Activate()
vModulo = 30
End Sub

Private Sub Form_Load()

vModulo = 30

tcMain.Item(0).Selected = True
tcPro.Item(0).Selected = True

vGrid.AppearanceStyle = fxGridStyle
vGridCargos.AppearanceStyle = fxGridStyle

dtpProgPrimerPago.Value = fxFechaServidor


With lswProg.ColumnHeaders
 .Clear
 .Add , , "No. Factura", 2000
 .Add , , "Prov.Id.", 1000, vbCenter
 .Add , , "Proveedor", 3400
 .Add , , "Total", 1400, vbRightJustify
 .Add , , "Fecha", 1100, vbCenter
 .Add , , "Estado", 1100, vbCenter
 .Add , , "Tipo", 1100, vbCenter
 .Add , , "F.Pago", 1100, vbCenter
 .Add , , "Divisa", 1000, vbCenter
 .Add , , "T.C.", 1100, vbRightJustify
 .Add , , "Importe Real", 1400, vbRightJustify
 .Add , , "Vencimiento", 1300, vbCenter

End With

vInicia = True
Call opt_Click(1)


cboTipo.Clear
cboTipo.AddItem "Contado"
cboTipo.AddItem "Crédito"
cboTipo.AddItem "Todas"
cboTipo.Text = "Todas"

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbProgCargosIni()
Dim strSQL As String

strSQL = "select cod_Cargo,descripcion,0 as Monto " _
       & " from cxp_cargos where Activo = 1"
Call sbCargaGrid(vGridCargos, 3, strSQL)
'Limpia espacios vacios
vGridCargos.MaxRows = vGridCargos.MaxRows - 1

txtTotalCargos.Text = Format(0, "Standard")

End Sub


Private Sub lswProg_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

Call sbProgCargosIni

lblProgFactura.Caption = Item.Text
lblProgProveedor.Caption = Item.SubItems(2)
lblProgProveedor.Tag = Item.SubItems(1)
lblMontoFactura.Caption = Item.SubItems(3)
lblFechaFactura.Caption = Item.SubItems(4)

lblDivisa.Caption = Item.SubItems(8)
lblTipoCambio.Caption = Item.SubItems(9)
lblImporteReal.Caption = Item.SubItems(10)
lblVencimiento.Caption = Item.SubItems(11)


txtProgNPagos.Tag = Mid(Item.SubItems(6), 1, 1)

vGridCargos.Enabled = True
vGridVence.Enabled = True

'txtProgNPagos.SetFocus

tcMain.Item(1).Selected = True


'Dias de Credito / Saldo
strSQL = "SELECT CREDITO_PLAZO,dbo.fxCxPSaldoCorte(cod_proveedor,dbo.Mygetdate()) as 'SALDO'" _
       & " FROM CXP_PROVEEDORES" _
       & " where cod_proveedor = " & lblProgProveedor.Tag
Call OpenRecordSet(rs, strSQL)
    lblDiasCredito.Caption = rs!credito_plazo
    lblSaldoProv.Caption = Format(rs!Saldo, "Standard")
rs.Close

'Saldo Factura
strSQL = "select isnull(sum(monto),0) as Saldo From cxp_pagoprov" _
       & " Where cod_proveedor = " & lblProgProveedor.Tag _
       & " and cod_factura = '" & lblProgFactura.Caption _
       & "' and tesoreria is null"
Call OpenRecordSet(rs, strSQL)
    lblSaldoFact.Caption = Format(rs!Saldo, "Standard")
rs.Close



If txtProgNPagos.Tag = "C" Then
  'Factura de Compras
  strSQL = "select CxP_Estado,Total,Imp_ventas from CPR_COMPRAS where cod_factura = '" & lblProgFactura.Caption _
         & "' and cod_proveedor = " & lblProgProveedor.Tag
Else
  'Factura Directa
  strSQL = "select CxP_Estado,Total, isnull(impuesto_Ventas,0) as 'Imp_ventas' from cxp_facturas where cod_factura = '" & lblProgFactura.Caption _
         & "' and cod_proveedor = " & lblProgProveedor.Tag
End If

  Call OpenRecordSet(rs, strSQL)
  If rs.EOF And rs.BOF Then
    mcurMonto = 0
    mcurImpVentas = 0
    mcurSubTotal = mcurMonto - mcurImpVentas
  Else
    mcurMonto = rs!Total
    mcurImpVentas = rs!imp_ventas
    mcurSubTotal = mcurMonto - mcurImpVentas
  End If
  rs.Close

'Calcula Plan de Pagos
Call sbPagos_Calcula

If Item.Tag = "P" Then
    tcPro.Item(0).Selected = True
Else
    tcPro.Item(1).Selected = True
'    Call ssTabPro_Click(1)
End If

End Sub

Private Sub opt_Click(Index As Integer)
Call sbProgLlenaLsw
End Sub



Private Function fxTesoreriaConsulta(NSolicitud As Long) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String

On Error GoTo vError

vCadena = ""

strSQL = "select C.estado,C.tipo,B.descripcion,C.beneficiario,C.monto" _
       & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_banco = B.id_Banco" _
       & " where C.Nsolicitud = " & NSolicitud
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  vCadena = "TIPO : " & rs!Tipo & vbCrLf & "BANCO : " & rs!Descripcion _
          & vbCrLf & "ESTADO : "
 Select Case rs!Estado
     Case "A"
        vCadena = vCadena & "ANULADO"
     Case "I", "T"
        vCadena = vCadena & "EMITIDO"
     Case "P"
        vCadena = vCadena & "PENDIENTE"
 End Select
           
 vCadena = vCadena & vbCrLf & "MONTO : " & Format(rs!Monto, "Standard") _
         & vbCrLf & "BENEFICIARIO : " & rs!Beneficiario
          
End If
rs.Close

vError:

fxTesoreriaConsulta = vCadena


End Function


Private Sub sbCargaGridLocal01(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.col = i
 vGrid.Text = ""
Next i

Call OpenRecordSet(rs, strSQL)
Do While rs.EOF = False
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i

    If rs.Fields(i - 1).Type = 135 Then
        vGrid.Text = Format((rs.Fields(i - 1).Value), "yyyy/mm/dd")
    Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value)
    End If
    
    If i = 7 Then
        'Carga Informacion de Tesoreria
        If CLng(vGrid.Text) Then
          vGrid.TextTip = TextTipFloating
          vGrid.CellNote = fxTesoreriaConsulta(CLng(vGrid.Text))
        End If
    End If
    
    vGrid.Tag = rs!forma_pago
    
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbProgLlenaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass


'Limpia Datos Primero
lblProgFactura.Caption = ""
lblProgProveedor.Caption = ""
lblProgProveedor.Tag = ""
txtProgNPagos = 1
txtProgFrecuencia = 15

vGridCargos.Enabled = False
vGridVence.Enabled = False

'Solo Muestra las Compras con Forma de Pago a Credito

If IsNumeric(txtLineas) Then
    strSQL = "select Top " & txtLineas & " cod_proveedor,cod_Factura,total,CxP_Estado,fecha,tipo"
Else
    strSQL = "select cod_proveedor,cod_Factura,total,CxP_Estado,fecha,Tipo"
End If

strSQL = strSQL & ",fecha_ingreso,Proveedor,forma_pago,cod_divisa,tipo_cambio,Vence,IMPORTE_DIVISA_REAL" _
       & " from vCxP_ProgramacionPago" _
       & " where cod_factura like '%" & txtFactura & "%'"

If IsNumeric(txtProgCodProv) Then
  strSQL = strSQL & " and cod_proveedor = " & txtProgCodProv
End If

If Len(Trim(txtPrgProveedor)) > 0 Then
    strSQL = strSQL & " and Proveedor like '%" & txtPrgProveedor & "%'"
End If

If Len(Trim(txtProgDivisa)) > 0 Then
    strSQL = strSQL & " and cod_divisa like '%" & txtProgDivisa & "%'"
End If


Select Case True
  Case opt.Item(1) 'Pendientes
     strSQL = strSQL & " and cxp_estado = 'P'"
  Case opt.Item(2) 'Generadas
     strSQL = strSQL & " and cxp_estado = 'G'"
  Case opt.Item(0) 'Todas
End Select

'Forma de Pago y Ordenamiento
If cboTipo.Text = "Todas" Then
    strSQL = strSQL & " Order by fecha desc"
Else
    strSQL = strSQL & " and forma_pago = '" & UCase(Mid(cboTipo.Text, 1, 2)) _
           & "' order by fecha desc"
End If

vPaso = True

lswProg.ListItems.Clear
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lswProg.ListItems.Add(, , rs!cod_Factura)
     itmX.SubItems(1) = rs!cod_proveedor
     itmX.SubItems(2) = rs!Proveedor
     itmX.SubItems(3) = Format(rs!Total, "Standard")
     itmX.SubItems(4) = Format(rs!fecha, "yyyy/mm/dd")
     itmX.SubItems(5) = IIf((rs!CxP_Estado = "P"), "Pendiente", "Generada")
     itmX.SubItems(6) = IIf((rs!Tipo = "I"), "Internas", "Compras")
     itmX.SubItems(7) = IIf((rs!forma_pago = "CR"), "CREDITO", "CONTADO")
     itmX.SubItems(8) = rs!cod_Divisa
     itmX.SubItems(9) = rs!TIPO_CAMBIO
     itmX.SubItems(10) = Format(rs!Importe_divisa_real, "Standard")
     itmX.SubItems(11) = Format(rs!Vence, "yyyy/mm/dd")
     
     itmX.Tag = rs!CxP_Estado
     
     If rs!CxP_Estado = "P" Then
       itmX.Bold = True
       itmX.TextBackColor = RGB(252, 243, 207)
     End If
     
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

vPaso = False

Exit Sub

vError:
 Me.MousePointer = vbDefault
 vPaso = False
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If lblProgFactura.Caption = "" Then
 MsgBox "Seleccione Una Factura Primero...", vbExclamation
 tcMain.Item(0).Selected = True
End If

End Sub

Private Sub tcPro_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

If Item.Index = 1 And Len(lblProgProveedor.Tag) > 0 Then
    strSQL = "select P.NPago,P.Cod_Factura,P.Cod_Proveedor,isnull(sum(C.monto),0) as Cargo" _
           & ",P.Monto,(P.monto - isnull(sum(C.monto),0)) as Neto,isnull(P.Tesoreria,0) as Tesoreria" _
           & ",P.fecha_Vencimiento,P.importe_divisa_real,P.cod_divisa,P.tipo_Cambio,P.forma_pago" _
           & " from cxp_pagoProv P left join cxp_pagoProvCargos C on P.npago = C.npago" _
           & " and P.cod_factura = C.cod_factura and P.cod_proveedor = C.cod_proveedor" _
           & " where P.cod_proveedor = " & lblProgProveedor.Tag _
           & " and P.cod_factura = '" & lblProgFactura.Caption _
           & "' group by  P.NPago,P.Cod_Factura,P.Cod_Proveedor,P.Monto,P.Tesoreria,P.fecha_Vencimiento,P.importe_divisa_real,P.cod_divisa,P.tipo_Cambio,P.forma_pago" _
           & " order by P.NPago"
    
    Call sbCargaGridLocal01(vGrid, 11, strSQL)
    vGrid.MaxRows = vGrid.MaxRows - 1
End If
End Sub

Private Sub txtFactura_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyReturn Then sbProgLlenaLsw
End Sub

Private Sub txtPrgProveedor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyReturn Then sbProgLlenaLsw
End Sub

Private Sub txtProgCodProv_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyReturn Then sbProgLlenaLsw
End Sub


Private Sub txtProgDivisa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyReturn Then sbProgLlenaLsw
End Sub


Private Sub txtProgFrecuencia_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbPagos_Calcula
End Sub

Private Sub txtProgMonto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyReturn Then sbProgLlenaLsw
End Sub


Private Sub txtProgNPagos_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbPagos_Calcula
End Sub

Private Sub vGrid_DblClick(ByVal col As Long, ByVal Row As Long)

If col = 7 Then
  vGrid.Row = Row
  vGrid.col = col
    
  

    If Not IsNumeric(vGrid.Text) Then Exit Sub
    If CCur(vGrid.Text) <= 0 Then Exit Sub
     
    Dim frm As Form
     
    
     Call sbFormsCall("frmTES_Transacciones")
     For Each frm In Forms
       If UCase(frm.Name) = UCase("frmTES_Transacciones") Then
         Call frm.sbTESDocConsulta(vGrid.Text)
         Exit For
       End If
     Next frm
 
End If


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 7
  
  If vGrid.Tag = "CO" Then Exit Sub
  If CInt(vGrid.Text) > 0 Then Exit Sub
  
  vGrid.col = 8
  strSQL = "update cxp_pagoprov set fecha_vencimiento = '" & vGrid.Text _
         & "' where npago = "
  vGrid.col = 1
  strSQL = strSQL & vGrid.Text & " and cod_factura = '"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "' and cod_proveedor = "
  vGrid.col = 3
  strSQL = strSQL & vGrid.Text
  
  Call ConectionExecute(strSQL)
  
  vGrid.col = 8
  MsgBox "Fecha de Vencimiento Cambiada a: " & vGrid.Text
  

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub vGridCargos_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim curMonto As Currency, i As Integer

On Error GoTo vError

curMonto = 0
With vGridCargos
  For i = 1 To .MaxRows
     .Row = i
     .col = 3
     curMonto = curMonto + CCur(.Text)
  Next i
End With

txtTotalCargos.Text = Format(curMonto, "Standard")
txtPagoMin.Text = Format(curMonto / CInt(txtProgNPagos.Text), "Standard")
Exit Sub

vError:
txtTotalCargos.Text = Format(curMonto, "Standard")

End Sub

Private Sub vGridVence_Change(ByVal col As Long, ByVal Row As Long)

Call sbPagos_Calcula_Cambio(col)

End Sub
