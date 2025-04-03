VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCajas_Servicios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cajas: Conceptos y Servicios"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6015
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   8535
      _Version        =   1572864
      _ExtentX        =   15055
      _ExtentY        =   10610
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
      Item(0).Caption =   "Conceptos"
      Item(0).ControlCount=   15
      Item(0).Control(0)=   "Label1(8)"
      Item(0).Control(1)=   "Label1(12)"
      Item(0).Control(2)=   "Label1(14)"
      Item(0).Control(3)=   "Label1(15)"
      Item(0).Control(4)=   "txtDescripcion"
      Item(0).Control(5)=   "txtContrato"
      Item(0).Control(6)=   "dtpFechaVence"
      Item(0).Control(7)=   "chkVence"
      Item(0).Control(8)=   "cboConcepto"
      Item(0).Control(9)=   "GroupBox1"
      Item(0).Control(10)=   "chkIntercambio"
      Item(0).Control(11)=   "chkConfirmaFondos"
      Item(0).Control(12)=   "txtCabys"
      Item(0).Control(13)=   "Label1(1)"
      Item(0).Control(14)=   "chkFactura"
      Item(1).Caption =   "Comisiones"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Item(2).Caption =   "Cajas vinculadas"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "lswCajas"
      Item(2).Control(1)=   "lblServicio"
      Begin XtremeSuiteControls.ListView lswCajas 
         Height          =   5655
         Left            =   -67000
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   5535
         _Version        =   1572864
         _ExtentX        =   9763
         _ExtentY        =   9975
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
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         DragMode        =   1  'Automatic
         Height          =   2895
         Left            =   120
         TabIndex        =   16
         Top             =   3360
         Width           =   8295
         _Version        =   1572864
         _ExtentX        =   14626
         _ExtentY        =   5101
         _StockProps     =   79
         Caption         =   "Configuración Contable"
         ForeColor       =   8421504
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
         Appearance      =   21
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtUnidadCod 
            Height          =   312
            Left            =   1560
            TabIndex        =   23
            Top             =   480
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCentroCod 
            Height          =   312
            Left            =   1560
            TabIndex        =   24
            Top             =   840
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuenta 
            Height          =   312
            Left            =   1560
            TabIndex        =   25
            Top             =   1440
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaComision 
            Height          =   312
            Left            =   1560
            TabIndex        =   26
            Top             =   1800
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaIV 
            Height          =   312
            Left            =   1560
            TabIndex        =   27
            Top             =   2160
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtUnidadDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   28
            Top             =   480
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCentroDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   29
            Top             =   840
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   30
            Top             =   1440
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaComisionDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   31
            Top             =   1800
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaIVDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   32
            Top             =   2160
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Comisión"
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
            Left            =   0
            TabIndex        =   22
            Top             =   1800
            Width           =   1692
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta I.V.A"
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
            Left            =   0
            TabIndex        =   21
            Top             =   2160
            Width           =   1692
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Principal"
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
            Left            =   0
            TabIndex        =   20
            Top             =   1440
            Width           =   1692
         End
         Begin VB.Label Label1 
            Caption         =   "Centro de Costo"
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
            Left            =   0
            TabIndex        =   19
            Top             =   840
            Width           =   1692
         End
         Begin VB.Label Label1 
            Caption         =   "Unidad"
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
            Left            =   0
            TabIndex        =   18
            Top             =   480
            Width           =   1692
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   6732
         _Version        =   1572864
         _ExtentX        =   11874
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtContrato 
         Height          =   312
         Left            =   1680
         TabIndex        =   12
         Top             =   840
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaVence 
         Height          =   312
         Left            =   4920
         TabIndex        =   13
         Top             =   840
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.CheckBox chkVence 
         Height          =   252
         Left            =   6600
         TabIndex        =   14
         Top             =   840
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Vence?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboConcepto 
         Height          =   312
         Left            =   1680
         TabIndex        =   15
         Top             =   1320
         Width           =   6732
         _Version        =   1572864
         _ExtentX        =   11880
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
      Begin XtremeSuiteControls.CheckBox chkIntercambio 
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   1800
         Width           =   6735
         _Version        =   1572864
         _ExtentX        =   11874
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Utilizar para Intercambio de valores por Efectivo en divisa local?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5535
         Left            =   -69280
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   7335
         _Version        =   524288
         _ExtentX        =   12938
         _ExtentY        =   9763
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
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frmCajas_Servicios.frx":0000
         VisibleRows     =   1
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.CheckBox chkConfirmaFondos 
         Height          =   255
         Left            =   1680
         TabIndex        =   36
         Top             =   2160
         Width           =   6735
         _Version        =   1572864
         _ExtentX        =   11874
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Valor en Tránsito, Requiere Proceso de Confirmación de Fondos?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtCabys 
         Height          =   315
         Left            =   1680
         TabIndex        =   37
         Top             =   2640
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3408
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkFactura 
         Height          =   255
         Left            =   3840
         TabIndex        =   39
         Top             =   2640
         Width           =   4695
         _Version        =   1572864
         _ExtentX        =   8281
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Genera Facturación?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin VB.Label Label1 
         Caption         =   "Código CABYS"
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
         Left            =   120
         TabIndex        =   38
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblServicio 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto/Servicios asigando a las siguentes cajas ..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   912
         Left            =   -69760
         TabIndex        =   34
         Top             =   720
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
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
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2052
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Contrato"
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1932
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   216
         Index           =   12
         Left            =   3720
         TabIndex        =   8
         Top             =   840
         Width           =   1200
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
         Index           =   8
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1932
      End
   End
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   252
      Left            =   3360
      TabIndex        =   4
      Top             =   480
      Width           =   1452
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activo?"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar2 
      Height          =   252
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgExplorer"
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
   Begin MSComctlLib.ImageList imgExplorer 
      Left            =   5880
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Servicios.frx":0788
            Key             =   "imgDocu"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Servicios.frx":1662
            Key             =   "imgFormu"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7125
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Usuario de Registro"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
      EndProperty
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   1452
      _Version        =   1572864
      _ExtentX        =   2561
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   216
      Index           =   0
      Left            =   312
      TabIndex        =   1
      Top             =   480
      Width           =   708
   End
End
Attribute VB_Name = "frmCajas_Servicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tRecaudador
    Codigo      As String
    Nombre      As String
    Cta         As String
    CtaDesc     As String
    CtaIV       As String
    CtaIVDesc   As String
    CtaCom      As String
    CtaComDesc  As String
    Vence       As Date
End Type

Dim mRecaudador As tRecaudador
Dim vScroll As Boolean, vPaso As Boolean
Dim vEdita  As Boolean, vCodigo As String



Private Sub cboConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuenta.SetFocus
End Sub

Private Sub dtpFechaVence_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboConcepto.SetFocus
End Sub

Private Sub FlatScrollBar2_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
'tcMain.Item(0).Selected = True
If vScroll Then
 If txtCodigo = "" Then txtCodigo = 0
    strSQL = "select Top 1 cod_servicio from cajas_servicios"
    
    If FlatScrollBar2.Value = 1 Then
       strSQL = strSQL & " where cod_recaudador = '" & mRecaudador.Codigo & "' and cod_servicio > '" & txtCodigo & "' order by cod_servicio asc"  ' and cod_recaudador = '" & mRecaudador.Codigo & "' order by cod_servicio asc"
    Else
       strSQL = strSQL & " where cod_recaudador = '" & mRecaudador.Codigo & "' and cod_servicio < '" & txtCodigo & "' order by cod_servicio desc" ' and cod_recaudador = '" & mRecaudador.Codigo & "' order by cod_servicio desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!COD_SERVICIO)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar2.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Activate()
vModulo = 5

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 5

vEdita = False

With lswCajas.ColumnHeaders
    .Clear
    .Add , , "", 4000
End With

strSQL = "select rtrim(cod_concepto) as 'IdX' , rtrim(descripcion)  as 'itmx'" _
       & " from sif_conceptos where activo = 1  and cod_concepto like 'CAJ%' order by cod_Concepto"
Call sbCbo_Llena_New(cboConcepto, strSQL, False)


'Carga datos del Recaudador
strSQL = "select R.*, dateadd(yyyy, 1,dbo.MyGetdate()) as 'Vence', isnull(Cta.Descripcion,'') as 'CtaDesc',isnull(CtaIv.Descripcion,'') as 'CtaDescIv', isnull(CtaCom.Descripcion,'') as 'CtaDescCom'" _
       & " from cajas_recaudador R left join CntX_Cuentas Cta on R.cod_Cuenta = Cta.Cod_Cuenta and Cta.Cod_Contabilidad = " & GLOBALES.gEnlace _
       & " left join CntX_Cuentas CtaIv on R.cod_Cuenta_Iv = CtaIv.Cod_Cuenta and CtaIv.Cod_Contabilidad = " & GLOBALES.gEnlace _
       & " left join CntX_Cuentas CtaCom on R.cod_Cuenta_Comision = CtaCom.Cod_Cuenta and CtaCom.Cod_Contabilidad = " & GLOBALES.gEnlace _
       & " where R.cod_recaudador = '" & GLOBALES.gTag & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  mRecaudador.Codigo = Trim(rs!COD_RECAUDADOR)
  mRecaudador.Nombre = Trim(rs!Descripcion)
  
  mRecaudador.Cta = Trim(rs!cod_cuenta)
  mRecaudador.CtaDesc = Trim(rs!CtaDesc)
  mRecaudador.CtaIV = Trim(rs!Cod_Cuenta_IV)
  mRecaudador.CtaIVDesc = Trim(rs!CtaDescIV)
  mRecaudador.CtaCom = Trim(rs!Cod_Cuenta_Comision)
  mRecaudador.CtaComDesc = Trim(rs!CtaDescCom)
  mRecaudador.Vence = rs!Vence
End If
rs.Close

Me.Caption = "Servicios del Recaudador .:" & mRecaudador.Nombre

Call sbLimpia

Call sbToolBarIconos(tlb, False)
Call sbToolBar(tlb, "nuevo")

vScroll = False
    FlatScrollBar2.Value = 0
vScroll = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbLimpia()

vCodigo = ""

txtDescripcion.Text = ""
txtContrato.Text = ""


txtCuenta.Text = mRecaudador.Cta
txtCuentaDesc.Text = mRecaudador.CtaDesc

txtCuentaComision.Text = mRecaudador.CtaCom
txtCuentaComisionDesc = mRecaudador.CtaComDesc

txtCuentaIV.Text = mRecaudador.CtaIV
txtCuentaIVDesc.Text = mRecaudador.CtaIVDesc

chkVence.Value = vbChecked

chkActivo.Value = vbChecked
chkIntercambio.Value = vbUnchecked

dtpFechaVence.Value = mRecaudador.Vence

StatusBarX.Panels.Item(1).Text = ""
StatusBarX.Panels.Item(2).Text = ""


tcMain.Item(0).Selected = True
tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False

End Sub


Private Sub lswCajas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert CAJAS_SERVICIOS_ASIGNADOS(cod_recaudador,cod_servicio,cod_Caja,registro_Fecha,registro_usuario)" _
          & " values('" & mRecaudador.Codigo & "','" & vCodigo & "','" & Item.Tag & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"

   Call Bitacora("Registra", "Asignación en Caja: " & Item.Tag & " (Serv.:" & vCodigo & " - Rec.:" & mRecaudador.Codigo & ")")
Else
   strSQL = "delete CAJAS_SERVICIOS_ASIGNADOS where cod_recaudador = '" & mRecaudador.Codigo & "' and cod_servicio = '" _
          & vCodigo & "' and cod_caja = '" & Item.Tag & "'"
   Call Bitacora("Elimina", "Asignación en Caja: " & Item.Tag & " (Serv.:" & vCodigo & " - Rec.:" & mRecaudador.Codigo & ")")
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vCodigo = "" Then Exit Sub

Select Case Item.Index
  Case 1 'Rangos de Comision
    strSQL = "Select linea,monto_inicio,monto_corte,comision_mnt_minimo,comision_porcentaje,iv_porcentaje from cajas_servicios_rangos " _
            & " where cod_servicio = '" & vCodigo & "' and cod_recaudador = '" & mRecaudador.Codigo _
            & "' order by Linea"
    Call sbCargaGrid(vGrid, 6, strSQL)
  
  Case 2 'Cajas Asignadas

    strSQL = "select C.cod_caja,C.descripcion,X.cod_caja as 'Asignado'" _
            & " from cajas_definicion C left join cajas_servicios_asignados X on C.cod_caja = X.cod_caja" _
            & " and X.cod_recaudador = '" & mRecaudador.Codigo & "' and X.cod_servicio = '" & vCodigo & "'"
    
    vPaso = True
    lswCajas.ListItems.Clear
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = lswCajas.ListItems.Add(, , Trim(rs!Descripcion))
          itmX.Tag = Trim(rs!COD_CAJA)
      
      If Not IsNull(rs!Asignado) Then
         itmX.Checked = True
      End If
          
      rs.MoveNext
    Loop
    rs.Close
    
    vPaso = False
  
End Select

End Sub

Private Sub txtCentroCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentroDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_centro_Costo"
   gBusquedas.Orden = "cod_centro_Costo"
   gBusquedas.Consulta = "select cod_centro_Costo,descripcion from Cntx_Centro_Costos"
   gBusquedas.Filtro = " and cod_Contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   If Len(gBusquedas.Resultado) > 0 Then
        txtCentroCod.Text = gBusquedas.Resultado
        txtCentroDesc.Text = gBusquedas.Resultado2
   End If
End If

End Sub

Private Sub txtCentroCod_LostFocus()
txtCentroDesc.Text = fxgCntCentroCostos(txtCentroCod.Text)
End Sub


Private Sub txtCodigo_Change()
Call sbLimpia
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_servicio"
   gBusquedas.Orden = "cod_servicio"
   gBusquedas.Consulta = "select cod_servicio,descripcion from cajas_servicios"
   gBusquedas.Filtro = " and cod_recaudador = '" & mRecaudador.Codigo & "'"
   frmBusquedas.Show vbModal
   txtCodigo.SetFocus
   txtCodigo.Text = gBusquedas.Resultado
   txtDescripcion.SetFocus
End If

End Sub

Private Sub txtCodigo_LostFocus()
If Trim(txtCodigo) <> "" Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtContrato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFechaVence.SetFocus
End Sub


Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuenta.Text = gCuenta
    txtCuentaDesc.Text = ""
End If

End Sub

Private Sub txtCuenta_LostFocus()
    txtCuenta.Text = fxgCntCuentaFormato(False, txtCuenta.Text)
    txtCuentaDesc.Text = fxgCntCuentaDesc(txtCuenta.Text)
    txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta.Text)
End Sub


Private Sub txtCuentaComision_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaComisionDesc.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaComision.Text = gCuenta
    txtCuentaComisionDesc.Text = ""
End If

End Sub


Private Sub txtCuentaComision_LostFocus()
    txtCuentaComision.Text = fxgCntCuentaFormato(False, txtCuentaComision.Text)
    txtCuentaComisionDesc.Text = fxgCntCuentaDesc(txtCuentaComision.Text)
    txtCuentaComision.Text = fxgCntCuentaFormato(True, txtCuentaComision.Text)
End Sub

Private Sub txtCuentaComisionDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaIV.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaComisionDesc.Text = fxgCntCuentaDesc(gCuenta)
    txtCuentaComision.Text = fxgCntCuentaFormato(True, gCuenta)
End If

End Sub

Private Sub txtCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaComision.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
    txtCuenta.Text = fxgCntCuentaFormato(True, gCuenta)
End If

End Sub

Private Sub txtCuentaIV_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaIVDesc.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaIV.Text = gCuenta
    txtCuentaIVDesc.Text = ""
End If

End Sub

Private Sub txtCuentaIV_LostFocus()
    txtCuentaIV.Text = fxgCntCuentaFormato(False, txtCuentaIV.Text)
    txtCuentaIVDesc.Text = fxgCntCuentaDesc(txtCuentaIV.Text)
    txtCuentaIV.Text = fxgCntCuentaFormato(True, txtCuentaIV.Text)
End Sub

Private Sub txtCuentaIVDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuentaIVDesc.Text = fxgCntCuentaDesc(gCuenta)
    txtCuentaIV.Text = fxgCntCuentaFormato(True, gCuenta)
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContrato.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "select cod_servicio,descripcion from cajas_servicios"
   gBusquedas.Filtro = " and cod_recaudador = '" & mRecaudador.Codigo & "'"
   frmBusquedas.Show vbModal
   txtCodigo.SetFocus
   txtCodigo.Text = gBusquedas.Resultado
   Call txtCodigo_LostFocus
End If

End Sub

Private Sub txtUnidadCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidadDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_unidad"
   gBusquedas.Orden = "cod_unidad"
   gBusquedas.Consulta = "select cod_unidad,descripcion from Cntx_Unidades"
   gBusquedas.Filtro = " and cod_Contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   If Len(gBusquedas.Resultado) > 0 Then
        txtUnidadCod.Text = gBusquedas.Resultado
        txtUnidadDesc.Text = gBusquedas.Resultado2
   End If
End If
End Sub


Private Sub txtUnidadCod_LostFocus()
txtUnidadDesc.Text = fxgCntUnidad(txtUnidadCod.Text)
End Sub

Private Sub txtUnidadDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or vbKeyTab Then txtCentroCod.SetFocus
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If


If KeyCode = vbKeyDelete Then
   'Aqui codigo de Borrado
   If MsgBox("¿Desea eliminar esta linea?", vbYesNo Or vbQuestion, "") = vbNo Then Exit Sub
      
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   If Trim(vGrid.Text) <> "" Then
      vGrid.Col = 1
        strSQL = "Delete cajas_servicios_rangos where cod_recaudador = '" & mRecaudador.Codigo & "' and linea = '" & vGrid.Text & "' " _
                & " and cod_servicio = '" & vCodigo & "' "
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Elimina", "Rango Servicio..: " & vCodigo & " .. Recaudador.:" & mRecaudador.Codigo)

   End If
   vGrid.DeleteRows vGrid.ActiveRow, 1
   vGrid.MaxRows = vGrid.MaxRows - 1
End If


If KeyCode = vbKeyInsert Then
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.InsertRows vGrid.ActiveRow, 1
  vGrid.Row = vGrid.ActiveRow
End If



End Sub



Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from cajas_servicios_rangos" _
       & " where linea =  " & vGrid.ActiveRow & " and cod_servicio = '" & vCodigo & "' and" _
       & " cod_recaudador = '" & mRecaudador.Codigo & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
    strSQL = "insert into cajas_servicios_rangos(cod_servicio,cod_recaudador,linea,monto_inicio,monto_corte" _
             & ",comision_mnt_minimo,comision_porcentaje,iv_porcentaje)" _
           & " values('" & vCodigo & "','" & mRecaudador.Codigo & "',"
    vGrid.Col = 2
    strSQL = strSQL & " " & vGrid.ActiveRow & ", " & CCur(vGrid.Text) & ","
    vGrid.Col = 3
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.Col = 4
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.Col = 5
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.Col = 6
    strSQL = strSQL & CCur(vGrid.Text) & ")"
    Call ConectionExecute(strSQL)
    
    
    Call Bitacora("Registra", "Rango Servicio..: " & vCodigo & " .. Recaudador.:" & mRecaudador.Codigo)
    vGrid.Col = 1
    vGrid.Text = CStr(vGrid.ActiveRow)

Else 'Actualizar
    
    vGrid.Col = 2
    strSQL = "update cajas_servicios_rangos set monto_inicio= " & CCur(vGrid.Text) & ","
    vGrid.Col = 3
    strSQL = strSQL & " monto_corte = " & CCur(vGrid.Text) & ","
    vGrid.Col = 4
    strSQL = strSQL & "comision_mnt_minimo = " & CCur(vGrid.Text) & ","
    vGrid.Col = 5
    strSQL = strSQL & "comision_porcentaje = " & CCur(vGrid.Text) & ","
    vGrid.Col = 6
    strSQL = strSQL & "iv_porcentaje = " & CCur(vGrid.Text) & ""
    vGrid.Col = 1
    strSQL = strSQL & " where linea =  " & vGrid.Text & " and cod_servicio = '" & vCodigo & "'"
    strSQL = strSQL & " and cod_recaudador = '" & mRecaudador.Codigo & "'"
    
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Rango Servicio..: " & vCodigo & " .. Recaudador.:" & mRecaudador.Codigo)

End If
rs.Close

fxGuardar = 1


Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function

Private Sub sbConsulta(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.*, Con.Descripcion as 'ConceptoDesc'" _
       & ", isnull(Cta.Descripcion,'') as 'CtaDesc',isnull(CtaIv.Descripcion,'') as 'CtaDescIv', isnull(CtaCom.Descripcion,'') as 'CtaDescCom'" _
       & " from cajas_servicios C inner join sif_conceptos Con on C.COD_CONCEPTO = Con.COD_CONCEPTO" _
       & " left join CntX_Cuentas Cta on C.cod_Cuenta = Cta.Cod_Cuenta and Cta.Cod_Contabilidad = " & GLOBALES.gEnlace _
       & " left join CntX_Cuentas CtaIv on C.cod_Cuenta_Iv = CtaIv.Cod_Cuenta and CtaIv.Cod_Contabilidad = " & GLOBALES.gEnlace _
       & " left join CntX_Cuentas CtaCom on C.cod_Cuenta_Comision = CtaCom.Cod_Cuenta and CtaCom.Cod_Contabilidad = " & GLOBALES.gEnlace _
       & " where C.cod_recaudador = '" & mRecaudador.Codigo & "' and C.cod_servicio = '" & pCodigo & "'"
       
Call OpenRecordSet(rs, strSQL)

Call sbLimpia


If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  txtCodigo.Text = rs!COD_SERVICIO
  vCodigo = Trim(rs!COD_SERVICIO)
  
  txtDescripcion.Text = rs!Descripcion
  txtContrato = rs!Contrato
  dtpFechaVence.Value = rs!vence_fecha
  chkActivo.Value = rs!Activo
  chkVence.Value = rs!vende_activo
  chkIntercambio.Value = rs!InterCambio
  
  Call sbCboAsignaDato(cboConcepto, rs!ConceptoDesc, True, rs!cod_Concepto)
  
  txtCuenta.Text = fxgCntCuentaFormato(True, rs!cod_cuenta)
  txtCuentaDesc.Text = rs!CtaDesc
  
  txtCuentaIV.Text = fxgCntCuentaFormato(True, rs!Cod_Cuenta_IV)
  txtCuentaIVDesc.Text = rs!CtaDescIV
  
  txtCuentaComision.Text = fxgCntCuentaFormato(True, rs!Cod_Cuenta_Comision)
  txtCuentaComisionDesc.Text = rs!CtaDescCom
  
  txtUnidadCod.Text = rs!Cod_Unidad & ""
  txtCentroCod.Text = rs!Cod_Centro_Costo & ""
  
  txtUnidadCod_LostFocus
  txtCentroCod_LostFocus
    
  
  StatusBarX.Panels.Item(1).Text = rs!REGISTRO_USUARIO & ""
  StatusBarX.Panels.Item(2).Text = rs!REGISTRO_FECHA & ""

  tcMain.Item(0).Selected = True
  tcMain.Item(1).Enabled = True
  tcMain.Item(2).Enabled = True

End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
        Call sbLimpia
        vEdita = False
        txtCodigo.SetFocus
        txtCodigo.Text = ""
       Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      'Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
     Call sbToolBar(tlb, "activo")
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpia
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_servicio,descripcion from cajas_servicios "
       gBusquedas.Filtro = " and cod_recaudador = '" & mRecaudador.Codigo & "'"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo.Text = gBusquedas.Resultado
       Call txtCodigo_LostFocus

    Case "REPORTES"

    Case "AYUDA"

End Select


End Sub


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update cajas_servicios set descripcion = '" & Trim(txtDescripcion) & "'" _
         & ",activo = " & chkActivo.Value & ",contrato = '" & txtContrato.Text & "'" _
         & ",vence_fecha = '" & Format(dtpFechaVence, "yyyy/mm/dd") _
         & "',vende_activo = " & chkVence.Value & ", Intercambio = " & chkIntercambio.Value _
         & ",cod_cuenta_comision = '" & fxgCntCuentaFormato(False, txtCuentaComision) & "'" _
         & ",cod_cuenta_iv = '" & fxgCntCuentaFormato(False, txtCuentaIV) & "', cod_concepto = '" & cboConcepto.ItemData(cboConcepto.ListIndex) & "'" _
         & ",cod_cuenta = '" & fxgCntCuentaFormato(False, txtCuenta) _
         & "',cod_Unidad = '" & txtUnidadCod.Text & "', cod_centro_costo = '" & txtCentroCod.Text & "'" _
         & " where cod_servicio = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Servicio: " & vCodigo)

Else
  vCodigo = txtCodigo

   strSQL = "insert into cajas_servicios(cod_recaudador,cod_servicio,descripcion,activo,contrato,vence_fecha,vende_activo" _
          & ",cod_Cuenta,cod_cuenta_comision,cod_cuenta_iv,cod_concepto,Intercambio,cod_unidad,cod_centro_costo, REGISTRO_USUARIO,REGISTRO_FECHA)" _
          & " values('" & mRecaudador.Codigo & "','" & vCodigo & "','" & Trim(txtDescripcion.Text) & "', " & chkActivo.Value & ",'" & txtContrato.Text & "'," _
          & "'" & Format(dtpFechaVence, "yyyy/mm/dd") & "'," & chkVence.Value & ",'" & fxgCntCuentaFormato(False, txtCuenta.Text) & "'," _
          & "'" & fxgCntCuentaFormato(False, txtCuentaComision) & "','" & fxgCntCuentaFormato(False, txtCuentaIV) & "','" _
          & cboConcepto.ItemData(cboConcepto.ListIndex) & "'," & chkIntercambio.Value & ",'" & txtUnidadCod.Text & "','" _
          & txtCentroCod.Text & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Servicio: " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValida()
Dim vMensaje As String

vMensaje = ""
fxValida = True


If Trim(txtCodigo) = "" Then vMensaje = vMensaje & vbCrLf & " - Código del Servicio no es válido ..."
If Trim(txtDescripcion) = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Servicio no es válido ..."

If Not fxgCntCuentaValida(txtCuenta.Text) Then vMensaje = vMensaje & vbCrLf & " - Cuenta Contable Prinicipal no es válida.."
If Not fxgCntCuentaValida(txtCuentaIV.Text) Then vMensaje = vMensaje & vbCrLf & " - Cuenta Contable para Impuesto de Ventas no es válida.."
If Not fxgCntCuentaValida(txtCuentaComision.Text) Then vMensaje = vMensaje & vbCrLf & " - Cuenta Contable para Comisiones no es válida.."

If Trim(txtUnidadDesc.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - La Unidad Contable no es válida ..."
If Trim(txtCentroCod.Text) <> "" And Trim(txtCentroDesc.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Centro de Costos no es válido ..."



If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If


End Function

