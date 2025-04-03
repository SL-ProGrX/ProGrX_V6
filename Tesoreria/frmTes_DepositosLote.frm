VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmTes_DepositosLote 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Carga de Depósitos en Cuenta (Archivo/Lote)"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   12090
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   11895
      _Version        =   1572864
      _ExtentX        =   20976
      _ExtentY        =   12933
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
      SelectedItem    =   1
      Item(0).Caption =   "Cargado"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "txtArchivo"
      Item(0).Control(1)=   "Label1(2)"
      Item(0).Control(2)=   "txtCuentaDesc"
      Item(0).Control(3)=   "txtCuenta"
      Item(0).Control(4)=   "cboCategoria"
      Item(0).Control(5)=   "vGrid"
      Item(0).Control(6)=   "Label2(10)"
      Item(0).Control(7)=   "Label1(1)"
      Item(0).Control(8)=   "fraAccion"
      Item(0).Control(9)=   "btnArchivo(0)"
      Item(0).Control(10)=   "btnArchivo(1)"
      Item(0).Control(11)=   "btnArchivo(2)"
      Item(1).Caption =   "Registro en Bancos"
      Item(1).ControlCount=   19
      Item(1).Control(0)=   "btnRegistro_Buscar"
      Item(1).Control(1)=   "cboFiltro"
      Item(1).Control(2)=   "cboFechas"
      Item(1).Control(3)=   "txtNombre"
      Item(1).Control(4)=   "txtNumDoc"
      Item(1).Control(5)=   "chkMarcas"
      Item(1).Control(6)=   "vGridId"
      Item(1).Control(7)=   "btnRegistro_Registrar"
      Item(1).Control(8)=   "btnRegistro_Actualizar"
      Item(1).Control(9)=   "btnRegistro_Desvincular"
      Item(1).Control(10)=   "Label2(9)"
      Item(1).Control(11)=   "Label2(8)"
      Item(1).Control(12)=   "Label2(7)"
      Item(1).Control(13)=   "Label2(4)"
      Item(1).Control(14)=   "Label2(5)"
      Item(1).Control(15)=   "dtpRegistroInicio"
      Item(1).Control(16)=   "dtpRegistroCorte"
      Item(1).Control(17)=   "txtCedula"
      Item(1).Control(18)=   "Label2(6)"
      Item(2).Caption =   "Inconsistencias"
      Item(2).ControlCount=   6
      Item(2).Control(0)=   "vGridInco"
      Item(2).Control(1)=   "btnInco_Buscar"
      Item(2).Control(2)=   "btnInco_Exportar"
      Item(2).Control(3)=   "Label2(11)"
      Item(2).Control(4)=   "dtpIncoInicio"
      Item(2).Control(5)=   "dtpIncoCorte"
      Begin XtremeSuiteControls.CheckBox chkMarcas 
         Height          =   252
         Left            =   720
         TabIndex        =   39
         Top             =   1560
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Marcar"
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
         Appearance      =   16
      End
      Begin VB.Frame fraAccion 
         BorderStyle     =   0  'None
         Height          =   732
         Left            =   -69880
         TabIndex        =   25
         Top             =   6600
         Visible         =   0   'False
         Width           =   11532
         Begin XtremeSuiteControls.PushButton btnAplicar 
            Height          =   492
            Left            =   9840
            TabIndex        =   26
            Top             =   120
            Width           =   1452
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmTes_DepositosLote.frx":0000
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.PushButton btnCancelar 
            Height          =   492
            Left            =   8400
            TabIndex        =   27
            Top             =   120
            Width           =   1452
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Cancelar"
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmTes_DepositosLote.frx":0727
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   312
            Left            =   960
            TabIndex        =   43
            Top             =   360
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
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
         Begin XtremeSuiteControls.FlatEdit txtCasos 
            Height          =   312
            Left            =   2520
            TabIndex        =   44
            Top             =   360
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
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
         Begin XtremeSuiteControls.FlatEdit txtSocios 
            Height          =   312
            Left            =   3480
            TabIndex        =   45
            Top             =   360
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
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
         Begin XtremeSuiteControls.FlatEdit txtContratos 
            Height          =   312
            Left            =   4440
            TabIndex        =   46
            Top             =   360
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
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
         Begin VB.Label Label2 
            Caption         =   "Totales"
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
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Casos"
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
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   30
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Existe ?"
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
            Height          =   255
            Index           =   2
            Left            =   3480
            TabIndex        =   29
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Ident.?"
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
            Height          =   255
            Index           =   3
            Left            =   4440
            TabIndex        =   28
            Top             =   120
            Width           =   975
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4692
         Left            =   -69880
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   11652
         _Version        =   524288
         _ExtentX        =   20553
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
         MaxCols         =   7
         SpreadDesigner  =   "frmTes_DepositosLote.frx":0CCB
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnRegistro_Buscar 
         Height          =   492
         Left            =   8520
         TabIndex        =   6
         Top             =   480
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Buscar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTes_DepositosLote.frx":143B
         ImageAlignment  =   4
      End
      Begin FPSpreadADO.fpSpread vGridId 
         Height          =   5175
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   11535
         _Version        =   524288
         _ExtentX        =   20346
         _ExtentY        =   9128
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
         MaxCols         =   15
         SpreadDesigner  =   "frmTes_DepositosLote.frx":1B3B
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnRegistro_Registrar 
         Height          =   492
         Left            =   10200
         TabIndex        =   8
         Top             =   480
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Registar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTes_DepositosLote.frx":25E2
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnRegistro_Actualizar 
         Height          =   492
         Left            =   8520
         TabIndex        =   9
         Top             =   1080
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Actualizar Identificación"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTes_DepositosLote.frx":2D09
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnRegistro_Desvincular 
         Height          =   492
         Left            =   10200
         TabIndex        =   10
         Top             =   1080
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Desvincular"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTes_DepositosLote.frx":3409
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.DateTimePicker dtpRegistroInicio 
         Height          =   312
         Left            =   1080
         TabIndex        =   16
         Top             =   1200
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
      Begin XtremeSuiteControls.DateTimePicker dtpRegistroCorte 
         Height          =   312
         Left            =   2400
         TabIndex        =   17
         Top             =   1200
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
      Begin FPSpreadADO.fpSpread vGridInco 
         Height          =   5892
         Left            =   -69880
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   11652
         _Version        =   524288
         _ExtentX        =   20553
         _ExtentY        =   10393
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
         MaxCols         =   8
         SpreadDesigner  =   "frmTes_DepositosLote.frx":39AD
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnInco_Buscar 
         Height          =   492
         Left            =   -66160
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Buscar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTes_DepositosLote.frx":408B
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnInco_Exportar 
         Height          =   492
         Left            =   -64840
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Exportar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTes_DepositosLote.frx":478B
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.DateTimePicker dtpIncoInicio 
         Height          =   312
         Left            =   -69040
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.DateTimePicker dtpIncoCorte 
         Height          =   312
         Left            =   -67720
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.ComboBox cboCategoria 
         Height          =   312
         Left            =   -68680
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.ComboBox cboFechas 
         Height          =   312
         Left            =   1080
         TabIndex        =   34
         Top             =   840
         Width           =   2652
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.ComboBox cboFiltro 
         Height          =   312
         Left            =   5280
         TabIndex        =   35
         Top             =   1200
         Width           =   3132
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.FlatEdit txtNumDoc 
         Height          =   312
         Left            =   1080
         TabIndex        =   36
         Top             =   480
         Width           =   2652
         _Version        =   1572864
         _ExtentX        =   4678
         _ExtentY        =   550
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   312
         Left            =   5280
         TabIndex        =   37
         Top             =   480
         Width           =   3132
         _Version        =   1572864
         _ExtentX        =   5524
         _ExtentY        =   550
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   5280
         TabIndex        =   38
         Top             =   840
         Width           =   3132
         _Version        =   1572864
         _ExtentX        =   5524
         _ExtentY        =   550
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuenta 
         Height          =   312
         Left            =   -65080
         TabIndex        =   40
         Top             =   1200
         Visible         =   0   'False
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
         Height          =   312
         Left            =   -63280
         TabIndex        =   41
         Top             =   1200
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   672
         Left            =   -68680
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   8772
         _Version        =   1572864
         _ExtentX        =   15473
         _ExtentY        =   1185
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   372
         Index           =   0
         Left            =   -59680
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTes_DepositosLote.frx":505C
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   372
         Index           =   1
         Left            =   -59200
         TabIndex        =   48
         Top             =   600
         Visible         =   0   'False
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTes_DepositosLote.frx":575C
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   372
         Index           =   2
         Left            =   -58720
         TabIndex        =   49
         Top             =   600
         Visible         =   0   'False
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTes_DepositosLote.frx":5E75
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha .:"
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
         Left            =   -69760
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Identificación.:"
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
         Left            =   3960
         TabIndex        =   18
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "No. Doc.:"
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
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre.:"
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
         Left            =   3960
         TabIndex        =   14
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha .:"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Filtro.:"
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
         Left            =   3960
         TabIndex        =   12
         Top             =   1200
         Width           =   852
      End
      Begin VB.Label Label2 
         Caption         =   "Base .:"
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
         TabIndex        =   11
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Categoría"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   -69760
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta.:"
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
         Left            =   -65920
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   -69760
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
      End
   End
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   312
      Left            =   2400
      TabIndex        =   32
      Top             =   240
      Width           =   7692
      _Version        =   1572864
      _ExtentX        =   13573
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      DataField       =   "Banco"
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
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTes_DepositosLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mBanco As Long, vPaso As Boolean

Private Sub sbLimpia()

    vGrid.MaxRows = 0
    vGridId.MaxRows = 0
    vGridInco.MaxRows = 0
    
'    tlbRegistro.Visible = IIf((ssTab.Tab = 1), True, False)
    
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtSocios.Text = 0
    txtContratos.Text = 0
    txtArchivo.Text = ""
End Sub



Private Sub btnAplicar_Click()
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen registros cargados...[verifique!]", vbExclamation
       Exit Sub
    End If
   
    Call sbProcesar
End Sub

Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String
  
Select Case Index
  
  Case 0 'buscar
        txtArchivo.Text = ""
        Call sbArchivoBusca

  Case 1 'cargar
       Call sbArchivoCarga


  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: Import" & vbCrLf _
              & " 3. Columnas.: DOCUMENTO, FECHA, MONTO, DESCRIPCION"
     
     MsgBox vMensaje, vbInformation
     
     
End Select
End Sub

Private Sub btnCancelar_Click()
    vGrid.MaxRows = 0
    txtArchivo.Text = ""
End Sub

Private Sub btnInco_Buscar_Click()
Call sbInconsistenciaBuscar
End Sub

Private Sub btnInco_Exportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 8
    vHeaders.Headers(1) = "# Documento"
    vHeaders.Headers(2) = "Monto"
    vHeaders.Headers(3) = "Fec.Doc."
    vHeaders.Headers(4) = "Descripcion"
    vHeaders.Headers(5) = "Inconsistencia"
    vHeaders.Headers(6) = "Registro.Fecha"
    vHeaders.Headers(7) = "Registro.Usuario"
    vHeaders.Headers(8) = "Banco"
Call sbSIFGridExportar(vGridInco, vHeaders, "Tes_ControlDepositos")


'Select Case ButtonMenu.Key
'  Case "Excel"
'      Call sbSIFGridExportar(vGridInco, vHeaders, "Tes_ControlDepositos")
'  Case "HTML"
'      Call sbSIFGridExportar(vGridInco, vHeaders, "Tes_ControlDepositos", "HTML")
'End Select
End Sub

Private Sub btnRegistro_Actualizar_Click()
    Call sbRegistroActualizar
End Sub

Private Sub btnRegistro_Buscar_Click()
    Call sbRegistroBuscar
End Sub

Private Sub btnRegistro_Desvincular_Click()
Dim strSQL As String, i As Long
Dim pDocumento As String, pCedula As String

On Error GoTo vError

Me.MousePointer = vbHourglass

With vGridId

strSQL = ""
For i = 1 To .MaxRows
  .Row = i
  .Col = 1
  If .Value = vbChecked Then
    .Col = 3
    pDocumento = .Text
    .Col = 9
    pCedula = .Text
    strSQL = strSQL & Space(10) & "exec spTES_Deposito_Desvincula " & mBanco & ",'" & pDocumento & "','" & pCedula & "','" & glogon.Usuario & "'"
  End If

  If Len(strSQL) > 25000 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
  End If
  

Next i
End With

  If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
  End If


Me.MousePointer = vbDefault

MsgBox "Casos registrados en Tesorería satisfactoriamente!", vbInformation

Call sbRegistroBuscar


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Call sbRegistroBuscar

End Sub

Private Sub btnRegistro_Registrar_Click()
    Select Case cboFiltro.Text
       Case "Identificados - No Registrados"
            Call sbRegistroAplicar
       Case "Identificados - Registrados"
            MsgBox "Los casos actuales ya fueron procesados!", vbInformation
       Case "No Identificados - Registrados"
            MsgBox "Los casos actuales ya fueron procesados!", vbInformation
       Case "No Identificados - No Registrados"
            Call sbRegistroAplicar
    End Select
End Sub

Private Sub cboBanco_Click()

If vPaso Then Exit Sub

 Call sbLimpia
 
If cboBanco.ListCount = 0 Then
 mBanco = 0
Else
 mBanco = cboBanco.ItemData(cboBanco.ListIndex)
End If

End Sub



Private Sub cboCategoria_Click()
Dim vCuenta As String

If vPaso Then Exit Sub


Select Case Mid(cboCategoria.Text, 1, 2)
   Case "01" 'Depositos en Cajas
    vCuenta = fxTesParametro("05")
   Case "02" 'Depositos sin Identificar
    vCuenta = fxTesParametro("06")
   Case "03" 'Depositos Otros.."
    vCuenta = fxTesParametro("07")
End Select

txtCuentaDesc.Text = fxgCntCuentaDesc(vCuenta)
txtCuenta.Text = fxgCntCuentaFormato(True, vCuenta, 0)

End Sub

Private Sub cboFechas_Click()
If vPaso Then Exit Sub
vGridId.MaxRows = 0
End Sub

Private Sub cboFiltro_Click()
If vPaso Then Exit Sub
vGridId.MaxRows = 0
End Sub

Private Sub chkMarcas_Click()
Dim i As Long


For i = 1 To vGridId.MaxRows
   vGridId.Row = i
   vGridId.Col = 1
   vGridId.Value = chkMarcas.Value
Next i


End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim vProceso As Long

vModulo = 9
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.AppearanceStyle = fxGridStyle

vPaso = True

cboFechas.AddItem "Documento"
cboFechas.AddItem "Identificación"
cboFechas.AddItem "Registro"
cboFechas.Text = "Documento"


cboCategoria.Clear
cboCategoria.AddItem "01 - Depósitos de Cajas"
cboCategoria.AddItem "02 - Depósitos Sin Identificar"
cboCategoria.AddItem "03 - Depósitos Otros..."
cboCategoria.Text = "02 - Depósitos Sin Identificar"

cboFiltro.AddItem "TODOS"
cboFiltro.AddItem "Identificados - No Registrados"
cboFiltro.AddItem "Identificados - Registrados"
cboFiltro.AddItem "No Identificados - Registrados"
cboFiltro.AddItem "No Identificados - No Registrados"
cboFiltro.Text = "Identificados - No Registrados"

strSQL = "exec spTes_Cuenta_Bancaria_Acceso '" & glogon.Usuario & "','DP','SOL'"

Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

vPaso = False

dtpRegistroInicio.Value = fxFechaServidor
dtpRegistroCorte.Value = dtpRegistroInicio.Value

dtpIncoInicio.Value = dtpRegistroInicio.Value
dtpIncoCorte.Value = dtpIncoInicio.Value

tcMain.Item(0).Selected = True

Call cboBanco_Click
Call cboCategoria_Click
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbArchivoCarga()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset

Dim i As Integer, iCampos As Integer, vExiste As Integer
Dim vFecha As Date, vDocumento As String, vMonto As Currency, vDescripcion As String
Dim vCedula As String, vNombre As String, vInconsistencia As String

Dim curMonto As Currency, lCasos As Long

On Error GoTo vError
vGrid.MaxRows = 0

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboBanco.ListCount <= 0 Then
    MsgBox "No existe ninguna Institución, no se puede procesar el archivo...", vbCritical
    Exit Sub
End If

Me.MousePointer = vbHourglass

vGrid.MaxRows = 0

curMonto = 0
lCasos = 0 'Total

Set rsExcel = Excel_Load(txtArchivo.Text, "Import")

'Verifica Estructura del Archivo

iCampos = 0
For i = 0 To rsExcel.Fields.Count - 1
   Select Case UCase(rsExcel.Fields(i).Name)
      Case "DOCUMENTO", "FECHA", "MONTO", "DESCRIPCION"
        iCampos = iCampos + 1
      Case Else
      
   End Select
Next i

If iCampos < 4 Then
   Me.MousePointer = vbDefault
   MsgBox "1. No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "2. Los campos son Documento, Fecha, Monto y Descripcion", vbExclamation
 
   Exit Sub
End If


With vGrid



    Do While Not rsExcel.EOF
        vDocumento = Trim(rsExcel!Documento)
        vFecha = rsExcel!fecha
        vMonto = rsExcel!Monto
        vDescripcion = rsExcel!Descripcion
       
      If vDocumento <> "" Then
            strSQL = "select dbo.fxTes_DP_Cargado(" & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & vDocumento & "',''," & vMonto & ") as Existe"
            Call OpenRecordSet(rs, strSQL)
              vExiste = rs!Existe
              If vExiste > 0 Then vExiste = 1
              
              Select Case rs!Existe
                    Case 0 'Sin Inconsistencia
                      vInconsistencia = ""
                    Case 1 'Existe  / Identificado
                      vInconsistencia = "Existe  / Identificado"
                    Case 2 'Existe  / No Identificado
                      vInconsistencia = "Existe  / No Identificado"
                    Case 3 'Existe Registro pero a nombre de otra persona
                      vInconsistencia = "Existe Registro pero a nombre de otra persona"
                    Case 4 'Existe Registro con Monto Diferente
                      vInconsistencia = "Existe Registro con Monto Diferente"
              End Select
              
            rs.Close
      
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .Col = 1
            .Value = vbChecked
            
            .Col = 2
            .Value = vExiste
            
            .Col = 3
            .Text = vDocumento
            .Col = 4
            .Text = CStr(vMonto)
            .Col = 5
            .Text = vFecha
            .Col = 6
            .Text = vDescripcion
            .Col = 7
            .Text = vInconsistencia
            
            curMonto = curMonto + vMonto
            txtCasos.Text = txtCasos.Text + 1
       
       End If
       rsExcel.MoveNext
    Loop
    rsExcel.Close
    
End With
        
'Totales
txtMonto.Text = Format(curMonto, "Standard")
Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente", vbInformation


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia
End Sub

Private Sub sbProcesar()
Dim strSQL As String
Dim i As Long, vDescripcion As String, vCuenta As String, vInconsistencia As String
Dim vRequiereId As Integer, vDocumento As String, vMonto As Currency, vFecha As Date, vExiste As Integer
Dim vMensaje As Boolean, vCasos As Long

On Error GoTo vError


vCuenta = fxgCntCuentaFormato(False, txtCuenta.Text, 0)

If Not fxgCntCuentaValida(vCuenta) Then
   MsgBox "La cuenta especificada para registro no es válida...verifique!", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass

vMensaje = False
vCasos = 0

'Inicializa
strSQL = ""
With vGrid
    For i = 1 To .MaxRows

       .Row = i
       .Col = 1
       vRequiereId = .Value
       .Col = 2
       vExiste = .Value
       .Col = 3
       vDocumento = .Text
       .Col = 4
       vMonto = CCur(.Text)
       .Col = 5
       vFecha = Format(.Text, "yyyy/mm/dd")
       .Col = 6
       vDescripcion = .Text
       .Col = 7
       vInconsistencia = .Text
       
        If vExiste = 0 Then
            strSQL = strSQL & Space(10) & "insert TES_DEPOSITOS_TRAMITE(id_Banco,documento,nsolicitud,fecha,monto,descripcion,registro_fecha,registro_usuario " _
                   & ",id_requerida,identificado, cod_cuenta)" _
                   & " values(" & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & vDocumento & "',0,'" & Format(vFecha, "yyyy/mm/dd") _
                   & "'," & vMonto & ",'" & vDescripcion & "',dbo.MyGetdate(),'" & glogon.Usuario & "'," & vRequiereId & ",0,'" & vCuenta & "')"
            
            vCasos = vCasos + 1
            
        Else
            strSQL = strSQL & Space(10) & "insert TES_DEPOSITOS_TRAMITE_INCONSISTENCIAS(id_Banco,documento,fecha,monto,descripcion,registro_fecha,registro_usuario " _
                   & ",inconsistencia)" _
                   & " values(" & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & vDocumento & "','" & Format(vFecha, "yyyy/mm/dd") _
                   & "'," & vMonto & ",'" & vDescripcion & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & vInconsistencia & "')"
           vMensaje = True
        End If
       
       
       If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
       End If
    Next i
End With

If Len(strSQL) > 0 Then
     Call ConectionExecute(strSQL)
     strSQL = ""
End If

Me.MousePointer = vbDefault

If vCasos = 0 Then
    MsgBox "No se procesaron casos *--Revisados--* para el control de depósitos!", vbExclamation
Else
    MsgBox "Carga realizada Satisfactoriamente... Registros Procesados :" & vCasos, vbInformation
End If

If vMensaje Then
    MsgBox "Se presentaron inconsistencias en la carga..Revise en el TAB de consulta de inconsistencias!", vbExclamation
End If


Call sbLimpia

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

tcMain.Height = Me.Height - (tcMain.top + 650)
tcMain.Width = Me.Width - 500

vGrid.Height = tcMain.Height - (vGrid.top + fraAccion.Height + 550)
vGrid.Width = tcMain.Width - 350
fraAccion.top = vGrid.top + vGrid.Height + 100

vGridId.Height = tcMain.Height - (vGridId.top + 200)
vGridId.Width = tcMain.Width - 350

vGridInco.Height = tcMain.Height - (vGridInco.top + 200)
vGridInco.Width = tcMain.Width - 350

'btnAplicar.Top = fraAccion.Top + 50
'btnCancelar.Top = btnAplicar.Top

End Sub


Private Sub tlbInconsistencias_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
     Call sbInconsistenciaBuscar

End Select

End Sub



Private Sub sbRegistroBuscar()
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Long

On Error GoTo vError

If cboBanco.ListCount = 0 Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

'cboFiltro.AddItem "Identificados - No Registrados"Tra.IDENTIFICADO = " & chkIdentificados.Value _
'cboFiltro.AddItem "Identificados - Registrados"
'cboFiltro.AddItem "No Identificados - Registrados"
'cboFiltro.AddItem "No Identificados - No Registrados"



strSQL = "select Tra.*, Bn.Descripcion as 'BancoDesc'" _
        & " From TES_DEPOSITOS_TRAMITE Tra inner join Tes_Bancos Bn on Tra.id_banco = Bn.id_Banco"

Select Case Mid(cboFechas.Text, 1, 1)
   Case "D"
        strSQL = strSQL & " Where Tra.Fecha between '"
   Case "I"
        strSQL = strSQL & " Where Tra.Identifica_Fecha between '"
   Case "R"
        strSQL = strSQL & " Where Tra.Tes_Aplicado_Fecha between '"
End Select
strSQL = strSQL & Format(dtpRegistroInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpRegistroCorte.Value, "yyyy/mm/dd") & " 23:59:59'"


If Len(Trim(txtNumDoc.Text)) > 0 Then
    strSQL = strSQL & " and Tra.Documento like '%" & txtNumDoc.Text & "%'"
End If

Select Case cboFiltro.Text
   Case "Identificados - No Registrados"
        strSQL = strSQL & " and Tra.Identificado = 1 and Tra.Tes_Aplicado = 0"
   Case "Identificados - Registrados"
        strSQL = strSQL & " and Tra.Identificado = 1 and Tra.Tes_Aplicado = 1"
   Case "No Identificados - Registrados"
        strSQL = strSQL & " and Tra.Identificado = 0 and Tra.Tes_Aplicado = 1"
   Case "No Identificados - No Registrados"
        strSQL = strSQL & " and Tra.Identificado = 0 and Tra.Tes_Aplicado = 0"
End Select

strSQL = strSQL & " and Tra.Id_Banco = " & mBanco

Call OpenRecordSet(rs, strSQL)

vGridId.MaxRows = 0


  Do While Not rs.EOF
    vGridId.MaxRows = vGridId.MaxRows + 1
    vGridId.Row = vGridId.MaxRows
         
    vGridId.Col = 1

    For i = 2 To vGridId.MaxCols
      vGridId.Col = i
      Select Case i
         Case 2 'Tramite Id
            vGridId.Text = CStr(rs!DP_TRAMITE_ID)
         Case 3 'Num Documento
            vGridId.Text = rs!Documento
         Case 4 'Fecha del Documento
            vGridId.Text = Format(rs!fecha, "dd/mm/yyyy")
         Case 5 'Monto
            vGridId.Text = Format(rs!Monto, "Standard")
         Case 6 'Descripcion
            vGridId.Text = rs!Descripcion
         Case 7 'Registro Fecha
            vGridId.Text = rs!REGISTRO_FECHA & ""
         Case 8 'Registro Usuario
            vGridId.Text = rs!REGISTRO_USUARIO & ""
         Case 9 'Cliente Id
            vGridId.Text = rs!CLIENTE_ID & ""
         Case 10 'Cliente Nombre
            vGridId.Text = rs!Cliente_Nombre & ""
         Case 11 'Identifica Fecha
            vGridId.Text = rs!IDENTIFICA_FECHA & ""
         Case 12 'Identifica Usuario
            vGridId.Text = rs!IDENTIFICA_USUARIO & ""
         Case 13 'Tes. Registro Fecha
            vGridId.Text = rs!TES_APLICADO_FECHA & ""
         Case 14 'Tes. Registro Usuario
            vGridId.Text = rs!TES_APLICADO_USUARIO & ""
         Case 15 'Tes. Solicitud
            vGridId.Text = CStr(rs!NSolicitud & "")
      
      End Select
    Next i
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbInconsistenciaBuscar()
Dim strSQL As String

On Error GoTo vError

If cboBanco.ListCount = 0 Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "select Tra.Documento, Tra.Monto, Tra.Fecha, Tra.Descripcion, Tra.Inconsistencia, Tra.Registro_Fecha, Tra.Registro_Usuario, Bn.Descripcion as 'Banco' " _
        & " From TES_DEPOSITOS_TRAMITE_INCONSISTENCIAS Tra inner join Tes_Bancos Bn on Tra.id_banco = Bn.id_Banco" _
        & " Where Tra.Fecha between '" & Format(dtpIncoInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpIncoCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
        & " and Tra.Id_Banco = " & mBanco
Call sbCargaGrid(vGridInco, vGridInco.MaxCols, strSQL, True)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbRegistroAplicar()
Dim strSQL As String, i As Long
Dim vRemesa As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

vRemesa = fxTesParametro("08")
vRemesa = vRemesa + 1
strSQL = "update tes_parametros set valor = '" & vRemesa & "' where cod_parametro = '08'"
Call ConectionExecute(strSQL)

With vGridId

strSQL = ""
For i = 1 To .MaxRows
  .Row = i
  .Col = 1
  If .Value = vbChecked Then
    .Col = 3
    strSQL = strSQL & Space(10) & "exec spTES_Deposito_Lote_Registra " & mBanco & ",'" & .Text & "','" & glogon.Usuario & "'," & vRemesa
  End If

  If Len(strSQL) > 25000 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
  End If
  
Next i
End With

  If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
  End If

Me.MousePointer = vbDefault

MsgBox "Casos registrados en Banking satisfactoriamente!", vbInformation

Call sbRegistroBuscar


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Call sbRegistroBuscar
  
End Sub


Private Sub sbRegistroActualizar()
Dim strSQL As String, i As Long

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "exec spTES_Deposito_Lote_Actualiza"
Call ConectionExecute(strSQL)
  
Me.MousePointer = vbDefault

MsgBox "Casos Revisados y Actualizados en Tesorería satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub




Private Sub sbArchivoBusca()


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

    
    txtArchivo.Text = .FileName
End With

End Sub



Private Function fxExisteRegistro(pDocumento As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

'strSQL = "select cod_contrato from fnd_contratos where cedula = '" & vCedula & "' And cod_operadora = " & mOperadora & "" _
'         & " and cod_plan = '" & mPlan & "' and estado ='A'"
'Call OpenRecordSet(rs, strSQL)
'If rs.EOF Then
'    fxExisteRegistro = False
'Else
'    fxExisteRegistro = True
'    mContrato = rs!cod_contrato
'End If
'rs.Close

End Function


Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   If gCuenta <> "" Then
       txtCuenta.Text = gCuenta
       txtCuenta.SetFocus
   End If
End If
End Sub

Private Sub txtCuenta_LostFocus()
Dim vCuenta As String

vCuenta = fxgCntCuentaFormato(False, txtCuenta.Text, 0)
txtCuentaDesc.Text = fxgCntCuentaDesc(vCuenta)
txtCuenta.Text = fxgCntCuentaFormato(True, vCuenta, 0)

End Sub

