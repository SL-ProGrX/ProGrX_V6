VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCxPProveedores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Proveedores"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   7.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9510
   Begin XtremeSuiteControls.TabControl ssTab 
      Height          =   5295
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   9615
      _Version        =   1572864
      _ExtentX        =   16960
      _ExtentY        =   9340
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
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   6
      Item(0).Caption =   "General"
      Item(0).ControlCount=   25
      Item(0).Control(0)=   "txtCodAlter"
      Item(0).Control(1)=   "cbo"
      Item(0).Control(2)=   "txtCedJur"
      Item(0).Control(3)=   "txtTelefono1"
      Item(0).Control(4)=   "txtTelefonoExt"
      Item(0).Control(5)=   "txtFax"
      Item(0).Control(6)=   "txtFaxExt"
      Item(0).Control(7)=   "cboClasificacion"
      Item(0).Control(8)=   "txtApartadoPostal"
      Item(0).Control(9)=   "txtObservacion"
      Item(0).Control(10)=   "txtDireccion"
      Item(0).Control(11)=   "Label14(1)"
      Item(0).Control(12)=   "Label18(3)"
      Item(0).Control(13)=   "Label4"
      Item(0).Control(14)=   "Label6"
      Item(0).Control(15)=   "Label7(0)"
      Item(0).Control(16)=   "Label14(0)"
      Item(0).Control(17)=   "Label15"
      Item(0).Control(18)=   "Label16"
      Item(0).Control(19)=   "Label7(1)"
      Item(0).Control(20)=   "Label8"
      Item(0).Control(21)=   "txtEmail2"
      Item(0).Control(22)=   "Label2(0)"
      Item(0).Control(23)=   "cboEstado"
      Item(0).Control(24)=   "txtEmail"
      Item(1).Caption =   "Adicional"
      Item(1).ControlCount=   24
      Item(1).Control(0)=   "txtCuentaContable"
      Item(1).Control(1)=   "txtSaldo"
      Item(1).Control(2)=   "txtUltCompra"
      Item(1).Control(3)=   "txtMontoCredito"
      Item(1).Control(4)=   "txtDiasCredito"
      Item(1).Control(5)=   "txtDescuento"
      Item(1).Control(6)=   "txtAtencionPagos"
      Item(1).Control(7)=   "txtAtencionCompra"
      Item(1).Control(8)=   "txtNitCodigo"
      Item(1).Control(9)=   "txtNitNombre"
      Item(1).Control(10)=   "Label12(1)"
      Item(1).Control(11)=   "Label17"
      Item(1).Control(12)=   "Label12(0)"
      Item(1).Control(13)=   "Label11"
      Item(1).Control(14)=   "Label10"
      Item(1).Control(15)=   "Label9"
      Item(1).Control(16)=   "Label18(1)"
      Item(1).Control(17)=   "Label18(0)"
      Item(1).Control(18)=   "Label18(2)"
      Item(1).Control(19)=   "btnCuentas"
      Item(1).Control(20)=   "lswCuentas"
      Item(1).Control(21)=   "cboBancos"
      Item(1).Control(22)=   "Label3"
      Item(1).Control(23)=   "txtCuentaContableDesc"
      Item(2).Caption =   "Autorizaciones"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "vGrid"
      Item(2).Control(1)=   "Label1(3)"
      Item(3).Caption =   "Fusiones"
      Item(3).ControlCount=   6
      Item(3).Control(0)=   "lblProvFusion"
      Item(3).Control(1)=   "Label1(2)"
      Item(3).Control(2)=   "Label1(1)"
      Item(3).Control(3)=   "lblFusion"
      Item(3).Control(4)=   "Label1(4)"
      Item(3).Control(5)=   "lsw"
      Item(4).Caption =   "Suspender"
      Item(4).ControlCount=   10
      Item(4).Control(0)=   "btnSuspender(0)"
      Item(4).Control(1)=   "btnSuspender(1)"
      Item(4).Control(2)=   "lswS"
      Item(4).Control(3)=   "scSuspensiones"
      Item(4).Control(4)=   "txtSNotas"
      Item(4).Control(5)=   "cboSMotivo"
      Item(4).Control(6)=   "Label5(0)"
      Item(4).Control(7)=   "Label5(1)"
      Item(4).Control(8)=   "chkSVence"
      Item(4).Control(9)=   "dtpSVence"
      Item(5).Caption =   "Auto-Gestión"
      Item(5).ControlCount=   3
      Item(5).Control(0)=   "chkWebPortal"
      Item(5).Control(1)=   "chkWebFerias"
      Item(5).Control(2)=   "tcAutoGestion"
      Begin XtremeSuiteControls.ListView lswCuentas 
         Height          =   1692
         Left            =   -69760
         TabIndex        =   30
         Top             =   3480
         Visible         =   0   'False
         Width           =   8652
         _Version        =   1572864
         _ExtentX        =   15261
         _ExtentY        =   2984
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswS 
         Height          =   2535
         Left            =   -69880
         TabIndex        =   66
         Top             =   2760
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   4471
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3372
         Left            =   -69640
         TabIndex        =   62
         Top             =   1800
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1572864
         _ExtentX        =   15049
         _ExtentY        =   5948
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcAutoGestion 
         Height          =   4095
         Left            =   -70000
         TabIndex        =   76
         Top             =   1200
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1572864
         _ExtentX        =   16960
         _ExtentY        =   7223
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
         Item(0).Caption =   "Usuarios"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "lswUsuarios"
         Item(0).Control(1)=   "btnWebPortalUser"
         Item(1).Caption =   "Eventos"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswEventos"
         Begin XtremeSuiteControls.ListView lswEventos 
            Height          =   3735
            Left            =   -70000
            TabIndex        =   78
            Top             =   360
            Visible         =   0   'False
            Width           =   9495
            _Version        =   1572864
            _ExtentX        =   16748
            _ExtentY        =   6588
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswUsuarios 
            Height          =   3735
            Left            =   0
            TabIndex        =   77
            Top             =   360
            Width           =   9495
            _Version        =   1572864
            _ExtentX        =   16748
            _ExtentY        =   6588
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
         Begin XtremeSuiteControls.PushButton btnWebPortalUser 
            Height          =   375
            Left            =   8280
            TabIndex        =   79
            Top             =   0
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Usuarios"
            BackColor       =   -2147483633
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmProveedores.frx":000C
         End
      End
      Begin XtremeSuiteControls.CheckBox chkSVence 
         Height          =   255
         Left            =   -63760
         TabIndex        =   72
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Vence"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnSuspender 
         Height          =   375
         Index           =   0
         Left            =   -64240
         TabIndex        =   64
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Suspender"
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
      End
      Begin XtremeSuiteControls.PushButton btnCuentas 
         Height          =   372
         Left            =   -62800
         TabIndex        =   31
         Tag             =   "1"
         Top             =   3080
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cuentas Bancarias"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4092
         Left            =   -68800
         TabIndex        =   23
         Top             =   1020
         Visible         =   0   'False
         Width           =   7572
         _Version        =   524288
         _ExtentX        =   13356
         _ExtentY        =   7218
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
         MaxCols         =   485
         ScrollBars      =   2
         SpreadDesigner  =   "frmProveedores.frx":072C
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   1680
         TabIndex        =   33
         Top             =   600
         Width           =   1932
         _Version        =   1572864
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
      Begin XtremeSuiteControls.ComboBox cboClasificacion 
         Height          =   312
         Left            =   1680
         TabIndex        =   34
         Top             =   1200
         Width           =   3852
         _Version        =   1572864
         _ExtentX        =   6800
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   6960
         TabIndex        =   35
         Top             =   1176
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboBancos 
         Height          =   312
         Left            =   -67720
         TabIndex        =   36
         Top             =   3120
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8493
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
      Begin XtremeSuiteControls.FlatEdit txtObservacion 
         Height          =   1035
         Left            =   1680
         TabIndex        =   38
         Top             =   4200
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
         _ExtentY        =   1826
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
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   330
         Left            =   1680
         TabIndex        =   39
         Top             =   3240
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
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
      Begin XtremeSuiteControls.FlatEdit txtEmail2 
         Height          =   330
         Left            =   1680
         TabIndex        =   40
         Top             =   3720
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
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
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   915
         Left            =   1680
         TabIndex        =   41
         Top             =   2160
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12933
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
      Begin XtremeSuiteControls.FlatEdit txtTelefono1 
         Height          =   330
         Left            =   1680
         TabIndex        =   42
         Top             =   1680
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
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
      Begin XtremeSuiteControls.FlatEdit txtTelefonoExt 
         Height          =   330
         Left            =   2880
         TabIndex        =   43
         Top             =   1680
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
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
      Begin XtremeSuiteControls.FlatEdit txtFax 
         Height          =   330
         Left            =   3840
         TabIndex        =   44
         Top             =   1680
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
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
      Begin XtremeSuiteControls.FlatEdit txtFaxExt 
         Height          =   330
         Left            =   5040
         TabIndex        =   45
         Top             =   1680
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
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
      Begin XtremeSuiteControls.FlatEdit txtApartadoPostal 
         Height          =   330
         Left            =   6960
         TabIndex        =   46
         Top             =   1680
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.FlatEdit txtCedJur 
         Height          =   330
         Left            =   3600
         TabIndex        =   47
         Top             =   600
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
      Begin XtremeSuiteControls.FlatEdit txtCodAlter 
         Height          =   330
         Left            =   6960
         TabIndex        =   48
         Top             =   600
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
      Begin XtremeSuiteControls.FlatEdit txtAtencionCompra 
         Height          =   330
         Left            =   -67720
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1572864
         _ExtentX        =   11663
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
      Begin XtremeSuiteControls.FlatEdit txtAtencionPagos 
         Height          =   330
         Left            =   -67720
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   6612
         _Version        =   1572864
         _ExtentX        =   11663
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
      Begin XtremeSuiteControls.FlatEdit txtNitNombre 
         Height          =   330
         Left            =   -65800
         TabIndex        =   54
         Top             =   1320
         Visible         =   0   'False
         Width           =   4692
         _Version        =   1572864
         _ExtentX        =   8276
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
      Begin XtremeSuiteControls.FlatEdit txtNitCodigo 
         Height          =   330
         Left            =   -67720
         TabIndex        =   53
         Top             =   1320
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCuentaContable 
         Height          =   330
         Left            =   -67720
         TabIndex        =   55
         Top             =   1800
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCuentaContableDesc 
         Height          =   330
         Left            =   -65800
         TabIndex        =   56
         Top             =   1800
         Visible         =   0   'False
         Width           =   4692
         _Version        =   1572864
         _ExtentX        =   8276
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUltCompra 
         Height          =   330
         Left            =   -63640
         TabIndex        =   57
         Top             =   2280
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   330
         Left            =   -63640
         TabIndex        =   58
         Top             =   2640
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
      Begin XtremeSuiteControls.FlatEdit txtMontoCredito 
         Height          =   330
         Left            =   -67720
         TabIndex        =   59
         Top             =   2640
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtDiasCredito 
         Height          =   330
         Left            =   -67720
         TabIndex        =   60
         Top             =   2280
         Visible         =   0   'False
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtDescuento 
         Height          =   330
         Left            =   -66280
         TabIndex        =   61
         Top             =   2280
         Visible         =   0   'False
         Width           =   612
         _Version        =   1572864
         _ExtentX        =   1080
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.PushButton btnSuspender 
         Height          =   375
         Index           =   1
         Left            =   -62560
         TabIndex        =   65
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Activar"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtSNotas 
         Height          =   915
         Left            =   -68200
         TabIndex        =   68
         Top             =   960
         Visible         =   0   'False
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12933
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
      Begin XtremeSuiteControls.ComboBox cboSMotivo 
         Height          =   330
         Left            =   -68200
         TabIndex        =   69
         Top             =   600
         Visible         =   0   'False
         Width           =   3735
         _Version        =   1572864
         _ExtentX        =   6588
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
      Begin XtremeSuiteControls.DateTimePicker dtpSVence 
         Height          =   330
         Left            =   -62320
         TabIndex        =   73
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.CheckBox chkWebPortal 
         Height          =   495
         Left            =   -69760
         TabIndex        =   74
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Portal Empresarial de Proveedores y Oferentes"
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
      End
      Begin XtremeSuiteControls.CheckBox chkWebFerias 
         Height          =   495
         Left            =   -65320
         TabIndex        =   75
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Web para Ventas en Ferias y Eventos"
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
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Index           =   1
         Left            =   -69640
         TabIndex        =   71
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detalle"
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
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Index           =   0
         Left            =   -69640
         TabIndex        =   70
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Motivo"
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
      End
      Begin XtremeShortcutBar.ShortcutCaption scSuspensiones 
         Height          =   375
         Left            =   -69880
         TabIndex        =   67
         Top             =   2400
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Historial"
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
      Begin VB.Label Label3 
         Caption         =   "Cuenta/Desembolsos"
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
         Left            =   -69640
         TabIndex        =   37
         Top             =   3120
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label2 
         Caption         =   "Email (2)"
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
         Left            =   240
         TabIndex        =   32
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Proveedor Fusionado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   4
         Left            =   -67600
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label lblFusion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   -65440
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   4332
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nuevo Proveedor"
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
         Left            =   -69640
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   8532
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Proveedores Fusionados"
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
         Left            =   -69640
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   8532
      End
      Begin VB.Label lblProvFusion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "..."
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
         Left            =   -69640
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   8532
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Listas de Personas autorizadas para Pago de Facturas a Terceros: "
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
         Height          =   312
         Index           =   3
         Left            =   -69760
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   6012
      End
      Begin VB.Label Label18 
         Caption         =   "NIT (Cod/Nom)"
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
         TabIndex        =   22
         ToolTipText     =   "Nombre de Información Tributaria"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "Atención en Compras"
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
         Left            =   -69640
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.Label Label18 
         Caption         =   "Atención en Pagos"
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
         Left            =   -69640
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label9 
         Caption         =   "% Desc."
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
         Left            =   -67120
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label10 
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
         Left            =   -64960
         TabIndex        =   18
         Top             =   2640
         Visible         =   0   'False
         Width           =   492
      End
      Begin VB.Label Label11 
         Caption         =   "Días de Crédito"
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
         Left            =   -69640
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Ult. Compra"
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
         Left            =   -64960
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label17 
         Caption         =   "Monto de Crédito"
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
         Left            =   -69640
         TabIndex        =   15
         Top             =   2640
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label12 
         Caption         =   "Cuenta x Pagar"
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
         Left            =   -69640
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Observación"
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
         TabIndex        =   13
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Telefono"
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
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Dirección"
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
         TabIndex        =   11
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Email (1)"
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
         TabIndex        =   10
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Apto. Postal"
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
         Left            =   5760
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Clasificación"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         Left            =   5760
         TabIndex        =   7
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label Label4 
         Caption         =   "Fax"
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
         Left            =   3480
         TabIndex        =   6
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "ID Alterno"
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
         Left            =   5760
         TabIndex        =   4
         Top             =   600
         Width           =   972
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8880
      TabIndex        =   1
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin ComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6540
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   6244
            MinWidth        =   6244
            Text            =   "Saldo Divisa Extranjera:"
            TextSave        =   "Saldo Divisa Extranjera:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   6244
            MinWidth        =   6244
            Text            =   "Saldo Divisa Local: "
            TextSave        =   "Saldo Divisa Local: "
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Text            =   "Divisa:"
            TextSave        =   "Divisa:"
            Object.Tag             =   ""
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
      Height          =   330
      Left            =   1320
      TabIndex        =   49
      Top             =   480
      Width           =   1092
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   582
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
      Height          =   330
      Left            =   2400
      TabIndex        =   50
      Top             =   480
      Width           =   6372
      _Version        =   1572864
      _ExtentX        =   11239
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
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repBoleta"
                  Text            =   "Boleta "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repListadoGeneral"
                  Text            =   "Listado General"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
      Begin XtremeSuiteControls.PushButton btnAdjuntos 
         Height          =   330
         Left            =   9000
         TabIndex        =   80
         ToolTipText     =   "Adjuntar Documentos"
         Top             =   0
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   582
         _StockProps     =   79
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
         Picture         =   "frmProveedores.frx":0C48
      End
   End
   Begin XtremeSuiteControls.CheckBox chkFiltra 
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   81
      Top             =   840
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Auto Gestión"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.CheckBox chkFiltra 
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   83
      Top             =   840
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ferias"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.ComboBox cboFiltro 
      Height          =   315
      Left            =   2400
      TabIndex        =   84
      Top             =   840
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar:"
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
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   82
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
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
      Height          =   312
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1092
   End
End
Attribute VB_Name = "frmCxPProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean

Private Sub sbUsuarios_EnLinea()

If vCodigo = 0 Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCxP_Proveedores_Usuarios_List " & vCodigo
Call OpenRecordSet(rs, strSQL)

With lswUsuarios.ListItems
    .Clear
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!Usuario)
          itmX.SubItems(1) = rs!Nombre
          itmX.SubItems(2) = rs!Ferias_Desc
          itmX.SubItems(3) = rs!Portal_Desc
          itmX.SubItems(4) = rs!Activo_Desc
          itmX.SubItems(5) = rs!REGISTRO_FECHA & ""
          itmX.SubItems(6) = rs!REGISTRO_USUARIO & ""
      rs.MoveNext
    Loop
    rs.Close

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbEventos_EnLinea()

If vCodigo = 0 Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCxP_Proveedores_Eventos_List " & vCodigo
Call OpenRecordSet(rs, strSQL)


vPaso = True
With lswEventos.ListItems
    .Clear
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!cod_Evento)
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = rs!Inicio
          itmX.SubItems(3) = rs!Corte
          itmX.SubItems(4) = rs!REGISTRO_FECHA & ""
          itmX.SubItems(5) = rs!REGISTRO_USUARIO & ""
          
          itmX.Checked = rs!Asignado
          
      rs.MoveNext
    Loop
    rs.Close

End With

vPaso = False
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnAdjuntos_Click()
 gGA.Modulo = "CXP"
 gGA.Llave_01 = txtCodigo.Text
 gGA.Llave_02 = ""
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End Sub

Private Sub btnCuentas_Click()
If vCodigo = 0 Then
   MsgBox "Consulte un Proveedor Primero...", vbExclamation
   ssTab.Item(0).Selected = True
   Exit Sub
End If

GLOBALES.gTag = Trim(txtCedJur)
GLOBALES.gTag2 = "CxP"

frmCC_Cuentas_Bancarias.Show vbModal

Call sbCuentas_Load

End Sub

Private Sub btnSuspender_Click(Index As Integer)
Dim pActiva As Integer, pVence As String

txtSNotas.Text = fxSysCleanTxtInject(txtSNotas.Text)

If cboSMotivo.ListCount = 0 Then Exit Sub

If Len(txtSNotas.Text) <= 10 Then
    MsgBox "Indique una Nota válida!", vbExclamation
    Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass


If chkSVence.Value = xtpChecked Then
    pVence = "'" & Format(dtpSVence.Value, "yyyy-mm-dd") & "'"
Else
    pVence = "Null"
End If

If Index = 0 Then
  pActiva = 1
Else
  pActiva = 0
End If



If Index = 0 Then
  pActiva = 1
Else
  pActiva = 0
End If

strSQL = "exec spCxP_Suspension " & vCodigo & ", '" & cboSMotivo.ItemData(cboSMotivo.ListIndex) _
        & "', " & pActiva & ", '" & Mid(txtSNotas.Text, 1, 1000) & "', " & pVence & ", '" & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Registro realizado satisfactoriamente!", vbInformation

Call sbConsulta(vCodigo)

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnWebPortalUser_Click()
If vCodigo = 0 Then Exit Sub

GLOBALES.gTag = vCodigo
GLOBALES.gTag2 = txtNombre.Text

Call sbFormsCall("frmCxP_Proveedor_Usuarios", vbModal, , , False, Me)

Call sbUsuarios_EnLinea

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedJur.SetFocus
End Sub

Private Sub cboClasificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono1.SetFocus
End Sub

Private Sub chkSVence_Click()
If chkSVence.Value = xtpChecked Then
  dtpSVence.Enabled = True
Else
  dtpSVence.Enabled = False
End If
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_proveedor from cxp_proveedores"
           
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_proveedor > " & IIf(txtCodigo = "", 0, txtCodigo)
    Else
       strSQL = strSQL & " where cod_proveedor < " & IIf(txtCodigo = "", 0, txtCodigo)
    End If
    
    If chkFiltra(0).Value = xtpChecked Then
        strSQL = strSQL & " and WEB_AUTO_GESTION = 1"
    End If
    If chkFiltra(1).Value = xtpChecked Then
        strSQL = strSQL & " and WEB_FERIAS = 1"
    End If
    
    If Mid(cboFiltro.Text, 1, 1) <> "T" Then
        strSQL = strSQL & " and Estado = '" & Mid(cboFiltro.Text, 1, 1) & "'"
    End If
    
    
    If FlatScrollBar.Value = 1 Then
        strSQL = strSQL & " order by cod_proveedor asc"
    Else
        strSQL = strSQL & " order by cod_proveedor desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_Proveedor)
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

Private Sub sbCuentas_Load()

On Error GoTo vError

lswCuentas.ListItems.Clear
If vCodigo > 0 Then
    strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
           & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
           & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
           & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
           & " where C.Identificacion = '" & Trim(txtCedJur.Text) & "'"
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswCuentas.ListItems.Add(, , rs!CUENTA_INTERNA)
           itmX.SubItems(1) = Trim(rs!Banco)
           itmX.SubItems(2) = rs!TipoDesc
           itmX.SubItems(3) = rs!cod_Divisa
           itmX.SubItems(4) = IIf(rs!CUENTA_INTERBANCA = 1, "Sí", "No")
           itmX.SubItems(5) = rs!Destino & ""
           itmX.SubItems(6) = IIf(rs!ACTIVA = 1, "Activa", "Cerrada")
           itmX.SubItems(7) = rs!REGISTRO_FECHA & ""
           itmX.SubItems(8) = rs!REGISTRO_USUARIO & ""
     
       rs.MoveNext
    Loop
    rs.Close
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub






Private Sub lswEventos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spCxP_Proveedores_Eventos_Asigna " & vCodigo & ", " & Item.Text _
       & ", " & IIf(Item.Checked, 1, 0) & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
   Call Bitacora(rs!Movimiento, rs!Mensaje)
Else
    Item.Checked = IIf(Item.Checked, False, True)
    MsgBox "Este Evento no puede ser modificado porque se encuentra vencido!", vbExclamation
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub ssTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

Select Case Item.Index
   Case 1 'Cuentas
       Call sbCuentas_Load
       txtAtencionCompra.SetFocus
       
   Case 2  'Autorizaciones
        If vCodigo = 0 Then
           MsgBox "Consulte un Proveedor Primero...", vbExclamation
           ssTab.Item(0).Selected = True
        Else
           'Cargar Grid
          strSQL = "select cedula,nombre from cxp_autorizaciones where cod_proveedor = " & vCodigo _
                 & " order by cedula"
          Call sbCargaGrid(vGrid, 2, strSQL)
        End If
   Case 3 'Fusiones
        If vCodigo = 0 Then
           MsgBox "Consulte un Proveedor Primero...", vbExclamation
           ssTab.Item(0).Selected = True
        Else
          lsw.ListItems.Clear
                 strSQL = "select P.cod_proveedor,P.descripcion" _
                        & " from cxp_fusiones F inner join cxp_proveedores P on F.cod_proveedor = P.cod_proveedor" _
                        & " inner join cxp_proveedores X On F.cod_proveedor_fus = X.cod_proveedor" _
                        & " Where F.cod_proveedor_fus = " & vCodigo
                Call OpenRecordSet(rs, strSQL)
                If Not rs.EOF And Not rs.BOF Then
                  lblProvFusion.Caption = rs!cod_Proveedor & " - " & rs!Descripcion
                Else
                  lblProvFusion.Caption = ""
                End If
                rs.Close
                 
                strSQL = "select X.cod_proveedor,X.descripcion,X.fusion" _
                        & " from cxp_fusiones F inner join cxp_proveedores P on F.cod_proveedor = P.cod_proveedor" _
                        & " inner join cxp_proveedores X On F.cod_proveedor_fus = X.cod_proveedor" _
                        & " Where F.cod_proveedor = " & vCodigo
                Call OpenRecordSet(rs, strSQL, 0)
                Do While Not rs.EOF
                 Set itmX = lsw.ListItems.Add(, , rs!cod_Proveedor)
                     itmX.SubItems(1) = rs!Descripcion
                     itmX.SubItems(2) = rs!fusion & ""
                 rs.MoveNext
                Loop
                rs.Close
        End If 'vCodigo = 0
        
     Case 4 'Suspension
            
        txtSNotas.Text = ""
            
        dtpSVence.Value = fxFechaServidor
            
        strSQL = "select Rtrim(COD_SUSPENSION) as 'IdX', rtrim(descripcion) as 'ItmX'" _
                & " from CXP_SUSPENSION_TIPOS order by descripcion"
        Call sbCbo_Llena_New(cboSMotivo, strSQL, False, True)
    
        chkSVence.Value = xtpUnchecked
        Call chkSVence_Click
        
        btnSuspender(0).Visible = True
        btnSuspender(1).Visible = True
        
        If Mid(cboEstado.Text, 1, 1) = "S" Then
            btnSuspender(0).Visible = False
        Else
            btnSuspender(1).Visible = False
        End If


        If vCodigo = 0 Then
           MsgBox "Consulte un Proveedor Primero...", vbExclamation
           ssTab.Item(0).Selected = True
        Else
          
          
          lswS.ListItems.Clear

                 
                strSQL = "select *" _
                        & " from vCxP_Suspensiones" _
                        & " Where cod_proveedor = " & vCodigo
                Call OpenRecordSet(rs, strSQL, 0)
                Do While Not rs.EOF
                 Set itmX = lswS.ListItems.Add(, , rs!Suspension_Id)
                     itmX.SubItems(1) = rs!Suspension_Desc
                     itmX.SubItems(2) = rs!Notas
                     itmX.SubItems(3) = rs!Vencimiento & ""
                     itmX.SubItems(4) = rs!REGISTRO_FECHA & ""
                     itmX.SubItems(5) = rs!REGISTRO_USUARIO & ""
                     itmX.SubItems(6) = rs!REACTIVA_FECHA & ""
                     itmX.SubItems(7) = rs!REACTIVA_USUARIO & ""
                     itmX.SubItems(8) = rs!REACTIVA_NOTAS & ""
                 rs.MoveNext
                Loop
                rs.Close
        End If 'vCodigo = 0

    Case 5 'Auto Gestion
        tcAutoGestion.Item(0).Selected = True
        Call sbUsuarios_EnLinea
        
End Select


vError:
End Sub

Private Sub Form_Activate()
vModulo = 30
End Sub

Private Sub Form_Load()

On Error GoTo vError

vModulo = 30

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True

 vGrid.AppearanceStyle = fxGridStyle
 
lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Prov.Id", 1200
lsw.ColumnHeaders.Add , , "Nombre", 3500
lsw.ColumnHeaders.Add , , "Fecha", 2100, vbCenter
 
With lswCuentas.ColumnHeaders
    .Clear
    .Add 1, , "Cuenta", 2500
    .Add 2, , "Banco", 3500
    .Add 3, , "Tipo", 1100, vbCenter
    .Add 4, , "Divisa", 1100, vbCenter
    .Add 5, , "Interbanca", 1100, vbCenter
    .Add 6, , "Destino", 1100, vbCenter
    .Add 7, , "Activa", 1100, vbCenter
    .Add 8, , "Fecha", 2500
    .Add 9, , "Usuario", 2500
End With

With lswS.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Motivo", 2500
    .Add , , "Notas", 3000
    .Add , , "Vence?", 2100
    .Add , , "Reg.Fecha", 2100
    .Add , , "Reg.Usuario", 2100
    .Add , , "Act.Fecha", 2100
    .Add , , "Act.Usuario", 2100
    .Add , , "Act.Notas", 3000
End With

With lswUsuarios.ColumnHeaders
    .Add , , "Usuario", 2100
    .Add , , "Nombre", 4100
    .Add , , "Ferias ?", 1100, vbCenter
    .Add , , "Portal ?", 1100, vbCenter
    .Add , , "Activo ?", 1100, vbCenter
    .Add , , "R.Fecha", 2100, vbCenter
    .Add , , "R.Usuario ?", 2100, vbCenter
End With


With lswEventos.ColumnHeaders
    .Add , , "Evento Id", 1400
    .Add , , "Descripción", 3500
    .Add , , "Inicio", 1500, vbCenter
    .Add , , "Finaliza", 1500, vbCenter
    .Add , , "R.Fecha", 2100, vbCenter
    .Add , , "R.Usuario ?", 2100, vbCenter
End With


cbo.Clear
cbo.AddItem "Persona Física"
cbo.AddItem "Entidad Juridica"

cboFiltro.Clear
cboFiltro.AddItem "Activos"
cboFiltro.AddItem "Inactivos"
cboFiltro.AddItem "Suspendidos"
cboFiltro.AddItem "Todos"
cboFiltro.Text = "Todos"



strSQL = "exec spCxP_Bancos_Autorizados"
Call sbCbo_Llena_New(cboBancos, strSQL, False, True)

strSQL = "select Rtrim(Cod_clasificacion) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from cxp_prov_clas order by descripcion"
Call sbCbo_Llena_New(cboClasificacion, strSQL, False, True)

 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()

ssTab.Item(0).Selected = True

vCodigo = 0
txtCodigo.Text = ""

StatusBarX.Panels(1).Text = "Saldo Divisa Foránea:"
StatusBarX.Panels(2).Text = "Saldo Divisa Local:"
StatusBarX.Panels(3).Text = "Divisa:"

txtCuentaContable.Tag = GLOBALES.gEnlace

cbo.Text = "Entidad Juridica"

cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "InActivo"
cboEstado.Text = "Activo"

txtNombre.Text = ""
txtObservacion.Text = ""

lblFusion.Caption = ""

txtCedJur.Text = ""

txtDireccion.Text = ""
txtApartadoPostal.Text = ""
txtEmail.Text = ""
txtEmail2.Text = ""

txtTelefono1.Text = ""
txtTelefonoExt.Text = ""
txtFax.Text = ""
txtFaxExt.Text = ""

txtAtencionCompra.Text = ""
txtAtencionPagos.Text = ""

txtCuentaContable.Text = ""
txtCuentaContableDesc.Text = ""

txtNitCodigo.Text = ""
txtNitNombre.Text = ""

txtDescuento.Text = "0"
txtDiasCredito.Text = "0"
txtMontoCredito.Text = "0"
txtSaldo.Text = "0"
txtUltCompra.Text = ""


chkWebFerias.Value = xtpUnchecked
chkWebPortal.Value = xtpUnchecked

txtCodigo.Enabled = True

End Sub


Private Sub tcAutoGestion_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
    Case 0
        Call sbUsuarios_EnLinea
    Case 1
        Call sbEventos_EnLinea
End Select

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtNombre.SetFocus
      txtCodigo.Enabled = False
      
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
         gBusquedas.Columna = "descripcion"
         gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtNombre.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select P.*,rtrim(C.descripcion) as 'TipoProv',isnull(Cta.Descripcion,'') as 'CuentaConta'" _
       & ", dbo.fxSys_Cuenta_Bancos_Desc(P.cod_Banco) as 'Banco_Desc'" _
       & " from cxp_proveedores P inner join cxp_prov_clas C on P.cod_clasificacion = C.cod_clasificacion" _
       & " left join CntX_Cuentas Cta on P.cod_Cuenta = Cta.cod_Cuenta and Cta.cod_contabilidad = " & GLOBALES.gEnlace _
       & " where P.cod_proveedor = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  
  ssTab.Item(0).Selected = True
  
  vEdita = True
  vCodigo = rs!cod_Proveedor
  txtCodigo.Text = CStr(rs!cod_Proveedor)
  
  txtCodAlter = rs!Cod_Alter & ""
  
    txtNombre.Text = rs!Descripcion & ""
    txtObservacion.Text = rs!Observacion & ""
    cboClasificacion.Text = Trim(rs!tipoprov)
    
    txtCedJur.Text = rs!CEDJUR & ""
    
    lblFusion.Caption = IIf(IsNull(rs!fusion), "", "Fusión : " & rs!fusion)
    
    Call sbCboAsignaDato(cboBancos, rs!Banco_Desc, True, rs!cod_banco)
    
    cboEstado.Clear
    cboEstado.AddItem "Activo"
    cboEstado.AddItem "InActivo"
    
    Select Case rs!Estado
        Case "A"
            cboEstado.Text = "Activo"
        Case "I"
            cboEstado.Text = "InActivo"
        Case "S"
            cboEstado.Clear
            cboEstado.AddItem "Suspendido"
            cboEstado.Text = "Suspendido"
    End Select
    
    
    Select Case rs!Tipo
      Case "P", "F"
          cbo.Text = "Persona Física"
      Case "E", "J"
          cbo.Text = "Entidad Juridica"
    End Select
    
    txtDireccion.Text = rs!direccion & ""
    txtApartadoPostal.Text = rs!aptopostal & ""
    txtEmail.Text = rs!Email & ""
    txtEmail2.Text = rs!Email_02 & ""
    
    txtTelefono1.Text = rs!telefono & ""
    txtTelefonoExt.Text = rs!telefono_ext & ""
    txtFax.Text = rs!fax & ""
    txtFaxExt.Text = rs!fax_ext & ""
    
    txtAtencionCompra.Text = rs!contacto_compras & ""
    txtAtencionPagos.Text = rs!contacto_ventas & ""
    
    txtCuentaContable.Text = fxgCntCuentaFormato(True, rs!cod_cuenta & "")
    txtCuentaContableDesc.Text = rs!CuentaConta & ""
    
    txtNitCodigo.Text = rs!Nit_Codigo & ""
    txtNitNombre.Text = rs!nit_nombre & ""
    
    txtDescuento.Text = Format(rs!descuento_porc, "Standard")
    txtDiasCredito.Text = CStr(rs!credito_plazo)
    txtMontoCredito.Text = Format(rs!credito_monto, "Standard")
    
    txtSaldo.Text = Format(rs!Saldo, "Standard")
    txtUltCompra.Text = IIf(IsNull(rs!ultima_compra), "", Format(rs!ultima_compra, "dd/mm/yyyy"))
    
    
    StatusBarX.Panels(1).Text = "Saldo Divisa Foránea: " & Format(rs!Saldo_Divisa_Real, "Standard")
    StatusBarX.Panels(2).Text = "Saldo Divisa Local: " & Format(rs!Saldo, "Standard")
    StatusBarX.Panels(3).Text = "Divisa: " & rs!cod_Divisa
    StatusBarX.Panels(3).Tag = Trim(rs!cod_Divisa)
   
    chkWebFerias.Value = rs!WEB_FERIAS
    chkWebPortal.Value = rs!WEB_AUTO_GESTION
   
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
Dim vCuenta As String, vDivisa As String
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Verifica que exista ningun otro proveedor con la misma cedula juridica
strSQL = "select isnull(count(*),0) as Existe from cxp_proveedores" _
       & " where cod_proveedor not in(" & vCodigo & ") and cedJur = '" _
       & Trim(txtCedJur) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   vMensaje = vMensaje & vbCrLf & " - Existe ya un Proveedor registrado con la misma Cédula Jurídica ..."
End If
rs.Close


'Si Existe Enlace con ContaExpress / Realizar esta verificacion
If txtCuentaContable.Tag = "" Then
  If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, txtCuentaContable, 0)) Then
     vMensaje = vMensaje & vbCrLf & " - No se especificó una cuenta contable válida..."
  End If
End If


txtEmail.Text = Trim(txtEmail.Text)
txtEmail2.Text = Trim(txtEmail2.Text)

If Not fxEmail_Valida(txtEmail.Text) Then
    vMensaje = vMensaje & " - El Email principal no es válido!" & vbCrLf
End If

If Len(Trim(txtEmail2.Text)) > 0 Then
    If Not fxEmail_Valida(txtEmail2.Text) Then
        vMensaje = vMensaje & " - El Email secundario no es válido!" & vbCrLf
    End If
End If


If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Proveedor no es válido ..."

If Len(lblFusion.Caption) > 0 Then vMensaje = vMensaje & vbCrLf & " - No se puede Modificar Proveedores con Fusionados ..."


'Validar que el cambio de cuenta contable la divisa no cambien a menos de que el saldo sea 0
If IsNumeric(txtSaldo.Text) Then
    If vEdita And CCur(txtSaldo.Text) > 0 And Len(vMensaje) = 0 Then
        vCuenta = fxgCntCuentaFormato(False, txtCuentaContable)
    
        strSQL = "select cod_divisa from Cntx_Cuentas where cod_contabilidad = " & GLOBALES.gEnlace _
               & " and cod_cuenta = '" & vCuenta & "'"
        Call OpenRecordSet(rs, strSQL)
            vDivisa = Trim(rs!cod_Divisa)
        rs.Close
        If vDivisa <> StatusBarX.Panels(3).Tag Then
            vMensaje = vMensaje & vbCrLf & " - El proveedor tiene Saldo por Pagar y se intenta cambiar la divisa del mismo vía cambio de cuenta contable..."
        End If
    
    End If
End If 'Isnumeric


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim vDivisa As String, vCuenta As String

On Error GoTo vError

vCuenta = fxgCntCuentaFormato(False, txtCuentaContable)

strSQL = "select cod_divisa from Cntx_Cuentas where cod_contabilidad = " & GLOBALES.gEnlace _
       & " and cod_cuenta = '" & vCuenta & "'"
Call OpenRecordSet(rs, strSQL)
    vDivisa = Trim(rs!cod_Divisa)
rs.Close

If vEdita Then
  strSQL = "update cxp_proveedores set descripcion = '" & UCase(Trim(txtNombre)) & "',cod_alter = '" & txtCodAlter _
         & "',cedJur = '" & txtCedJur & "',tipo = '" & Mid(cbo.Text, 1, 1) _
         & "',observacion = '" & txtObservacion & "',estado = '" & Mid(cboEstado.Text, 1, 1) _
         & "',direccion = '" & txtDireccion & "',aptopostal = '" & txtApartadoPostal _
         & "',email = '" & txtEmail & "', telefono = '" & txtTelefono1 _
         & "',email_02 = '" & txtEmail2.Text & "', telefono_ext = '" & txtTelefonoExt & "',fax = '" & txtFax _
         & "',fax_ext = '" & txtFaxExt & "',contacto_compras = '" & txtAtencionCompra _
         & "',contacto_ventas = '" & txtAtencionPagos & "',cod_cuenta = '" & vCuenta _
         & "',descuento_porc = " & txtDescuento _
         & ",credito_plazo = " & txtDiasCredito & ",credito_monto = " & CCur(txtMontoCredito) _
         & ",cod_clasificacion = '" & cboClasificacion.ItemData(cboClasificacion.ListIndex) _
         & "',nit_Codigo = '" & txtNitCodigo & "',nit_nombre = '" & txtNitNombre _
         & "',cod_divisa = '" & vDivisa & "', cod_Banco = " & cboBancos.ItemData(cboBancos.ListIndex) _
         & ", WEB_AUTO_GESTION = " & chkWebPortal.Value & ", WEB_FERIAS = " & chkWebFerias.Value _
         & ", MODIFICA_FECHA = dbo.MyGetdate(), MODIFICA_USUARIO = '" & glogon.Usuario & "'" _
         & "  where cod_proveedor = " & vCodigo
  Call ConectionExecute(strSQL)
  
  
  
  Call Bitacora("Modifica", "Proveedor Cod: " & vCodigo)

Else
   strSQL = "select isnull(max(cod_proveedor),0) as ultimo from cxp_proveedores"
   Call OpenRecordSet(rs, strSQL)
     txtCodigo = rs!ultimo + 1
     vCodigo = txtCodigo
   rs.Close
   
   strSQL = "insert into cxp_proveedores(cod_proveedor,tipo,cod_clasificacion,descripcion,cod_alter, observacion," _
          & "estado,contacto_ventas,contacto_compras,telefono,telefono_ext,fax,fax_ext,email,email_02,aptopostal," _
          & "direccion,credito_plazo,credito_monto,descuento_porc,saldo,cod_cuenta," _
          & "cedJur,Nit_Codigo,Nit_Nombre,cod_divisa,saldo_divisa_real,cod_banco, WEB_AUTO_GESTION, WEB_FERIAS, REGISTRO_FECHA, REGISTRO_USUARIO) values(" _
          & vCodigo & ",'" & Mid(cbo.Text, 1, 1) & "','" & cboClasificacion.ItemData(cboClasificacion.ListIndex) _
          & "','" & txtNombre & "','" & txtCodAlter & "','" & txtObservacion & "','" & Mid(cboEstado.Text, 1, 1) & "','" _
          & txtAtencionPagos & "','" & txtAtencionCompra & "','" & txtTelefono1 & "','" & txtTelefonoExt _
          & "','" & txtFax & "','" & txtFaxExt & "','" & txtEmail & "','" & txtEmail2.Text & "','" & txtApartadoPostal _
          & "','" & txtDireccion & "'," & txtDiasCredito & "," & CCur(txtMontoCredito) & "," _
          & txtDescuento & ",0,'" & vCuenta & "','" & txtCedJur & "','" & txtNitCodigo _
          & "','" & txtNitNombre & "','" & vDivisa & "',0," & cboBancos.ItemData(cboBancos.ListIndex) _
          & ", " & chkWebPortal.Value & ", " & chkWebFerias.Value & ", dbo.MyGetdate(), '" & glogon.Usuario & "'" & ")"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Proveedor Cod: " & vCodigo)
    
   txtCodigo.Enabled = True
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete cxp_proveedores where cod_proveedor = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Proveedor Cod: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtApartadoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub txtAtencionCompra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAtencionPagos.SetFocus
End Sub

Private Sub txtAtencionPagos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNitCodigo.SetFocus
End Sub

Private Sub txtCedJur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodAlter.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Nombre"
  gBusquedas.Col3Name = "Id. Real"
  
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedjur"
  gBusquedas.Orden = "cedjur"
  gBusquedas.Consulta = "select cod_proveedor, Descripcion, CedJur from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If


End Sub


Private Sub txtCedJur_LostFocus()
If Not vEdita And txtNitCodigo.Text = "" Then txtNitCodigo.Text = txtCedJur.Text
End Sub

Private Sub txtCodAlter_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboClasificacion.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Id Proveedor"
  gBusquedas.Col2Name = "Id Alterno"
  gBusquedas.Col3Name = "Nombre"
  
  gBusquedas.Columna = "cod_alter"
  gBusquedas.Orden = "cod_alter"
  gBusquedas.Consulta = "select cod_proveedor,cod_Alter,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

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
  
    If chkFiltra(0).Value = xtpChecked Then
        gBusquedas.Filtro = gBusquedas.Filtro & " and WEB_AUTO_GESTION = 1"
    End If
    If chkFiltra(1).Value = xtpChecked Then
        gBusquedas.Filtro = gBusquedas.Filtro & " and WEB_FERIAS = 1"
    End If
    If Mid(cboFiltro.Text, 1, 1) <> "T" Then
        gBusquedas.Filtro = gBusquedas.Filtro & " and Estado = '" & Mid(cboFiltro.Text, 1, 1) & "'"
    End If
  
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub


Private Sub txtCuentaContable_GotFocus()
On Error GoTo vError
If txtCuentaContable.Tag = "S" Then
    txtCuentaContable = fxgCntCuentaFormato(False, txtCuentaContable)
End If
vError:
End Sub

Private Sub txtCuentaContable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDiasCredito.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaContable = gCuenta
End If
End Sub

Private Sub txtCuentaContable_LostFocus()
On Error GoTo vError
txtCuentaContable = fxgCntCuentaFormato(True, txtCuentaContable)
txtCuentaContableDesc.Text = fxSIFCCodigos("D", fxgCntCuentaFormato(False, fxgCntCuentaFormato(False, txtCuentaContable)), "cuentas")
vError:
End Sub

Private Sub txtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMontoCredito.SetFocus
End Sub

Private Sub txtDiasCredito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescuento.SetFocus
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub


Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail2.SetFocus
End Sub

Private Sub txtEmail2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservacion.SetFocus
End Sub


Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFaxExt.SetFocus
End Sub


Private Sub txtFaxExt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApartadoPostal.SetFocus
End Sub

Private Sub txtMontoCredito_GotFocus()
On Error GoTo vError
 txtMontoCredito = CCur(txtMontoCredito)
vError:
End Sub

Private Sub txtMontoCredito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtMontoCredito_LostFocus()
On Error GoTo vError
 txtMontoCredito = Format(CCur(txtMontoCredito), "Standard")
vError:
End Sub

Private Sub txtNitCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNitNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "NIT_CODIGO"
  gBusquedas.Orden = "NIT_CODIGO"
  gBusquedas.Consulta = "select cod_proveedor,descripcion,NIT_CODIGO,NIT_NOMBRE from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If
End Sub

Private Sub txtNitNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaContable.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "NIT_NOMBRE"
  gBusquedas.Orden = "NIT_NOMBRE"
  gBusquedas.Consulta = "select cod_proveedor,descripcion,NIT_CODIGO,NIT_NOMBRE from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  
    If chkFiltra(0).Value = xtpChecked Then
        gBusquedas.Filtro = gBusquedas.Filtro & " and WEB_AUTO_GESTION = 1"
    End If
    If chkFiltra(1).Value = xtpChecked Then
        gBusquedas.Filtro = gBusquedas.Filtro & " and WEB_FERIAS = 1"
    End If
    If Mid(cboFiltro.Text, 1, 1) <> "T" Then
        gBusquedas.Filtro = gBusquedas.Filtro & " and Estado = '" & Mid(cboFiltro.Text, 1, 1) & "'"
    End If
    
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtNombre_LostFocus()
If Not vEdita And txtNitNombre.Text = "" Then txtNitNombre.Text = txtNombre.Text
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    ssTab.Item(1).Selected = True
    txtAtencionCompra.SetFocus
End If
End Sub

Private Sub txtTelefono1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefonoExt.SetFocus
End Sub

Private Sub txtTelefonoExt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFax.SetFocus
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  strSQL = "select isnull(count(0),0) as Existe from cxp_autorizaciones" _
         & " where cod_proveedor = " & vCodigo _
         & " and cedula = '" & vGrid.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then
     strSQL = "insert cxp_autorizaciones(cod_proveedor,cedula,nombre) values(" _
            & vCodigo & ",'" & vGrid.Text & "','"
     vGrid.Col = 2
     strSQL = strSQL & vGrid.Text & "')"
  Else
     vGrid.Col = 2
     strSQL = "update cxp_autorizaciones set nombre = '" & vGrid.Text _
            & "' where cod_proveedor = " & vCodigo & " and cedula = '"
     vGrid.Col = 1
     strSQL = strSQL & vGrid.Text & "'"
  End If
  Call ConectionExecute(strSQL)
  rs.Close
  
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


'Inserta Linea
If KeyCode = vbKeyDelete Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   strSQL = "delete cxp_autorizaciones where cod_proveedor = " & vCodigo _
          & " and cedula = '" & vGrid.Text & "'"
   Call ConectionExecute(strSQL)
   
'   Call ssTab_SelectedChanged
End If


End Sub
