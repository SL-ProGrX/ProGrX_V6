VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_APA_Acreedores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Acreedores"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9405
   Icon            =   "frmCR_APA_Acreedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Acreedores"
      TabPicture(0)   =   "frmCR_APA_Acreedores.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vGridAcreedores"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Datos"
      TabPicture(1)   =   "frmCR_APA_Acreedores.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2(7)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2(8)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2(10)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label2(12)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label2(13)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label2(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label2(14)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(0)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(2)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(3)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1(4)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1(5)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label1(6)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtCod_CCargos"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtCod_CGastos"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtCod_CTransitoria"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtCuentaDesc"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtCuenta"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtTelefono2"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtTelefono1"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtWebsite"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtCod_Acreedor"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtDireccion"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cboEstado"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtDescripcion"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtCod_CTransitoriaDesc"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtCod_CGastosDesc"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "txtCod_CCargosDesc"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtBanco_DC"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txtBanco_CK"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txtCod_CComision"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txtCod_CComisionDesc"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtBanco_CKDesc"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "txtBanco_DCDesc"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).ControlCount=   36
      TabCaption(2)   =   "Contactos"
      TabPicture(2)   =   "frmCR_APA_Acreedores.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Line3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtAcreedorContactos"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "vGridContactos"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Autorizados"
      TabPicture(3)   =   "frmCR_APA_Acreedores.frx":68A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Line7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "txtAcreedorAutorizado"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "vGridAutorizados"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin XtremeSuiteControls.FlatEdit txtBanco_DCDesc 
         Height          =   330
         Left            =   -71280
         TabIndex        =   40
         Top             =   6120
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtBanco_CKDesc 
         Height          =   330
         Left            =   -71280
         TabIndex        =   38
         Top             =   5760
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCod_CComisionDesc 
         Height          =   330
         Left            =   -71280
         TabIndex        =   36
         Top             =   5040
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCod_CComision 
         Height          =   330
         Left            =   -73200
         TabIndex        =   35
         Top             =   5040
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtBanco_CK 
         Height          =   330
         Left            =   -73200
         TabIndex        =   37
         Top             =   5760
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtBanco_DC 
         Height          =   330
         Left            =   -73200
         TabIndex        =   39
         Top             =   6120
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCod_CCargosDesc 
         Height          =   330
         Left            =   -71280
         TabIndex        =   34
         Top             =   4680
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCod_CGastosDesc 
         Height          =   330
         Left            =   -71280
         TabIndex        =   32
         Top             =   4320
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCod_CTransitoriaDesc 
         Height          =   330
         Left            =   -71280
         TabIndex        =   30
         Top             =   3960
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   330
         Left            =   -73200
         TabIndex        =   15
         Top             =   1080
         Width           =   6975
         _Version        =   1441793
         _ExtentX        =   12303
         _ExtentY        =   582
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
      Begin FPSpreadADO.fpSpread vGridAcreedores 
         Height          =   5772
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   8412
         _Version        =   524288
         _ExtentX        =   14838
         _ExtentY        =   10181
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
         MaxCols         =   491
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_APA_Acreedores.frx":68C2
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridContactos 
         Height          =   5412
         Left            =   -74760
         TabIndex        =   3
         Top             =   1080
         Width           =   8532
         _Version        =   524288
         _ExtentX        =   15050
         _ExtentY        =   9546
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
         MaxCols         =   490
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_APA_Acreedores.frx":6E94
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridAutorizados 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   11
         Top             =   1080
         Width           =   8775
         _Version        =   524288
         _ExtentX        =   15478
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
         MaxCols         =   486
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_APA_Acreedores.frx":74FD
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   -73200
         TabIndex        =   13
         Top             =   1440
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   735
         Left            =   -73200
         TabIndex        =   14
         Top             =   2160
         Width           =   6975
         _Version        =   1441793
         _ExtentX        =   12303
         _ExtentY        =   1296
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
      Begin XtremeSuiteControls.FlatEdit txtCod_Acreedor 
         Height          =   495
         Left            =   -73200
         TabIndex        =   16
         Top             =   480
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
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
      Begin XtremeSuiteControls.FlatEdit txtWebsite 
         Height          =   330
         Left            =   -73200
         TabIndex        =   19
         Top             =   3000
         Width           =   6975
         _Version        =   1441793
         _ExtentX        =   12303
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtTelefono1 
         Height          =   330
         Left            =   -73200
         TabIndex        =   20
         Top             =   1800
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono2 
         Height          =   330
         Left            =   -68160
         TabIndex        =   21
         Top             =   1800
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuenta 
         Height          =   330
         Left            =   -73200
         TabIndex        =   27
         Top             =   3600
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
         Height          =   330
         Left            =   -71280
         TabIndex        =   28
         Top             =   3600
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCod_CTransitoria 
         Height          =   330
         Left            =   -73200
         TabIndex        =   29
         Top             =   3960
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCod_CGastos 
         Height          =   330
         Left            =   -73200
         TabIndex        =   31
         Top             =   4320
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCod_CCargos 
         Height          =   330
         Left            =   -73200
         TabIndex        =   33
         Top             =   4680
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtAcreedorContactos 
         Height          =   330
         Left            =   -74760
         TabIndex        =   41
         Top             =   480
         Width           =   8535
         _Version        =   1441793
         _ExtentX        =   15055
         _ExtentY        =   582
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
         BackColor       =   16777152
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAcreedorAutorizado 
         Height          =   330
         Left            =   -74880
         TabIndex        =   42
         Top             =   600
         Width           =   8775
         _Version        =   1441793
         _ExtentX        =   15478
         _ExtentY        =   582
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
         BackColor       =   16777152
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   6
         Left            =   -74640
         TabIndex        =   26
         Top             =   3000
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Sitio Web"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   -74640
         TabIndex        =   25
         Top             =   2160
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Dirección"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   -69480
         TabIndex        =   24
         Top             =   1800
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Teléfono No.2"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   -74640
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Teléfono No.1"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   22
         Top             =   1440
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   18
         Top             =   1080
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Descripción"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   495
         Index           =   0
         Left            =   -74640
         TabIndex        =   17
         Top             =   480
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Código"
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
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000004&
         X1              =   -74760
         X2              =   -68160
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Banco DC"
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
         Left            =   -74640
         TabIndex        =   10
         ToolTipText     =   "Banco para Débitos de Cuenta"
         Top             =   6180
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Banco CK"
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
         Left            =   -74640
         TabIndex        =   9
         ToolTipText     =   "Banco para Cheques"
         Top             =   5820
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Cta Comisión"
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
         Left            =   -74640
         TabIndex        =   8
         Top             =   5100
         Width           =   1572
      End
      Begin VB.Label Label2 
         Caption         =   "Cta Acreedor"
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
         Left            =   -74640
         TabIndex        =   7
         Top             =   3660
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Cta Cargos"
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
         Left            =   -74640
         TabIndex        =   6
         Top             =   4740
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Cta Gastos"
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
         Left            =   -74640
         TabIndex        =   5
         Top             =   4380
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Cta Transitoria"
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
         Left            =   -74640
         TabIndex        =   4
         Top             =   4020
         Width           =   1692
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000004&
         X1              =   -74760
         X2              =   -68160
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   -74640
         X2              =   -68040
         Y1              =   5640
         Y2              =   5640
      End
   End
   Begin MSComctlLib.ImageList imgMenu 
      Left            =   7440
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Acreedores.frx":7A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Acreedores.frx":E295
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Acreedores.frx":14AF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Acreedores.frx":1B359
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Acreedores.frx":21BBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_Acreedores.frx":36D2D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   556
      ButtonWidth     =   1826
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Agregar Acreedor"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Acreedor"
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Editar"
            Key             =   "Editar"
            Object.ToolTipText     =   "Editar Acreedor"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar"
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar los cambios"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprime Boleta del Traslado"
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.Label lblTitulo 
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   240
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Registro de Acreedores"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line6 
      X1              =   3240
      X2              =   4440
      Y1              =   3360
      Y2              =   3840
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmCR_APA_Acreedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim vEdita As Boolean, vCodigo As Long, vCedulaPrincipal As String
'' Variable para controlar si se realizó algún cambio para guardar
Dim vCambios As Boolean
Dim mAcreedor As String

Private Sub sbCargarListaAcreedores()
'' Procedimiento para cargar la lista de acreedores
    Dim strSQL As String
    
    'Consulta la lista de acreedores
    strSQL = "select COD_ACREEDOR, DESCRIPCION, case ESTADO when 'A' then 'Activo' " & _
        " when 'I' then 'Inactivo'" & _
        " when 'B' then 'Bloqueado' else ESTADO end " & _
        " from CRD_APA_ACREEDORES "
        
    Call sbCargaGridCheckIni(vGridAcreedores, 3, strSQL)
    vGridAcreedores.MaxRows = vGridAcreedores.MaxRows - 1
End Sub

Private Sub cboEstado_Click()
    vCambios = True
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtTelefono1.SetFocus
    End If
End Sub



Private Sub Form_Activate()

    vModulo = 14
    
End Sub



Private Sub Form_Load()


On Error GoTo vError

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

    If GLOBALES.gEnlace = 0 Then
        Call sbgCntParametros
    End If
    
    '' Carga nombre de la terminal
    If Len(glogon.Maquina) = 0 Then
        Call sbMaquina
    End If

    vEdita = True
    vCambios = False
    ssTab.Tab = 0
    Call sbCargaComboEstado
    Call sbCargarListaAcreedores
    Call sbLimpiaPantalla
    Call ssTab_Click(0)
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbCargaComboEstado()
'' Carga el combo de los estados de los acreedores
    cboEstado.Clear
    cboEstado.AddItem "Activo"
    cboEstado.AddItem "Inactivo"
    cboEstado.AddItem "Bloqueado"
    cboEstado.Text = "Activo"
End Sub

Private Sub sbLimpiaPantalla()
'' Procedimiento para limpiar los campos del tab de datos del acreedor
    txtCod_Acreedor = Empty
    txtDescripcion = Empty
    txtTelefono1 = Empty
    txtTelefono2 = Empty
    txtDireccion = Empty
    txtWebsite = Empty
    
    txtCuenta = Empty
    txtCod_CTransitoria = Empty
    txtCod_CGastos = Empty
    txtCod_CCargos = Empty
    txtCod_CComision = Empty
    txtBanco_CK = Empty
    txtBanco_DC = Empty
    
    txtCuentaDesc = Empty
    txtCod_CTransitoriaDesc = Empty
    txtCod_CGastosDesc = Empty
    txtCod_CCargosDesc = Empty
    txtCod_CComisionDesc = Empty
    txtBanco_CKDesc = Empty
    txtBanco_DCDesc = Empty
End Sub



Private Sub ssTab_Click(PreviousTab As Integer)
Dim mAcreedor As String
' Eventos al cambiar de tab
    Select Case ssTab.Tab
    Case 0 'Acredores
        'Activa Tabs y Botones
        ssTab.TabEnabled(0) = True
        ssTab.TabEnabled(1) = False
        ssTab.TabEnabled(2) = True
        ssTab.TabEnabled(3) = True
        
        tlbPrincipal.Buttons(1).Enabled = True
        tlbPrincipal.Buttons(3).Enabled = True
        tlbPrincipal.Buttons(5).Enabled = False
        
    Case 1 'Datos del acreedor
    
        ssTab.TabEnabled(0) = True
        ssTab.TabEnabled(1) = True
        ssTab.TabEnabled(2) = False
        ssTab.TabEnabled(3) = False
        
        tlbPrincipal.Buttons(1).Enabled = False
        tlbPrincipal.Buttons(3).Enabled = False
        tlbPrincipal.Buttons(5).Enabled = True
        
        If vEdita = True Then
            txtCod_Acreedor.Locked = True
        Else
            txtCod_Acreedor.Locked = False
        End If
        
    Case 2 ' Contactos
    
        ssTab.TabEnabled(0) = True
        ssTab.TabEnabled(1) = False
        ssTab.TabEnabled(2) = True
        ssTab.TabEnabled(3) = True
        
        tlbPrincipal.Buttons(1).Enabled = False
        tlbPrincipal.Buttons(3).Enabled = False
        tlbPrincipal.Buttons(5).Enabled = False
        
        mAcreedor = fxGridValorMarcado(vGridAcreedores, 2)
        If mAcreedor <> Empty Then
            vEdita = False
            txtAcreedorContactos.Text = fxGridValorMarcado(vGridAcreedores, 3)
            txtAcreedorContactos.Tag = Trim(mAcreedor)
            ' Carga grid de contactos por acreedor
            Call sbCargarContactos
            vCambios = False
        Else
          MsgBox "Debe seleccionar el acreedor que desea ver los contactos"
        End If
        
        vGridContactos.SetFocus
    
    Case 3 ' Autorizados
    
        ssTab.TabEnabled(0) = True
        ssTab.TabEnabled(1) = False
        ssTab.TabEnabled(2) = True
        ssTab.TabEnabled(3) = True
        
        tlbPrincipal.Buttons(1).Enabled = False
        tlbPrincipal.Buttons(3).Enabled = False
        tlbPrincipal.Buttons(5).Enabled = False
        
        mAcreedor = fxGridValorMarcado(vGridAcreedores, 2)
        If mAcreedor <> Empty Then
            vEdita = False
            txtAcreedorAutorizado.Text = fxGridValorMarcado(vGridAcreedores, 3)
            txtAcreedorAutorizado.Tag = Trim(mAcreedor)
            ' Carga grid de autorizados por acreedor
            Call sbCargarAutorizados
            vCambios = False
        Else
          MsgBox "Debe seleccionar el acreedor que desea ver los autorizados"
        End If
        
        vGridContactos.SetFocus
        
    End Select
End Sub



Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
'Opciones del Menú


Select Case UCase(Button.Key)
    'Insertar Acreedor
    Case "INSERTAR", "NUEVO"
          ' Desactiva los Checks del grid de acreedores
          Call sbGridLimpiarMarcas(vGridAcreedores)
          vEdita = False 'Variable Indica que no es editar
          vCambios = False 'Controla si hay cambios que guardar
          ssTab.Tab = 1
          Call sbLimpiaPantalla
          txtCod_Acreedor.Locked = False
          txtCod_Acreedor.SetFocus
      
    ' Modificar Acreedor
    Case "MODIFICAR", "EDITAR"
        ' Cargar el código del acredor marcado
        mAcreedor = fxGridValorMarcado(vGridAcreedores, 2)
        If mAcreedor <> Empty Then
            vEdita = True
            'Carga los datos del acreedor seleccionado
            Call sbConsulta(Trim(mAcreedor))
            vCambios = False
            ssTab.Tab = 1
        Else
            MsgBox "Debe seleccionar el acreedor que desea modificar"
        End If
        
    ' Proceso para guardar nuevos y editar acreedores
    Case "GUARDAR", "SALVAR"
        'Verifica guardar solo si hubo cambios
        If vCambios = True Then
            If fxValida Then
                Call sbGuardar
            Else
                Exit Sub
            End If
        End If
        Call sbGridLimpiarMarcas(vGridAcreedores)
        Call sbLimpiaPantalla
        ssTab.Tab = 0
        
End Select

End Sub


Private Sub sbConsulta(Acreedor As String)
' Carga los datos de un Acreedor
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer


On Error GoTo vError

    Me.MousePointer = vbHourglass

    strSQL = "select COD_ACREEDOR, DESCRIPCION, ESTADO, TELEFONO1, TELEFONO2, DIRECCION, WEBSITE, COD_CUENTA " _
              & ",COD_CUENTA_TRANSITORIA, COD_CUENTA_GASTOS, COD_CUENTA_CARGOS, COD_CUENTA_COMISION, BANCO_CK, " _
              & " BANCO_DC from crd_apa_acreedores " _
              & " where cod_acreedor = '" & Trim(Acreedor) & "'"
    Call OpenRecordSet(rs, strSQL)

    If Not rs.BOF And Not rs.EOF Then
    '  Call sbToolBar(tlbPrincipal, "activo")
  
        ssTab.Tab = 1
        vEdita = True
        txtCod_Acreedor.Text = IIf(IsNull(rs!Cod_Acreedor), Empty, rs!Cod_Acreedor)
        txtDescripcion.Text = IIf(IsNull(rs!Descripcion), Empty, rs!Descripcion)
        txtTelefono1.Text = IIf(IsNull(rs!TELEFONO1), Empty, rs!TELEFONO1)
        txtTelefono2.Text = IIf(IsNull(rs!TELEFONO2), Empty, rs!TELEFONO2)
        txtDireccion.Text = IIf(IsNull(rs!DIRECCION), Empty, rs!DIRECCION)
        txtWebsite.Text = IIf(IsNull(rs!WEBSITE), Empty, rs!WEBSITE)
        txtCuenta.Text = IIf(IsNull(rs!cod_cuenta), Empty, rs!cod_cuenta)
        txtCod_CTransitoria.Text = IIf(IsNull(rs!COD_CUENTA_TRANSITORIA), Empty, rs!COD_CUENTA_TRANSITORIA)
        txtCod_CGastos.Text = IIf(IsNull(rs!COD_CUENTA_GASTOS), Empty, rs!COD_CUENTA_GASTOS)
        txtCod_CCargos.Text = IIf(IsNull(rs!COD_CUENTA_CARGOS), Empty, rs!COD_CUENTA_CARGOS)
        txtCod_CComision.Text = IIf(IsNull(rs!COD_CUENTA_COMISION), Empty, rs!COD_CUENTA_COMISION)
        
        txtBanco_CK.Text = IIf(IsNull(rs!BANCO_CK), Empty, rs!BANCO_CK)
        txtBanco_DC.Text = IIf(IsNull(rs!BANCO_DC), Empty, rs!BANCO_DC)
        
        Select Case Trim(rs!Estado)
            Case "A"
                cboEstado.Text = "Activo"
            Case "B"
                cboEstado.Text = "Bloqueado"
            Case "I"
                cboEstado.Text = "Inactivo"
        End Select
        
        Call sbCuentas_Desc(txtCuenta, txtCuentaDesc)
        Call sbCuentas_Desc(txtCod_CTransitoria, txtCod_CTransitoriaDesc)
        Call sbCuentas_Desc(txtCod_CGastos, txtCod_CGastosDesc)
        Call sbCuentas_Desc(txtCod_CCargos, txtCod_CCargosDesc)
        Call sbCuentas_Desc(txtCod_CComision, txtCod_CComisionDesc)
        
        Call sbCargarNombreBancos
        
    
    Else 'Caso no encuentra acreedor con ese código
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

Private Sub sbCargarNombreBancos()
Dim rs As New ADODB.Recordset, strSQL As String
On Error GoTo vError

    If Trim(txtBanco_CK) <> Empty Then
        '' Consulta la descripcion del banco de cheques
        strSQL = "select isnull(DESCRIPCION,'') from BANCOS where ID_BANCO = " & txtBanco_CK
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            txtBanco_CKDesc.Text = rs.Fields(0)
        End If
        rs.Close
    End If
    
    If Trim(txtBanco_DC) <> Empty Then
        '' Consulta la descripcion del banco de debito de cuenta
        strSQL = "select isnull(DESCRIPCION,'') from BANCOS where ID_BANCO = " & txtBanco_DC
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            txtBanco_DCDesc.Text = rs.Fields(0)
        End If
        rs.Close
    End If
   
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
'Valida los datos antes de almacenarlos en la bd
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

    vMensaje = ""
    fxValida = True

    If txtCod_Acreedor.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar el código del acreedor"
    If txtDescripcion.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe completar la descripción del acreedor"
    If cboEstado.Text = Empty Then vMensaje = vMensaje & vbCrLf & " Debe de seleccionar el estado del acreedor"
    
    
    
    If Not fxgCntCuentaValida(txtCod_CCargos.Text) Then vMensaje = vMensaje & vbCrLf & " Cuenta Contable para Cargos no es válida"
    If Not fxgCntCuentaValida(txtCod_CComision.Text) Then vMensaje = vMensaje & vbCrLf & " Cuenta Contable para Comisión no es válida"
    If Not fxgCntCuentaValida(txtCod_CGastos.Text) Then vMensaje = vMensaje & vbCrLf & " Cuenta Contable para Gastos no es válida"
    If Not fxgCntCuentaValida(txtCod_CTransitoria.Text) Then vMensaje = vMensaje & vbCrLf & " Cuenta Contable Transitoria no es válida"
    If Not fxgCntCuentaValida(txtCuenta.Text) Then vMensaje = vMensaje & vbCrLf & " Cuenta Contable para Acreedor no es válida"
    
    If vEdita = False Then
    
        'Verifica que exista ningun acreedor con ese código
        strSQL = "select isnull(count(*),0) as Existe from CRD_APA_ACREEDORES" _
               & " where COD_ACREEDOR = '" & txtCod_Acreedor.Text & "' "
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe > 0 Then
           vMensaje = vMensaje & vbCrLf & " Ya Existe un Acreedor con ese código"
        End If
        rs.Close
        
    End If

    If Len(vMensaje) > 0 Then
      fxValida = False
      MsgBox vMensaje, vbCritical
    End If

End Function

Private Sub sbGuardar()
' Procedimiento para guardar nuevos acreedores y guardar cambios
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

    If vEdita Then
        strSQL = "update crd_apa_acreedores set descripcion = '" & Trim(txtDescripcion) _
               & "',estado = '" & Mid(cboEstado, 1, 1) _
               & "', telefono1 = '" & Trim(txtTelefono1) _
               & "',telefono2 = '" & Trim(txtTelefono2) _
               & "',direccion = '" & Trim(txtDireccion) _
               & "', website = '" & Trim(txtWebsite) _
               & "',cod_cuenta = '" & fxgCntCuentaFormato(False, txtCuenta.Text, 0) _
               & "',cod_cuenta_transitoria = '" & fxgCntCuentaFormato(False, txtCod_CTransitoria, 0) _
               & "',cod_cuenta_gastos = '" & fxgCntCuentaFormato(False, txtCod_CGastos, 0) _
               & "',cod_cuenta_cargos = '" & fxgCntCuentaFormato(False, txtCod_CCargos, 0) _
               & "',cod_cuenta_comision = '" & fxgCntCuentaFormato(False, txtCod_CComision, 0) _
               & "',banco_ck = '" & Trim(txtBanco_CK) _
               & "',banco_dc = '" & Trim(txtBanco_DC) _
               & "' where cod_acreedor = '" & Trim(txtCod_Acreedor) & "'"
        
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Modifica", "APA Acreedor: " & Trim(txtCod_Acreedor) & " - " & Trim(txtDescripcion))
        
    Else
    
        strSQL = "insert crd_apa_acreedores(COD_ACREEDOR, DESCRIPCION, ESTADO, TELEFONO1, TELEFONO2, DIRECCION, WEBSITE, COD_CUENTA " _
               & ",COD_CUENTA_TRANSITORIA, COD_CUENTA_GASTOS, COD_CUENTA_CARGOS, COD_CUENTA_COMISION, " _
               & "BANCO_CK,BANCO_DC) values('" & Trim(txtCod_Acreedor) & "','" & txtDescripcion _
               & "','" & Mid(cboEstado, 1, 1) & "','" & Trim(txtTelefono1.Text) & "','" & Trim(txtTelefono2.Text) & "','" & Trim(txtDireccion.Text) _
               & "','" & Trim(txtWebsite.Text) & "','" & fxgCntCuentaFormato(False, txtCuenta.Text, 0) & "','" _
               & fxgCntCuentaFormato(False, txtCod_CTransitoria.Text, 0) & "','" & fxgCntCuentaFormato(False, txtCod_CGastos.Text, 0) _
               & "','" & fxgCntCuentaFormato(False, txtCod_CCargos.Text, 0) & "','" & fxgCntCuentaFormato(False, txtCod_CComision.Text, 0) _
               & "','" & Trim(txtBanco_CK) & "','" & Trim(txtBanco_DC) & "')" _
               
        Call ConectionExecute(strSQL)
                     
        Call Bitacora("Registra", "APA Acreedor: " & Trim(txtCod_Acreedor) & " - " & Trim(txtDescripcion))
             
    End If

    MsgBox "Información guardada satisfactoriamente...", vbInformation
    Call sbCargarListaAcreedores

    'Call sbToolBar(tlbPrincipal, "activo")
    Call RefrescaTags(Me)

    Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Function fxGuardarContacto() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

    fxGuardarContacto = 0
    vGridContactos.Row = vGridContactos.ActiveRow
    vGridContactos.Col = 1
    
    strSQL = "select isnull(count(*),0) as Existe from crd_apa_contactos " _
           & " where cod_contacto = '" & vGridContactos.Text & "'" _
           & " and cod_acreedor = '" & txtAcreedorContactos.Tag & "'"
    
    Call OpenRecordSet(rs, strSQL)

    If rs!Existe = 0 Then 'Insertar
        
        If Trim(vGridContactos.Text) = "" Then
            MsgBox "Debe asignar un código para el contacto"
            Exit Function
        End If
        
        strSQL = "insert into crd_apa_contactos(cod_acreedor,cod_contacto,nombre,tel_cel," _
               & " tel_trabajo, tel_fax, email) values('" & txtAcreedorContactos.Tag & "','" _
               & UCase(Trim(vGridContactos.Text)) & "','"
        
        vGridContactos.Col = 2
        strSQL = strSQL & UCase(Trim(vGridContactos.Text)) & "','"
        vGridContactos.Col = 3
        strSQL = strSQL & UCase(Trim(vGridContactos.Text)) & "','"
        vGridContactos.Col = 4
        strSQL = strSQL & UCase(Trim(vGridContactos.Text)) & "','"
        vGridContactos.Col = 5
        strSQL = strSQL & UCase(Trim(vGridContactos.Text)) & "','"
        vGridContactos.Col = 6
        strSQL = strSQL & UCase(Trim(vGridContactos.Text)) & "')"
    
        Call ConectionExecute(strSQL)
        
        vGridContactos.Col = 1
        Call Bitacora("REGISTRA", "APA Contacto: " & Trim(vGridContactos.Text) & " Acreedor: " & Trim(txtAcreedorContactos.Tag) & " - " & Trim(txtAcreedorContactos))

    Else 'Actualizar

        vGridContactos.Col = 2
        strSQL = "update crd_apa_contactos set nombre = '" & Trim(vGridContactos.Text) & "',tel_cel = '"
        vGridContactos.Col = 3
        strSQL = strSQL & UCase(Trim(vGridContactos.Text)) & "' , tel_trabajo = '"
        vGridContactos.Col = 4
        strSQL = strSQL & UCase(Trim(vGridContactos.Text)) & "' , tel_fax = '"
        vGridContactos.Col = 5
        strSQL = strSQL & UCase(Trim(vGridContactos.Text)) & "' , email = '"
        vGridContactos.Col = 6
        strSQL = strSQL & UCase(Trim(vGridContactos.Text)) & "' where cod_contacto = '"
        vGridContactos.Col = 1
        strSQL = strSQL & Trim(vGridContactos.Text) & "' and cod_acreedor = '" & txtAcreedorContactos.Tag & "'"
     
        Call ConectionExecute(strSQL)
    
        vGridContactos.Col = 1
        Call Bitacora("MODIFICA", "APA Contacto: " & Trim(vGridContactos.Text) & " Acreedor: " & Trim(txtAcreedorContactos.Tag) & " - " & Trim(txtAcreedorContactos))

    End If
    rs.Close
    fxGuardarContacto = 1
    Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Function fxGuardarAutorizado() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

    fxGuardarAutorizado = 0
    vGridAutorizados.Row = vGridAutorizados.ActiveRow
    vGridAutorizados.Col = 1
    
    strSQL = "select isnull(count(*),0) as Existe from CRD_APA_AUTORIZADOSCK " _
           & " where CEDULA = '" & vGridAutorizados.Text & "'" _
           & " and COD_ACREEDOR = '" & txtAcreedorAutorizado.Tag & "'"
    
    Call OpenRecordSet(rs, strSQL)

    If rs!Existe = 0 Then 'Insertar
        
        If Trim(vGridAutorizados.Text) = "" Then
            MsgBox "Debe asignar un código para el contacto"
            Exit Function
        End If
        
        strSQL = "insert into CRD_APA_AUTORIZADOSCK(COD_ACREEDOR,CEDULA,NOMBRE)" _
               & " values('" & txtAcreedorAutorizado.Tag & "','"
        
        vGridAutorizados.Col = 1
        strSQL = strSQL & UCase(Trim(vGridAutorizados.Text)) & "','"
        vGridAutorizados.Col = 2
        strSQL = strSQL & UCase(Trim(vGridAutorizados.Text)) & "')"
    
        Call ConectionExecute(strSQL)
        
        vGridAutorizados.Col = 1
        Call Bitacora("REGISTRA", "APA Autorizado: " & Trim(vGridAutorizados.Text) & " Acreedor: " & Trim(txtAcreedorAutorizado.Tag) & " - " & Trim(txtAcreedorAutorizado))


    Else 'Actualizar

        vGridAutorizados.Col = 2
        strSQL = "update CRD_APA_AUTORIZADOSCK set NOMBRE = '" & UCase(Trim(vGridAutorizados.Text)) & "'"
        strSQL = strSQL & " where CEDULA = '"
        vGridAutorizados.Col = 1
        strSQL = strSQL & Trim(vGridAutorizados.Text) & "' and cod_acreedor = '" & txtAcreedorAutorizado.Tag & "'"
     
        Call ConectionExecute(strSQL)
        
        vGridAutorizados.Col = 1
        Call Bitacora("MODIFICA", "APA Autorizado: " & Trim(vGridAutorizados.Text) & " Acreedor: " & Trim(txtAcreedorAutorizado.Tag) & " - " & Trim(txtAcreedorAutorizado))

    End If
    rs.Close
    fxGuardarAutorizado = 1
    Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub sbCargarContactos()
' Llena la lista de contactos
Dim strSQL As String
On Error GoTo vError
    
    strSQL = "select cod_contacto,nombre,tel_cel,tel_trabajo,tel_fax,email from crd_apa_contactos" _
           & " where cod_acreedor = '" & txtAcreedorContactos.Tag & "' " _
           & " order by nombre"
    
    Call sbCargaGrid(vGridContactos, 6, strSQL)
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargarAutorizados()
' Llena la lista de Autorizados
Dim strSQL As String
On Error GoTo vError
    
    strSQL = "select CEDULA, NOMBRE from CRD_APA_AUTORIZADOSCK " _
           & " where COD_ACREEDOR = '" & txtAcreedorAutorizado.Tag & "' " _
           & " order by NOMBRE"
    
    Call sbCargaGrid(vGridAutorizados, 2, strSQL)
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub txtBanco_CK_Change()
    vCambios = True
End Sub

Private Sub txtBanco_CK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "DESCRIPCION"
        gBusquedas.Orden = "ID_BANCO"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select ID_BANCO, DESCRIPCION " _
                            & " from BANCOS "
        frmBusquedas.Show vbModal
        txtBanco_CK.Text = gBusquedas.Resultado
        txtBanco_CKDesc.Text = gBusquedas.Resultado2
    
    End If
    
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtBanco_DC.SetFocus
    End If
End Sub

Private Sub txtBanco_CK_LostFocus()
    Call sbCargarNombreBancos
End Sub



Private Sub txtBanco_DC_Change()
    vCambios = True
End Sub

Private Sub txtBanco_DC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        gBusquedas.Columna = "DESCRIPCION"
        gBusquedas.Orden = "ID_BANCO"
        gBusquedas.Filtro = ""
        gBusquedas.Consulta = "select ID_BANCO, DESCRIPCION " _
                            & " from BANCOS "
        frmBusquedas.Show vbModal
        txtBanco_DC.Text = gBusquedas.Resultado
        txtBanco_DCDesc.Text = gBusquedas.Resultado2
    
    End If
    
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtCod_Acreedor.SetFocus
    End If
End Sub

Private Sub txtBanco_DC_LostFocus()
    Call sbCargarNombreBancos
End Sub

Private Sub txtCod_Acreedor_Change()
    vCambios = True
End Sub

Private Sub txtCod_Acreedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtDescripcion.SetFocus
    End If
End Sub

Private Sub txtCod_CCargos_Change()
    vCambios = True
End Sub

Private Sub txtCod_CCargos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusqueda(4)
    
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtCod_CComision.SetFocus
    End If
    
End Sub

Private Sub txtCod_CCargos_LostFocus()
    Call sbCuentas_Desc(txtCod_CCargos, txtCod_CCargosDesc)
End Sub

Private Sub txtCod_CComision_Change()
    vCambios = True
End Sub

Private Sub txtCod_CComision_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusqueda(5)
    
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtBanco_CK.SetFocus
    End If
    
End Sub

Private Sub txtCod_CComision_LostFocus()
    Call sbCuentas_Desc(txtCod_CComision, txtCod_CComisionDesc)
End Sub

Private Sub txtCod_CGastos_Change()
    vCambios = True
End Sub

Private Sub txtCod_CGastos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusqueda(3)
    
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtCod_CCargos.SetFocus
    End If
End Sub

Private Sub txtCod_CGastos_LostFocus()
    Call sbCuentas_Desc(txtCod_CGastos, txtCod_CGastosDesc)
End Sub

Private Sub txtCod_CTransitoria_Change()
    vCambios = True
End Sub

Private Sub txtCod_CTransitoria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusqueda(2)
    
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtCod_CGastos.SetFocus
    End If
End Sub

Private Sub txtCod_CTransitoria_LostFocus()
    Call sbCuentas_Desc(txtCod_CTransitoria, txtCod_CTransitoriaDesc)
End Sub

Private Sub txtCuenta_Change()
    vCambios = True
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then Call sbBusqueda(1)
        
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtCod_CTransitoria.SetFocus
    End If
End Sub

Private Sub txtCuenta_LostFocus()
    Call sbCuentas_Desc(txtCuenta, txtCuentaDesc)
End Sub

Private Sub txtDescripcion_Change()
    vCambios = True
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        cboEstado.SetFocus
    End If
End Sub

Private Sub txtDireccion_Change()
    vCambios = True
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtWebsite.SetFocus
    End If
End Sub

Private Sub txtTelefono1_Change()
    vCambios = True
End Sub



Private Sub txtTelefono1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtTelefono2.SetFocus
    End If
End Sub

Private Sub txtTelefono2_Change()
    vCambios = True
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtDireccion.SetFocus
    End If
End Sub

Private Sub txtWebsite_Change()
    vCambios = True
End Sub

Private Sub txtWebsite_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab) Then
        txtCuenta.SetFocus
    End If
End Sub

Private Sub vGridAcreedores_Click(ByVal Col As Long, ByVal Row As Long)
    Call sbGridMarcarSoloUno(vGridAcreedores, Row)
End Sub

Private Sub sbBusqueda(Index As Integer)

On Error GoTo vError

Call sbgCntCuentaConsulta

If gBusquedas.Resultado <> "" Then
    Select Case Index
     Case 1
         txtCuenta.Text = gBusquedas.Resultado
         Call sbCuentas_Desc(txtCuenta, txtCuentaDesc)
         txtCuenta.SetFocus
     Case 2
         txtCod_CTransitoria.Text = gBusquedas.Resultado
         Call sbCuentas_Desc(txtCod_CTransitoria, txtCod_CTransitoriaDesc)
         txtCod_CTransitoria.SetFocus
     Case 3
         txtCod_CGastos.Text = gBusquedas.Resultado
         Call sbCuentas_Desc(txtCod_CGastos, txtCod_CGastosDesc)
         txtCod_CGastos.SetFocus
     Case 4
         txtCod_CCargos.Text = gBusquedas.Resultado
         Call sbCuentas_Desc(txtCod_CCargos, txtCod_CCargosDesc)
         txtCod_CCargos.SetFocus
    Case 5
         txtCod_CComision.Text = gBusquedas.Resultado
         Call sbCuentas_Desc(txtCod_CComision, txtCod_CComisionDesc)
         txtCod_CComision.SetFocus
    End Select
End If

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub




Private Sub vGridAutorizados_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

If vGridAutorizados.ActiveCol = vGridAutorizados.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
    i = fxGuardarAutorizado
    If i = 0 Then Exit Sub
        vGridAutorizados.Row = vGridAutorizados.ActiveRow
        If vGridAutorizados.MaxRows <= vGridAutorizados.ActiveRow Then
            vGridAutorizados.MaxRows = vGridAutorizados.MaxRows + 1
            vGridAutorizados.Row = vGridAutorizados.MaxRows
        End If
    End If

    'Inserta Linea
    If KeyCode = vbKeyInsert Then
        vGridAutorizados.MaxRows = vGridAutorizados.MaxRows + 1
        vGridAutorizados.InsertRows vGridAutorizados.ActiveRow, 1
        vGridAutorizados.Row = vGridAutorizados.ActiveRow
    End If

    If KeyCode = vbKeyDelete Then
        Dim Autorizado As String

        vGridAutorizados.Row = vGridAutorizados.ActiveRow
        vGridAutorizados.Col = 1
        If vGridAutorizados.Text <> "" Then
            Autorizado = vGridAutorizados.Text
            strSQL = "delete CRD_APA_AUTORIZADOSCK where  CEDULA = " & Autorizado _
                & " and COD_ACREEDOR = '" & txtAcreedorAutorizado.Tag & "'"
            Call ConectionExecute(strSQL)
            
            Call Bitacora("BORRA", "APA Autorizado: " & Trim(Autorizado) & " Acreedor: " & Trim(txtAcreedorAutorizado.Tag) & " - " & Trim(txtAcreedorAutorizado))

            Call sbCargarAutorizados
        End If
    End If



End Sub

Private Sub vGridContactos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

If vGridContactos.ActiveCol = vGridContactos.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
    i = fxGuardarContacto
    If i = 0 Then Exit Sub
        vGridContactos.Row = vGridContactos.ActiveRow
        If vGridContactos.MaxRows <= vGridContactos.ActiveRow Then
            vGridContactos.MaxRows = vGridContactos.MaxRows + 1
            vGridContactos.Row = vGridContactos.MaxRows
        End If
    End If

    'Inserta Linea
    If KeyCode = vbKeyInsert Then
        vGridContactos.MaxRows = vGridContactos.MaxRows + 1
        vGridContactos.InsertRows vGridContactos.ActiveRow, 1
        vGridContactos.Row = vGridContactos.ActiveRow
    End If

    If KeyCode = vbKeyDelete Then
        Dim contacto As Integer

        vGridContactos.Row = vGridContactos.ActiveRow
        vGridContactos.Col = 1
        If vGridContactos.Text <> "" Then
            contacto = vGridContactos.Text
            strSQL = "delete crd_apa_contactos where  cod_contacto = " & contacto _
                & " and cod_acreedor = '" & txtAcreedorContactos.Tag & "'"
            Call ConectionExecute(strSQL)
            
            Call Bitacora("BORRA", "APA Contacto: " & contacto & " Acreedor: " & Trim(txtAcreedorContactos.Tag) & " - " & Trim(txtAcreedorContactos))
        
            Call sbCargarContactos
        End If
    End If


End Sub

Private Sub sbCuentas_Desc(ObjId As Object, ObjDesc As Object)

ObjId.Text = fxgCntCuentaFormato(False, ObjId.Text, 0)
ObjDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, ObjId.Text, 0))
ObjId.Text = fxgCntCuentaFormato(True, ObjId.Text, 0)

End Sub
