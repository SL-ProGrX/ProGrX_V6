VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_PromotoresPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promotores / Ejecutivos de Cuenta"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   10215
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7425
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Usuario de Modificación"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha Modificacion"
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
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AFRA"
                  Text            =   "Resumen de Afiliaciones"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AFDA"
                  Text            =   "Detalle de Afiliaciones"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AFSEP1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AFLP"
                  Text            =   "Listado General"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   9480
      TabIndex        =   1
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   9975
      _Version        =   1441793
      _ExtentX        =   17595
      _ExtentY        =   11245
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
      Item(0).ControlCount=   31
      Item(0).Control(0)=   "Label3(0)"
      Item(0).Control(1)=   "Label3(1)"
      Item(0).Control(2)=   "Label3(2)"
      Item(0).Control(3)=   "Label3(3)"
      Item(0).Control(4)=   "Label3(4)"
      Item(0).Control(5)=   "Label3(5)"
      Item(0).Control(6)=   "Label3(6)"
      Item(0).Control(7)=   "Label3(7)"
      Item(0).Control(8)=   "Label3(8)"
      Item(0).Control(9)=   "Label3(9)"
      Item(0).Control(10)=   "Label3(10)"
      Item(0).Control(11)=   "chkComision"
      Item(0).Control(12)=   "txtCedJur"
      Item(0).Control(13)=   "txtPagarA"
      Item(0).Control(14)=   "txtUsuario"
      Item(0).Control(15)=   "txtTelefono1"
      Item(0).Control(16)=   "txtTelefonoExt"
      Item(0).Control(17)=   "txtFax"
      Item(0).Control(18)=   "txtFaxExt"
      Item(0).Control(19)=   "txtApartadoPostal"
      Item(0).Control(20)=   "txtEmail"
      Item(0).Control(21)=   "txtDireccion"
      Item(0).Control(22)=   "txtObservacion"
      Item(0).Control(23)=   "cboEstado"
      Item(0).Control(24)=   "cboTipo"
      Item(0).Control(25)=   "lswCuentas"
      Item(0).Control(26)=   "btnCuentas"
      Item(0).Control(27)=   "cboBancos"
      Item(0).Control(28)=   "Label3(11)"
      Item(0).Control(29)=   "Label3(13)"
      Item(0).Control(30)=   "cboTipoPago"
      Item(1).Caption =   "Listado"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "cboListadoTipo"
      Item(1).Control(1)=   "Opt(0)"
      Item(1).Control(2)=   "Opt(1)"
      Item(1).Control(3)=   "lsw"
      Item(1).Control(4)=   "txtFiltro"
      Item(1).Control(5)=   "btnExport"
      Item(1).Control(6)=   "scTitulo(0)"
      Begin XtremeSuiteControls.ListView lswCuentas 
         Height          =   1572
         Left            =   1320
         TabIndex        =   29
         Top             =   4596
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   2773
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
         View            =   3
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4935
         Left            =   -70000
         TabIndex        =   40
         Top             =   1320
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1441793
         _ExtentX        =   17595
         _ExtentY        =   8705
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton Opt 
         Height          =   372
         Index           =   0
         Left            =   -69760
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Activos ?"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtCedJur 
         Height          =   312
         Left            =   1320
         TabIndex        =   16
         Top             =   480
         Width           =   1692
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkComision 
         Height          =   252
         Left            =   1320
         TabIndex        =   7
         Top             =   3840
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica comisión?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtPagarA 
         Height          =   312
         Left            =   4200
         TabIndex        =   17
         Top             =   480
         Width           =   5652
         _Version        =   1441793
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   1320
         TabIndex        =   18
         Top             =   840
         Width           =   1692
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono1 
         Height          =   312
         Left            =   1320
         TabIndex        =   19
         Top             =   1320
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
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
      Begin XtremeSuiteControls.FlatEdit txtTelefonoExt 
         Height          =   312
         Left            =   2520
         TabIndex        =   20
         Top             =   1320
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFax 
         Height          =   312
         Left            =   4200
         TabIndex        =   21
         Top             =   1320
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
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
      Begin XtremeSuiteControls.FlatEdit txtFaxExt 
         Height          =   312
         Left            =   5400
         TabIndex        =   22
         Top             =   1320
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApartadoPostal 
         Height          =   312
         Left            =   7920
         TabIndex        =   23
         Top             =   1320
         Width           =   1932
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   312
         Left            =   1320
         TabIndex        =   24
         Top             =   1800
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
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
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   672
         Left            =   1320
         TabIndex        =   25
         Top             =   2160
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   1185
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
      Begin XtremeSuiteControls.FlatEdit txtObservacion 
         Height          =   792
         Left            =   1320
         TabIndex        =   26
         Top             =   2880
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   7920
         TabIndex        =   27
         Top             =   840
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3413
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
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   4200
         TabIndex        =   28
         Top             =   840
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2990
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
      End
      Begin XtremeSuiteControls.PushButton btnCuentas 
         Height          =   372
         Left            =   8160
         TabIndex        =   30
         Top             =   4200
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cuentas Bancarias"
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
      End
      Begin XtremeSuiteControls.ComboBox cboBancos 
         Height          =   312
         Left            =   3240
         TabIndex        =   31
         Top             =   4236
         Width           =   4812
         _Version        =   1441793
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
      Begin XtremeSuiteControls.ComboBox cboListadoTipo 
         Height          =   312
         Left            =   -62080
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3413
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
      End
      Begin XtremeSuiteControls.RadioButton Opt 
         Height          =   372
         Index           =   1
         Left            =   -68080
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Inactivos ?"
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
      Begin XtremeSuiteControls.ComboBox cboTipoPago 
         Height          =   312
         Left            =   8160
         TabIndex        =   32
         Top             =   3840
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Left            =   -66520
         TabIndex        =   42
         Top             =   1005
         Visible         =   0   'False
         Width           =   5655
         _Version        =   1441793
         _ExtentX        =   9975
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   255
         Left            =   -60400
         TabIndex        =   43
         ToolTipText     =   "Exportar Listado a Excel"
         Top             =   1005
         Visible         =   0   'False
         Width           =   255
         _Version        =   1441793
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   16
         Picture         =   "frmAF_Promotores.frx":0000
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Index           =   0
         Left            =   -70000
         TabIndex        =   44
         Top             =   960
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1441793
         _ExtentX        =   17595
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Listado de Promotores                       Filtros:"
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   13
         Left            =   5040
         TabIndex        =   41
         Top             =   3840
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Emitir"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   11
         Left            =   1320
         TabIndex        =   33
         Top             =   4200
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta/Desembolso"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Observacion"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   9
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "E Mail"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   7
         Left            =   6360
         TabIndex        =   12
         Top             =   1320
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Apto. Postal"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   6
         Left            =   3240
         TabIndex        =   11
         Top             =   1320
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fax"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Teléfono"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   4
         Left            =   3240
         TabIndex        =   9
         Top             =   840
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   3
         Left            =   6360
         TabIndex        =   8
         Top             =   840
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
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
         Alignment       =   1
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   2
         Left            =   3240
         TabIndex        =   6
         Top             =   480
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Nombre"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Identificación"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1680
      TabIndex        =   34
      Top             =   480
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   2760
      TabIndex        =   35
      Top             =   480
      Width           =   6612
      _Version        =   1441793
      _ExtentX        =   11663
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
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   45
      Top             =   840
      Visible         =   0   'False
      Width           =   9975
      _Version        =   1441793
      _ExtentX        =   17595
      _ExtentY        =   238
      _StockProps     =   93
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   12
      Left            =   240
      TabIndex        =   36
      Top             =   480
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Ejecutivo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAF_PromotoresPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean, vPaso As Boolean

Private Sub btnExport_Click()
Call Excel_Exportar_Lsw(lsw, ProgressBarX)
End Sub

Private Sub cboBancos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And btnCuentas.Enabled Then btnCuentas.SetFocus
End Sub

Private Sub sbCuentas_Load()

On Error GoTo vError

lswCuentas.ListItems.Clear
If vCodigo > 0 Then
    strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
           & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
           & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
           & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
           & " where C.Identificacion = '" & Trim(txtCedJur.Text) & "' and C.Modulo = 'AFI'"
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswCuentas.ListItems.Add(, , rs!CUENTA_INTERNA)
           itmX.SubItems(1) = Trim(rs!Banco)
           itmX.SubItems(2) = rs!TipoDesc
           itmX.SubItems(3) = rs!cod_Divisa
           itmX.SubItems(4) = IIf(rs!CUENTA_INTERBANCA = 1, "Sí", "No")
           itmX.SubItems(5) = rs!Destino & ""
           itmX.SubItems(6) = IIf(rs!Activa = 1, "Activa", "Cerrada")
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           itmX.SubItems(8) = rs!Registro_Usuario & ""
     
       rs.MoveNext
    Loop
    rs.Close
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnCuentas_Click()

If vCodigo = 0 Then
   MsgBox "Consulte un Ejecutivo Primero...", vbExclamation
   tcMain.Item(0).Selected = True
   Exit Sub
End If

GLOBALES.gTag = Trim(txtCedJur.Text)
GLOBALES.gTag2 = "AFI"

frmCC_Cuentas_Bancarias.Show vbModal

Call sbCuentas_Load
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono1.SetFocus
End Sub

Private Sub cboListadoTipo_Click()
If vPaso Then Exit Sub

Call sbListado_Load

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub

Private Sub cboTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboBancos.SetFocus
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If txtCodigo = "" Or Not IsNumeric(txtCodigo) Then txtCodigo.Text = "0"

If vScroll Then
    strSQL = "select Top 1 id_promotor from promotores"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where id_promotor > " & txtCodigo & " order by id_promotor asc"
    Else
       strSQL = strSQL & " where id_promotor < " & txtCodigo & " order by id_promotor desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!ID_PROMOTOR)
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 1

On Error GoTo vError
 
vScroll = False
 FlatScrollBar.Value = 0
vScroll = True
 
 
vPaso = True
 
lswCuentas.ColumnHeaders.Add 1, , "Cuenta", 2500
lswCuentas.ColumnHeaders.Add 2, , "Banco", 3500
lswCuentas.ColumnHeaders.Add 3, , "Tipo", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 4, , "Divisa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 5, , "Interbanca", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 6, , "Destino", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 7, , "Activa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 8, , "Fecha", 2500
lswCuentas.ColumnHeaders.Add 9, , "Usuario", 2500

lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add 1, , "Ejecutivo Id", 1000
lsw.ColumnHeaders.Add 2, , "Identificación", 1800, vbCenter
lsw.ColumnHeaders.Add 3, , "Nombre", 4000
lsw.ColumnHeaders.Add 4, , "Estado", 1500
lsw.ColumnHeaders.Add 5, , "Ingreso", 1300, vbCenter
lsw.ColumnHeaders.Add 6, , "Comisión", 1500, vbCenter


strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBancos, strSQL, False, True)
 
 
cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "Inactivo"
 
cboTipo.Clear
cboTipo.AddItem "Promotor"
cboTipo.AddItem "Comité"
cboTipo.AddItem "Externo"
 
cboListadoTipo.Clear
cboListadoTipo.AddItem "Promotor"
cboListadoTipo.AddItem "Comité"
cboListadoTipo.AddItem "Externo"
cboListadoTipo.Text = "Promotor"



cboTipoPago.Clear
cboTipoPago.AddItem fxTipoDocumento("CK")
cboTipoPago.AddItem fxTipoDocumento("TE")

vPaso = False

 
 vEdita = True
 
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()

tcMain.Item(0).Selected = True

vCodigo = 0
txtCodigo = ""

StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""

cboEstado.Text = "Activo"
cboTipo.Text = "Promotor"

chkComision.Value = vbUnchecked

cboTipoPago.Text = fxTipoDocumento("TE")

txtPagarA.Text = ""
txtNombre.Text = ""
txtObservacion.Text = ""

txtCedJur.Text = ""
txtUsuario.Text = ""

txtDireccion.Text = ""
txtApartadoPostal.Text = ""
txtEmail.Text = ""

txtTelefono1.Text = ""
txtTelefonoExt.Text = ""
txtFax.Text = ""
txtFaxExt.Text = ""

lswCuentas.ListItems.Clear

txtCodigo.Enabled = True

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lsw.ListItems.Count > 0 Then
    Call sbConsulta(Item.Text)
End If
End Sub

Private Sub opt_Click(Index As Integer)

Call sbListado_Load

End Sub


Private Sub sbListado_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

    lsw.ListItems.Clear

    strSQL = "Select P.*,B.descripcion as Banco" _
           & " from Promotores P inner join Tes_Bancos B on P.cod_banco = B.id_Banco" _
           & " where P.estado = " & IIf(opt(0).Value, 1, 0) & " and Tipo = '" & Mid(cboListadoTipo.Text, 1, 1) _
           & "'  and P.Nombre like '%" & txtFiltro.Text & "%'" _
           & " order by P.nombre"
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
        Set itmX = lsw.ListItems.Add(, , rs!ID_PROMOTOR)
            itmX.SubItems(1) = rs!Cod_Comision & ""
            itmX.SubItems(2) = rs!Nombre
            itmX.SubItems(3) = IIf((rs!Estado = 0), "Inactivo", "Activo")
            itmX.SubItems(4) = Format(rs!FECHAING, "dd/mm/yyyy")
            itmX.SubItems(5) = IIf(rs!apl_Comision = 1, "Sí", "No")
        rs.MoveNext
     Loop
     rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
    Call sbListado_Load
End If

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      txtNombre.SetFocus
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
         gBusquedas.Columna = "nombre"
         gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select id_promotor,nombre from promotores"
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

Private Sub sbConsulta(pEjecutivo As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select P.*,B.descripcion as Banco from promotores P inner join Tes_Bancos B on P.cod_banco = B.id_banco" _
       & " where P.id_promotor = " & pEjecutivo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  
  tcMain.Item(0).Selected = True
  
  vEdita = True
  vCodigo = rs!ID_PROMOTOR
  txtCodigo.Text = CStr(rs!ID_PROMOTOR)
  
    txtNombre.Text = rs!Nombre & ""
    txtObservacion.Text = rs!observacion & ""
    
    txtCedJur.Text = rs!Cod_Comision & ""
    txtUsuario.Text = rs!User_Referencia & ""
    
    If rs!Estado = 1 Then
      cboEstado.Text = "Activo"
    Else
      cboEstado.Text = "Inactivo"
    End If
    
    cboTipoPago.Text = fxTipoDocumento(rs!TIPO_DOCUMENTO)
    
    Select Case rs!Tipo
      Case "P"
        cboTipo.Text = "Promotor"
      Case "C"
        cboTipo.Text = "Comité"
      Case "E"
        cboTipo.Text = "Externo"
      Case Else
        cboTipo.Text = "Promotor"
    End Select
    
    chkComision.Value = rs!apl_Comision
    
    txtDireccion.Text = rs!direccion & ""
    txtApartadoPostal.Text = rs!aptopostal & ""
    txtEmail.Text = rs!Email & ""
    
    txtTelefono1.Text = rs!telefono & ""
    txtTelefonoExt.Text = rs!telefono_ext & ""
    txtFax.Text = rs!fax & ""
    txtFaxExt.Text = rs!fax_ext & ""
    
    
    txtPagarA.Text = rs!Nombre_Contacto & ""
    
    Call sbCboAsignaDato(cboBancos, Trim(rs!Banco), True, rs!cod_banco)


   StatusBarX.Panels(1).Text = rs!Usuario & ""
   StatusBarX.Panels(2).Text = rs!fecha & ""

   Call sbCuentas_Load
Else
  
  MsgBox "No se encontró registro verifique...", vbInformation
End If

Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If cboBancos.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No se especificó una Cuenta Bancaria para Desembolsos ..."

If Mid(cboTipoPago.Text, 1, 1) = "T" _
   And lswCuentas.ListItems.Count = 0 Then vMensaje = vMensaje & vbCrLf & " - No se especificó la cuenta para las transferencias..."

If txtNombre.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion del Ejecutivo de Cuenta no es válido ..."
If txtPagarA.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Ejecutivo de Cuenta no es válido ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()

On Error GoTo vError

If vEdita Then
  strSQL = "update promotores set nombre = '" & Trim(txtNombre.Text) _
         & "',cedula_contacto = '" & txtCedJur.Text & "',nombre_contacto = '" & txtPagarA.Text _
         & "',observacion = '" & txtObservacion.Text & "',estado = " & IIf(Mid(cboEstado.Text, 1, 1) = "A", 1, 0) _
         & ",tipo_documento = '" & fxTipoDocumento(cboTipoPago.Text) _
         & "',direccion = '" & txtDireccion.Text & "',aptoPostal = '" & txtApartadoPostal.Text _
         & "',email = '" & txtEmail.Text & "',telefono = '" & txtTelefono1.Text _
         & "',telefono_ext = '" & txtTelefonoExt.Text & "',fax = '" & txtFax.Text _
         & "',fax_ext = '" & txtFaxExt.Text & "',cod_banco = " & cboBancos.ItemData(cboBancos.ListIndex) _
         & ",comite = " & IIf(cboTipo.Text = "Comité", 1, 0) _
         & ",apl_comision = " & chkComision.Value & ",cod_comision = '" & txtCedJur.Text _
         & "',Tipo = '" & Mid(cboTipo.Text, 1, 1) & "', user_referencia = '" & txtUsuario.Text _
         & "',usuario = '" & glogon.Usuario & "',fecha = dbo.MyGetdate()" _
         & " where id_promotor = " & vCodigo
         
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Ejecutivo de Cuenta Id: " & vCodigo)

Else
   
   strSQL = "insert into promotores(Tipo,nombre,observacion,cod_comision,fechaIng" _
          & ",estado,telefono,telefono_ext,fax,fax_ext,email,aptopostal,direccion,tipo_documento" _
          & ",cod_banco,cedula_contacto,nombre_contacto,comite,apl_comision,usuario,fecha,user_referencia)" _
          & " values('" & Mid(cboTipo.Text, 1, 1) & "','" & txtNombre.Text & "','" & txtObservacion.Text & "','" & txtCedJur.Text & "',dbo.MyGetdate()," _
          & IIf((Mid(cboEstado.Text, 1, 1) = "A"), 1, 0) & ",'" & txtTelefono1.Text & "','" & txtTelefonoExt.Text & "','" & txtFax.Text & "','" & txtFaxExt.Text & "','" _
          & txtEmail & "','" & txtApartadoPostal & "','" & txtDireccion & "','" & fxTipoDocumento(cboTipoPago.Text) _
          & "'," & cboBancos.ItemData(cboBancos.ListIndex) & ",'" & txtCedJur.Text & "','" & txtPagarA.Text _
          & "'," & IIf(cboTipo.Text = "Comité", 1, 0) & "," & chkComision.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'" _
          & txtUsuario.Text & "')"
   
   Call ConectionExecute(strSQL)
    
   strSQL = "select isnull(max(id_promotor),0) as ultimo from promotores"
   Call OpenRecordSet(rs, strSQL)
     txtCodigo.Text = CStr(rs!ultimo)
     vCodigo = txtCodigo
   rs.Close
    
   Call Bitacora("Registra", "Ejecutivo de Cuenta Id: " & vCodigo)
    
   txtCodigo.Enabled = True
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbToolBar(tlb, "activo")
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
  strSQL = "delete promotores where id_promotor = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Ejecutivo de Cuenta Id: " & vCodigo)
  
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case ButtonMenu.Key
Case "AFRA"
 GLOBALES.gstrReporte = "Resumen"
 frmAF_PromotoresReportes.Show vbModal

Case "AFDA"
 GLOBALES.gstrReporte = "Detalle"
 frmAF_PromotoresReportes.Show vbModal

Case "AFLP"
   With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Personas"
        
    .Connect = glogon.ConectRPT
        
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .ReportFileName = SIFGlobal.fxPathReportes("Personas_ListadoPromotores.rpt")
    
    .PrintReport
   End With
   
End Select
Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtApartadoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub

Private Sub txtCedJur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPagarA.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Ejectuvo Id"
  gBusquedas.Col2Name = "Identificación"
  gBusquedas.Columna = "id_promotor"
  gBusquedas.Orden = "id_promotor"
  gBusquedas.Consulta = "select id_promotor,cod_Comision,nombre from promotores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub


Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservacion.SetFocus
End Sub


Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFaxExt.SetFocus
End Sub

Private Sub txtFaxExt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApartadoPostal.SetFocus
End Sub


Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbListado_Load
End If
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedJur.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Ejectuvo Id"
  gBusquedas.Col2Name = "Identificación"
  gBusquedas.Columna = "id_promotor"
  gBusquedas.Orden = "id_promotor"
  gBusquedas.Consulta = "select id_promotor,cod_Comision,nombre from promotores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoPago.SetFocus
End Sub

Private Sub txtPagarA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUsuario.SetFocus
End Sub

Private Sub txtTelefono1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefonoExt.SetFocus
End Sub

Private Sub txtTelefonoExt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFax.SetFocus
End Sub


Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Nombre"
  gBusquedas.Consulta = "select Nombre,Descripcion from usuarios"
  gBusquedas.Filtro = " and estado = 'A'"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Resultado = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
     txtUsuario.Text = gBusquedas.Resultado
  End If


End If

End Sub
