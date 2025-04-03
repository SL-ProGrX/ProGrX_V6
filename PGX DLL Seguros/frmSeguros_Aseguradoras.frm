VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmSeguros_Aseguradoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Aseguradoras"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   10080
      Top             =   240
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
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
      Left            =   10560
      TabIndex        =   1
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11280
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Aseguradoras.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Aseguradoras.frx":3492
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Aseguradoras.frx":6924
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Aseguradoras.frx":6A42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   6735
      _Version        =   1441792
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   2055
      _Version        =   1441792
      _ExtentX        =   3619
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5415
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   10935
      _Version        =   1441792
      _ExtentX        =   19283
      _ExtentY        =   9546
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
      Item(0).Caption =   "General"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "gbDireccion"
      Item(0).Control(1)=   "chkActivo"
      Item(0).Control(2)=   "GroupBox3"
      Item(0).Control(3)=   "txtContacto"
      Item(0).Control(4)=   "Label4(8)"
      Item(0).Control(5)=   "txtCedJuridica"
      Item(0).Control(6)=   "Label4(0)"
      Item(1).Caption =   "Cuentas"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "lswCuentas"
      Item(1).Control(1)=   "btnCuentas"
      Item(1).Control(2)=   "cboBancos"
      Item(1).Control(3)=   "Label3"
      Item(1).Control(4)=   "GroupBox1(0)"
      Item(1).Control(5)=   "GroupBox1(1)"
      Begin XtremeSuiteControls.ListView lswCuentas 
         Height          =   1575
         Left            =   -69880
         TabIndex        =   34
         Top             =   870
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1441792
         _ExtentX        =   18865
         _ExtentY        =   2778
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
      Begin XtremeSuiteControls.GroupBox gbDireccion 
         Height          =   1695
         Left            =   0
         TabIndex        =   6
         Top             =   3000
         Width           =   10815
         _Version        =   1441792
         _ExtentX        =   19071
         _ExtentY        =   2984
         _StockProps     =   79
         Caption         =   "Dirección"
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
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   312
            Left            =   1440
            TabIndex        =   7
            Top             =   480
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.ComboBox cboCanton 
            Height          =   312
            Left            =   1440
            TabIndex        =   8
            Top             =   840
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.ComboBox cboDistrito 
            Height          =   312
            Left            =   1440
            TabIndex        =   9
            Top             =   1200
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.FlatEdit txtDireccion 
            Height          =   1092
            Left            =   3840
            TabIndex        =   10
            Top             =   480
            Width           =   6852
            _Version        =   1441792
            _ExtentX        =   12086
            _ExtentY        =   1926
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   11
            Left            =   120
            TabIndex        =   13
            Top             =   1200
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Distrito"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   12
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cantón"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   13
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Provincia"
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
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   255
         Left            =   9480
         TabIndex        =   14
         Top             =   480
         Width           =   1095
         _Version        =   1441792
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activa?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1815
         Left            =   0
         TabIndex        =   15
         Top             =   1080
         Width           =   10815
         _Version        =   1441792
         _ExtentX        =   19071
         _ExtentY        =   3196
         _StockProps     =   79
         Caption         =   "Información de Contacto"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtWebSite 
            Height          =   312
            Left            =   5400
            TabIndex        =   16
            Top             =   360
            Width           =   5292
            _Version        =   1441792
            _ExtentX        =   9334
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
            Left            =   5400
            TabIndex        =   17
            Top             =   720
            Width           =   5292
            _Version        =   1441792
            _ExtentX        =   9334
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
         Begin XtremeSuiteControls.FlatEdit txtEmail2 
            Height          =   312
            Left            =   5400
            TabIndex        =   18
            Top             =   1080
            Width           =   5292
            _Version        =   1441792
            _ExtentX        =   9334
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
         Begin XtremeSuiteControls.FlatEdit txtAptoPostal 
            Height          =   312
            Left            =   5400
            TabIndex        =   19
            Top             =   1440
            Width           =   5292
            _Version        =   1441792
            _ExtentX        =   9334
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
         Begin XtremeSuiteControls.FlatEdit txtTelefono 
            Height          =   312
            Left            =   1440
            TabIndex        =   20
            Top             =   360
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3619
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
         Begin XtremeSuiteControls.FlatEdit txtTelefono2 
            Height          =   312
            Left            =   1440
            TabIndex        =   21
            Top             =   720
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3619
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
         Begin XtremeSuiteControls.FlatEdit txtTelFax 
            Height          =   312
            Left            =   1440
            TabIndex        =   22
            Top             =   1080
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3619
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   7
            Left            =   3840
            TabIndex        =   29
            Top             =   1440
            Width           =   1332
            _Version        =   1441792
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   6
            Left            =   3840
            TabIndex        =   28
            Top             =   1080
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Email (2)"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   5
            Left            =   3840
            TabIndex        =   27
            Top             =   720
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Email (1)"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   4
            Left            =   3840
            TabIndex        =   26
            Top             =   360
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Web Site"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   3
            Left            =   0
            TabIndex        =   25
            Top             =   1080
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tel. Fax"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   2
            Left            =   0
            TabIndex        =   24
            Top             =   720
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Teléfono (2)"
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   360
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Teléfono (1)"
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
      Begin XtremeSuiteControls.FlatEdit txtContacto 
         Height          =   312
         Left            =   1440
         TabIndex        =   30
         Top             =   4800
         Width           =   9252
         _Version        =   1441792
         _ExtentX        =   16319
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
      Begin XtremeSuiteControls.FlatEdit txtCedJuridica 
         Height          =   315
         Left            =   1440
         TabIndex        =   32
         Top             =   480
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3619
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
      Begin XtremeSuiteControls.PushButton btnCuentas 
         Height          =   375
         Left            =   -60880
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441792
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
         Height          =   330
         Left            =   -66880
         TabIndex        =   36
         Top             =   510
         Visible         =   0   'False
         Width           =   5295
         _Version        =   1441792
         _ExtentX        =   9340
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1215
         Index           =   0
         Left            =   -69880
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1441792
         _ExtentX        =   18865
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Códigos de Deducción"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboFormato 
            Height          =   315
            Left            =   1920
            TabIndex        =   39
            Top             =   840
            Width           =   2295
            _Version        =   1441792
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.FlatEdit txtRetCod 
            Height          =   315
            Left            =   1920
            TabIndex        =   44
            Top             =   120
            Width           =   2295
            _Version        =   1441792
            _ExtentX        =   4048
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtRetDesc 
            Height          =   315
            Left            =   4200
            TabIndex        =   45
            Top             =   120
            Width           =   6495
            _Version        =   1441792
            _ExtentX        =   11456
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            Caption         =   "Formato Trama"
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
            Index           =   18
            Left            =   480
            TabIndex        =   41
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Retención"
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
            Left            =   480
            TabIndex        =   40
            Top             =   480
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1335
         Index           =   1
         Left            =   -69880
         TabIndex        =   42
         Top             =   3960
         Visible         =   0   'False
         Width           =   10695
         _Version        =   1441792
         _ExtentX        =   18865
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Cuenta Contable"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCuentaCod 
            Height          =   315
            Left            =   1920
            TabIndex        =   46
            Top             =   480
            Width           =   2295
            _Version        =   1441792
            _ExtentX        =   4048
            _ExtentY        =   556
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
            Height          =   315
            Left            =   4200
            TabIndex        =   47
            Top             =   480
            Width           =   6495
            _Version        =   1441792
            _ExtentX        =   11456
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   480
            TabIndex        =   43
            Top             =   480
            Width           =   975
         End
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
         Height          =   255
         Left            =   -69760
         TabIndex        =   37
         Top             =   510
         Visible         =   0   'False
         Width           =   2655
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   33
         Top             =   480
         Width           =   1335
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ced. Juridica"
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
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   4800
         Width           =   1332
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Contacto"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Aseguradora"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   14
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmSeguros_Aseguradoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vCantonMascara As String, vDistritoMascara As String, vFechaActual As Date
Dim vScroll As Boolean, vPaso As Boolean

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Sub sbCuentas_Load()

On Error GoTo vError

lswCuentas.ListItems.Clear
    strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
           & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
           & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
           & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
           & " where C.Identificacion = '" & Trim(txtCedJuridica.Text) & "' and C.Modulo = 'CdS'"
    
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



Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub btnCuentas_Click()
If txtCodigo.Text = "" Then
   MsgBox "Consulte una Aseguradora Primero...", vbExclamation
   tcMain.Item(0).Selected = True
   Exit Sub
End If

GLOBALES.gTag = Trim(txtCedJuridica)
GLOBALES.gTag2 = "CdS"

frmCC_Cuentas_Bancarias.Show vbModal

Call sbCuentas_Load
End Sub

Private Sub cboCanton_Click()

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
            & "' and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistrito, strSQL, False, True)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistrito.AddItem " "
cboDistrito.Text = " "
End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub

Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub



Private Sub cboProvincia_Click()

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click


End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cod_aseguradora from SEGUROS_ASEGURADORAS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_aseguradora > '" & txtCodigo.Text & "' order by cod_aseguradora asc"
    Else
       strSQL = strSQL & " where cod_aseguradora < '" & txtCodigo.Text & "' order by cod_aseguradora desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!cod_Aseguradora
      Call txtCodigo_LostFocus
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
vModulo = 17
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 17

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
tcMain.Item(0).Selected = True

vEdita = False

lswCuentas.ColumnHeaders.Add 1, , "Cuenta", 2500
lswCuentas.ColumnHeaders.Add 2, , "Banco", 3500
lswCuentas.ColumnHeaders.Add 3, , "Tipo", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 4, , "Divisa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 5, , "Interbanca", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 6, , "Destino", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 7, , "Activa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 8, , "Fecha", 2500
lswCuentas.ColumnHeaders.Add 9, , "Usuario", 2500

vPaso = True
    strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
    Call sbCbo_Llena_New(cboBancos, strSQL, False, True)
vPaso = False

Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
vCodigo = ""
txtCodigo = ""

txtCedJuridica.Text = ""

txtNombre = ""
txtTelefono.Text = ""
txtTelefono2.Text = ""
txtWebSite.Text = ""
txtEmail.Text = ""
txtEmail2.Text = ""
txtAptoPostal.Text = ""

txtDireccion = ""
txtContacto.Text = ""

txtRetCod.Text = ""
txtRetDesc.Text = ""
txtCuentaCod.Text = ""
txtCuentaDesc.Text = ""

chkActivo.Value = vbChecked

cboFormato.Clear
cboFormato.AddItem "ISABS"
cboFormato.AddItem "SISEPO"
cboFormato.Text = "ISABS"

tcMain.Item(0).Selected = True

End Sub




Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 1 Then
  Call sbCuentas_Load
End If
End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

vFechaActual = Format(fxFechaServidor, "dd/mm/yyyy")

vPaso = True
'Call sbCargaCbo(cboProvincia, "provincias")
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select cod_aseguradora,nombre from SEGUROS_ASEGURADORAS"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(xCodigo As String)


On Error GoTo vError

If Not fxSIFValidaCadena(txtCodigo.Text) Then
   Exit Sub
End If

Me.MousePointer = vbHourglass
'       & ",case when P.Tipo_Emision = 'TE' then 'Transferencia' when P.Tipo_Emision = 'CK' then 'Cheque' else 'Transferencia' end as 'Tipo_Pago' " _

strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & ",isnull(Cat.Descripcion,'') as 'RetDesc', isnull(Ban.Descripcion, '') as 'BancoDesc'" _
       & " from SEGUROS_ASEGURADORAS P " _
       & " left join Provincias Prov on P.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Provincia = Cant.Provincia and P.Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Provincia = Dist.Provincia and convert(int,P.Canton) = convert(int,Dist.Canton) and P.distrito = Dist.distrito" _
       & " left join Catalogo Cat on P.codigo_Retencion = Cat.Codigo" _
       & " left join Tes_Bancos Ban on P.Cod_Banco = Ban.Id_Banco" _
       & " where P.cod_aseguradora = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCodigo = rs!cod_Aseguradora
  txtCodigo = rs!cod_Aseguradora

  txtNombre = rs!Nombre & ""
  
  txtCedJuridica.Text = rs!cedula_juridica & ""
  txtTelefono.Text = rs!Telefono_01 & ""
  txtTelefono2.Text = rs!Telefono_02 & ""
  txtTelFax.Text = rs!Tel_Fax & ""


  txtWebSite.Text = rs!Sitio_Web & ""
  txtEmail.Text = rs!Email_01 & ""
  txtEmail2.Text = rs!Email_02 & ""
  txtAptoPostal.Text = rs!apto_postal & ""

  txtContacto.Text = rs!Nombre_Contacto & ""

  Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")
  Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")
  Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
     
  cboDistrito.ToolTipText = Trim(rs!distrito) & ""
  txtDireccion.Text = rs!Direccion

'  txtBancoId.Text = rs!Cod_Banco & ""
'  txtBancoDesc.Text = rs!BancoDesc & ""
'  txtCuentaCliente.Text = rs!Cuenta_Cliente & ""
    
  Call sbCboAsignaDato(cboBancos, rs!BancoDesc, True, rs!Cod_Banco & "")

 
  txtRetCod.Text = rs!Codigo_Retencion & ""
  txtRetDesc.Text = rs!RetDesc

  txtCuentaDesc.Text = fxgCntCuentaDesc(rs!cod_cuenta & "")
  txtCuentaCod.Text = fxgCntCuentaFormato(True, rs!cod_cuenta & "", 0)
  
  Call sbCboAsignaDato(cboFormato, rs!Formato_Tramas & "")

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre no es válido ..."
If txtCedJuridica.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Cédula Jurídica no es válida! ..."

If Trim(txtCuentaDesc.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Cuenta Contable no es válida!..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim vEstadoCivil As String

On Error GoTo vError


If vEdita Then
  strSQL = "update SEGUROS_ASEGURADORAS set nombre = '" & Trim(txtNombre.Text) & "', Cedula_Juridica = '" & txtCedJuridica.Text _
         & "',Telefono_01 = '" & txtTelefono.Text & "',Telefono_02 = '" & txtTelefono2.Text _
         & "',Tel_Fax = '" & txtTelFax.Text & "',Sitio_Web = '" & txtWebSite.Text & "',apto_postal = '" & txtAptoPostal _
         & "',email_01 = '" & txtEmail & "', email_02 = '" & txtEmail2.Text & "',direccion = '" & txtDireccion _
         & "',distrito = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "',canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
         & "',provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
         & "',Nombre_Contacto = '" & txtContacto.Text & "',Activo = " & chkActivo.Value _
         & ",Codigo_Retencion = '" & txtRetCod.Text & "',formato_Tramas = '" & cboFormato.Text _
         & "', cod_Banco = " & cboBancos.ItemData(cboBancos.ListIndex) _
         & " , Cod_Cuenta = '" & fxgCntCuentaFormato(False, txtCuentaCod.Text, 0) & "'" _
         & " where cod_aseguradora = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Aseguradora: " & vCodigo)

Else
  vCodigo = txtCodigo

   strSQL = "insert into SEGUROS_ASEGURADORAS(cod_aseguradora,nombre,cedula_Juridica,Telefono_01,Telefono_02,Tel_fax,Activo,registro_fecha,registro_usuario" _
          & ",apto_postal,email_01,email_02,Sitio_Web,direccion,distrito,provincia,canton,nombre_contacto,cod_banco,cod_Cuenta, formato_Tramas)" _
          & " values('" & vCodigo & "','" & txtNombre.Text & "','" & txtCedJuridica.Text & "','" & txtTelefono.Text & "','" & txtTelefono2.Text & "','" & txtTelFax.Text _
          & "'," & chkActivo.Value & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtAptoPostal.Text & "','" & txtEmail.Text & "','" & txtEmail2.Text & "','" & txtWebSite.Text _
          & "','" & txtDireccion.Text & "','" & cboDistrito.ItemData(cboDistrito.ListIndex) & "','" _
          & cboProvincia.ItemData(cboProvincia.ListIndex) & "','" & cboCanton.ItemData(cboCanton.ListIndex) _
          & "','" & txtContacto.Text & "'," & cboBancos.ItemData(cboBancos.ListIndex) _
          & ",'" & fxgCntCuentaFormato(False, txtCuentaCod.Text) & "','" & cboFormato.Text & "')"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Aseguradora: " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(txtCodigo.Text)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete SEGUROS_ASEGURADORAS where cod_aseguradora = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Aseguradora: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCedJuridica_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_aseguradora"
  gBusquedas.Orden = "cod_aseguradora"
  gBusquedas.Consulta = "select cod_aseguradora,nombre from SEGUROS_ASEGURADORAS"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()
  Call sbConsulta(txtCodigo.Text)
'  txtNombre.SetFocus
End Sub




Private Sub txtCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    gBusquedas.Resultado = gCuenta
    txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
    txtCuentaCod.Text = fxgCntCuentaFormato(True, gCuenta, 0)
End If

End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContacto.SetFocus
End Sub

Private Sub txtEMail2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAptoPostal.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedJuridica.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cod_aseguradora,nombre from SEGUROS_ASEGURADORAS"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtRetCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRetDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Codigo"
  gBusquedas.Orden = "Codigo"
  gBusquedas.Consulta = "select Codigo,Descripcion from Catalogo"
  gBusquedas.Filtro = " and Retencion = 'S'"
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtRetCod.Text = gBusquedas.Resultado
    txtRetDesc.Text = gBusquedas.Resultado2
  End If
End If

End Sub


Private Sub txtRetDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaCod.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Consulta = "select Codigo,Descripcion from Catalogo"
  gBusquedas.Filtro = " and Retencion = 'S'"
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtRetCod.Text = gBusquedas.Resultado
    txtRetDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelFax.SetFocus
End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail2.SetFocus
End Sub


Private Sub txtTelFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtWebSite.SetFocus
End Sub

Private Sub txtWebSite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub






