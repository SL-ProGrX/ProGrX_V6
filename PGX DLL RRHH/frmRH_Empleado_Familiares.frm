VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmRH_Empleado_Familiares 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contactos y Familares"
   ClientHeight    =   8235
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7812
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9852
      _Version        =   1310720
      _ExtentX        =   17378
      _ExtentY        =   13779
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
      Item(0).Caption =   "Listado"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   17
      Item(1).Control(0)=   "txtApellido1"
      Item(1).Control(1)=   "txtApellido2"
      Item(1).Control(2)=   "txtNombre"
      Item(1).Control(3)=   "txtCedula"
      Item(1).Control(4)=   "dtpFechaNacimiento"
      Item(1).Control(5)=   "Label2"
      Item(1).Control(6)=   "Label3"
      Item(1).Control(7)=   "Label4(1)"
      Item(1).Control(8)=   "Lbl4"
      Item(1).Control(9)=   "Lbl3(0)"
      Item(1).Control(10)=   "Lbl1"
      Item(1).Control(11)=   "cboParentesco"
      Item(1).Control(12)=   "Label15(1)"
      Item(1).Control(13)=   "txtCodigo"
      Item(1).Control(14)=   "gbPersona(1)"
      Item(1).Control(15)=   "cboSexo"
      Item(1).Control(16)=   "Label14"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   6852
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   9732
         _Version        =   1310720
         _ExtentX        =   17166
         _ExtentY        =   12086
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboParentesco 
         Height          =   312
         Left            =   -68320
         TabIndex        =   2
         Top             =   1800
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1310720
         _ExtentX        =   4048
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
         Appearance      =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   315
         Left            =   -68320
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1310720
         _ExtentX        =   4043
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaNacimiento 
         Height          =   312
         Left            =   -62200
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   -63760
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1310720
         _ExtentX        =   5101
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido2 
         Height          =   312
         Left            =   -66040
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1310720
         _ExtentX        =   4043
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido1 
         Height          =   312
         Left            =   -68320
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1310720
         _ExtentX        =   4043
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   -61720
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   852
         _Version        =   1310720
         _ExtentX        =   1503
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox gbPersona 
         Height          =   5052
         Index           =   1
         Left            =   -68320
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1310720
         _ExtentX        =   13568
         _ExtentY        =   8911
         _StockProps     =   79
         Caption         =   "Datos de Localización"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   312
            Left            =   1440
            TabIndex        =   17
            Top             =   3120
            Width           =   1692
            _Version        =   1310720
            _ExtentX        =   2990
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.ComboBox cboCanton 
            Height          =   312
            Left            =   3240
            TabIndex        =   18
            Top             =   3120
            Width           =   1932
            _Version        =   1310720
            _ExtentX        =   3413
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.ComboBox cboDistrito 
            Height          =   312
            Left            =   5280
            TabIndex        =   19
            Top             =   3120
            Width           =   2052
            _Version        =   1310720
            _ExtentX        =   3625
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail 
            Height          =   312
            Left            =   1440
            TabIndex        =   20
            Top             =   1320
            Width           =   5892
            _Version        =   1310720
            _ExtentX        =   10393
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail_02 
            Height          =   312
            Left            =   1440
            TabIndex        =   21
            Top             =   1680
            Width           =   5892
            _Version        =   1310720
            _ExtentX        =   10393
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtApartado 
            Height          =   312
            Left            =   1440
            TabIndex        =   22
            Top             =   2520
            Width           =   5892
            _Version        =   1310720
            _ExtentX        =   10393
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDireccion 
            Height          =   912
            Left            =   1440
            TabIndex        =   23
            Top             =   3480
            Width           =   5892
            _Version        =   1310720
            _ExtentX        =   10393
            _ExtentY        =   1609
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtWebSite 
            Height          =   312
            Left            =   1440
            TabIndex        =   24
            Top             =   2160
            Width           =   5892
            _Version        =   1310720
            _ExtentX        =   10393
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelMovil 
            Height          =   312
            Left            =   1440
            TabIndex        =   25
            Top             =   480
            Width           =   1692
            _Version        =   1310720
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono1 
            Height          =   312
            Left            =   1440
            TabIndex        =   26
            Top             =   840
            Width           =   1692
            _Version        =   1310720
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono2 
            Height          =   312
            Left            =   5640
            TabIndex        =   27
            Top             =   480
            Width           =   1692
            _Version        =   1310720
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   552
            Left            =   1440
            TabIndex        =   38
            Top             =   4440
            Width           =   5892
            _Version        =   1310720
            _ExtentX        =   10393
            _ExtentY        =   974
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtTelFax 
            Height          =   312
            Left            =   5640
            TabIndex        =   40
            Top             =   840
            Width           =   1692
            _Version        =   1310720
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono/Fax"
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
            Left            =   4200
            TabIndex        =   41
            Top             =   840
            Width           =   1332
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas"
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
            Left            =   0
            TabIndex        =   39
            Top             =   4440
            Width           =   732
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. Móvil"
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
            TabIndex        =   35
            Top             =   480
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono Trab."
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
            Left            =   4200
            TabIndex        =   34
            Top             =   480
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono Hab."
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
            Left            =   0
            TabIndex        =   33
            Top             =   840
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Email No.1"
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
            Left            =   0
            TabIndex        =   32
            Top             =   1320
            Width           =   1092
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Site / Blog"
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
            Left            =   0
            TabIndex        =   31
            Top             =   2160
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Email No.2"
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
            Left            =   0
            TabIndex        =   30
            Top             =   1680
            Width           =   1092
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
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
            Height          =   252
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   3120
            Width           =   732
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
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
            Height          =   252
            Left            =   0
            TabIndex        =   28
            Top             =   2520
            Width           =   1332
         End
      End
      Begin XtremeSuiteControls.ComboBox cboSexo 
         Height          =   312
         Left            =   -65200
         TabIndex        =   36
         Top             =   1800
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2566
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
         Appearance      =   2
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Genero"
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
         Left            =   -65920
         TabIndex        =   37
         Top             =   1800
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Apellido 1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   312
         Left            =   -68320
         TabIndex        =   15
         Top             =   1020
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Apellido 2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   312
         Left            =   -66040
         TabIndex        =   14
         Top             =   1020
         Visible         =   0   'False
         Width           =   2412
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   312
         Index           =   1
         Left            =   -63760
         TabIndex        =   13
         Top             =   1020
         Visible         =   0   'False
         Width           =   2892
      End
      Begin VB.Label Lbl4 
         Caption         =   "Nacimiento"
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
         Left            =   -63520
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Lbl3 
         Caption         =   "Parentesco"
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
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Lbl1 
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
         Height          =   252
         Left            =   -69640
         TabIndex        =   10
         Top             =   636
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Linea Id:"
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
         Left            =   -63160
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1092
      End
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   264
      Left            =   240
      TabIndex        =   42
      Top             =   120
      Width           =   3456
      _ExtentX        =   6085
      _ExtentY        =   476
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
            Object.Visible         =   0   'False
            Key             =   "Reportes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   8
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "socprov"
                  Text            =   "Asociados x Provincia"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "exsocprov"
                  Text            =   "Ex-Asociados x Provincia"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Xsep1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LisIng"
                  Text            =   "Listados Generales"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ImpBol"
                  Text            =   "Boleta"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Xsep2"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "socup"
                  Text            =   "Resumen x Unidad"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "desocup"
                  Text            =   "Detalle x Unidad"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRH_Empleado_Familiares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vEmpleadoId As String
Dim vPaso As Boolean

Private Sub sbCargaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "select Cod_Familiar,Identificacion,Nombre,Parentesco_Desc,fecha_nacimiento" _
       & " from  vRH_Personas_Familiares" _
       & " where Empleado_Id = '" & vEmpleadoId & "'"
Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Cod_Familiar)
    itmX.SubItems(1) = rs!Identificacion
    itmX.SubItems(2) = rs!Nombre
    itmX.SubItems(3) = Format(rs!fecha_nacimiento, "dd/mm/yyyy")
    itmX.SubItems(4) = rs!parentesco_Desc
   rs.MoveNext
Loop
     
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cboCanton_Click()
If vPaso Then Exit Sub

Dim strSQL As String

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

Dim strSQL As String

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click
End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 1

On Error GoTo vError

vEmpleadoId = GLOBALES.gCedulaActual
 
 vEdita = True
 Call sbToolBarIconos(tlbPrincipal)
 Call sbToolBar(tlbPrincipal, "nuevo")
 
 strSQL = "select rtrim(cod_Parentesco) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
        & " from sys_Parentescos where activo = 1"
 Call sbCbo_Llena_New(cboParentesco, strSQL, False, True)
 
 
vPaso = True

'Provincias
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False
 
cboSexo.AddItem "Masculino"
cboSexo.AddItem "Femenino"
cboSexo.Text = "Masculino"



 With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Identificación", 1400
    .Add , , "Nombre", 3400
    .Add , , "Fec.Nac.", 1400, vbCenter
    .Add , , "Parentesco", 1400
   
 End With
 
 Call sbLimpiaPantalla(0)

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla(Optional Index As Integer = 1)
Dim strSQL As String, rs As New ADODB.Recordset


tcMain.Item(Index).Selected = True

Select Case Index
    Case 0 'Lista
        Call sbCargaLsw
    Case 1 'Caso
        vCodigo = 0
        txtCodigo = ""
        
        strSQL = "select isnull(count(*),0) + 1 as Consec from RH_PERSONA_FAMILIARES where Empleado_Id = '" & vEmpleadoId & "'"
        Call OpenRecordSet(rs, strSQL)
          txtCedula = Trim(vEmpleadoId) & "-" & Format(rs!consec, "00")
        
        rs.Close
        
        txtApellido1 = ""
        txtApellido2 = ""
        txtNombre = ""
        
        dtpFechaNacimiento.MaxDate = fxFechaServidor
        dtpFechaNacimiento.Value = dtpFechaNacimiento.MaxDate

        
        txtNotas.Text = ""
        txtDireccion.Text = ""
        txtApartado.Text = ""
        txtEmail.Text = ""
        txtEmail_02.Text = ""
        txtTelefono1.Text = ""
        txtTelefono2.Text = ""
        txtTelMovil.Text = ""
        txtTelFax.Text = ""
        
        
End Select

End Sub



Private Sub lsw_DblClick()
If lsw.ListItems.Count > 0 Then
   Call sbConsulta(lsw.SelectedItem.Text)
End If
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 0 Then

    Call sbCargaLsw

End If

End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCedula.SetFocus
      Call sbToolBar(tlbPrincipal, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCedula.SetFocus
      Call sbToolBar(tlbPrincipal, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlbPrincipal, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlbPrincipal, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select consec,cedulaBN,nombre from beneficiarios"
       gBusquedas.Filtro = " and cedula = '" & vEmpleadoId & "'"
       frmBusquedas.Show vbModal
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       Call sbConsulta(txtCodigo)
        
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vApellido1 As String, vApellido2 As String, vNombre1 As String, vNombre2 As String
Dim vEspacio As Integer, i As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vRH_Personas_Familiares" _
       & " where Empleado_Id = '" & vEmpleadoId & "' and Cod_Familiar = " & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlbPrincipal, "activo")
  
  tcMain.Item(1).Selected = True
      
  vEdita = True
  vCodigo = rs!Cod_Familiar
  txtCodigo.Text = rs!Cod_Familiar
  
  txtCedula.Text = Trim(rs!Identificacion)
 
   vEspacio = 1
   For i = 1 To Len(Trim(rs!Nombre))
     If Mid(Trim(rs!Nombre), i, 1) <> " " Then
        Select Case vEspacio
         Case 1
          vApellido1 = vApellido1 & Mid(Trim(rs!Nombre), i, 1)
         Case 2
          vApellido2 = vApellido2 & Mid(Trim(rs!Nombre), i, 1)
         Case 3
          vNombre1 = vNombre1 & Mid(Trim(rs!Nombre), i, 1)
         Case Is >= 4
          vNombre2 = vNombre2 & Mid(Trim(rs!Nombre), i, 1)
        End Select
     Else
        vEspacio = vEspacio + 1
     End If
   Next i
   txtApellido1 = vApellido1
   txtApellido2 = vApellido2
   txtNombre = vNombre1 & " " & vNombre2
   
   dtpFechaNacimiento.Value = rs!fecha_nacimiento
   
   Call sbCboAsignaDato(cboParentesco, rs!parentesco_Desc, True, rs!cod_Parentesco)
    
   cboSexo.Text = IIf(rs!sexo = "M", "Masculino", "Femenino")
   
   'Contacto
   txtTelefono1.Text = rs!telefono1 & ""
   txtTelefono2.Text = rs!telefono2 & ""
   txtTelMovil.Text = rs!Tel_Movil & ""
   txtTelFax.Text = rs!Fax & ""
   
   txtEmail.Text = Trim(rs!Email_01 & "")
   txtEmail_02.Text = Trim(rs!Email_02 & "")
   txtApartado.Text = Trim(rs!apto_postal & "")
   
   txtWebSite.Text = Trim(rs!WebSite & "")
     
   Call sbCboAsignaDato(cboProvincia, rs!ProvinciaDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
   Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
   Call sbCboAsignaDato(cboDistrito, rs!DistritoDesc & "")
     
   txtDireccion = Trim(rs!Direccion) & ""
   txtNotas = Trim(rs!notas & "")
       
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""
fxValida = True

If cboParentesco.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No se ha seleccionado ningún parentesco..."

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Beneficiario no es válido ..."
If txtApellido1 = "" Then vMensaje = vMensaje & vbCrLf & " - txtApellido 1 del Beneficiario no es válido ..."
If txtApellido2 = "" Then vMensaje = vMensaje & vbCrLf & " - txtApellido 2 del Beneficiario no es válido ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

'spRH_PERSONA_FAMILIARES_Registra(@EmpleadoId varchar(20), @FamiliarId int, @Identificacion varchar(30), @Nombre varchar(100)
'                                                , @FechaNac datetime, @Sexo char(1), @Parentesco varchar(10)
'                                                , @Telefono1 varchar(10), @Telefono2 varchar(10), @Movil varchar(10), @Fax varchar(10)
'                                                , @Email_01 varchar(100), @Email_02 varchar(100), @WebSise varchar(100)
'                                                , @Provincia varchar(10), @Canton varchar(10), @Distrito varchar(10)
'                                                , @Notas varchar(500), @Direccion varchar(500), @AptoPostal varchar(100)
'                                                , @TipoMov char(1), @Usuario varchar(30))

strSQL = "exec spRH_PERSONA_FAMILIARES_Registra '" & vEmpleadoId & "'," & vCodigo & ",'" & txtCedula.Text _
                & "','" & Trim(txtApellido1) & " " & Trim(txtApellido2) & " " & Trim(txtNombre) _
                & "','" & Format(dtpFechaNacimiento.Value, "yyyy/mm/dd") & "','" & Mid(cboSexo.Text, 1, 1) _
                & "','" & cboParentesco.ItemData(cboParentesco.ListIndex) & "','" & txtTelefono1.Text _
                & "','" & txtTelefono2.Text & "','" & txtTelMovil.Text & "','" & txtTelFax.Text _
                & "','" & txtEmail.Text & "','" & txtEmail_02.Text & "','" & txtWebSite.Text _
                & "','" & cboProvincia.ItemData(cboProvincia.ListIndex) _
                & "','" & cboCanton.ItemData(cboCanton.ListIndex) _
                & "','" & cboDistrito.ItemData(cboDistrito.ListIndex) & "','" & txtNotas.Text _
                & "','" & txtDireccion.Text & "','" & txtApartado.Text _
                & "','A','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If Not glogon.error Then
  vCodigo = rs!FamiliarId
  txtCodigo.Text = rs!FamiliarId
  rs.Close

Else
  Exit Sub
End If

If vEdita Then
  
  Call Bitacora("Modifica", "Familiar Id:  " & txtCedula.Text & " Consec.: " & vCodigo)

End If


   
MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbToolBar(tlbPrincipal, "activo")
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
 
strSQL = "exec spRH_PERSONA_FAMILIARES_Registra '" & vEmpleadoId & "'," & vCodigo & ",'" & txtCedula.Text _
                & "','" & Trim(txtApellido1) & " " & Trim(txtApellido2) & " " & Trim(txtNombre) _
                & "','" & Format(dtpFechaNacimiento.Value, "yyyy/mm/dd") & "','" & Mid(cboSexo.Text, 1, 1) _
                & "','" & cboParentesco.ItemData(cboParentesco.ListIndex) & "','" & txtTelefono1.Text _
                & "','" & txtTelefono2.Text & "','" & txtTelMovil.Text & "','" & txtTelFax.Text _
                & "','" & txtEmail.Text & "','" & txtEmail_02.Text & "','" & txtWebSite.Text _
                & "','" & cboProvincia.ItemData(cboProvincia.ListIndex) _
                & "','" & cboCanton.ItemData(cboCanton.ListIndex) _
                & "','" & cboDistrito.ItemData(cboDistrito.ListIndex) & "','" & txtNotas.Text _
                & "','" & txtDireccion.Text & "','" & txtApartado.Text _
                & "','E','" & glogon.Usuario & "'"
  
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Familiar Id:  " & txtCedula.Text & " Consec.: " & vCodigo)
  
  Call sbLimpiaPantalla(0)
  Call sbToolBar(tlbPrincipal, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtApartado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApellido1.SetFocus
End Sub

Private Sub txtApellido1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApellido2.SetFocus
End Sub

Private Sub txtApellido2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboParentesco.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  
  gBusquedas.Col1Name = "Registro Id"
  gBusquedas.Col2Name = "Identificacion"
  gBusquedas.Col3Name = "Nombre Completo"
  
  gBusquedas.Consulta = "select Cod_Familiar,Identificacion,Nombre from RH_PERSONA_FAMILIARES"
  gBusquedas.Filtro = " and EmpleadoId = '" & vEmpleadoId & "'"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub cboParentesco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFechaNacimiento.SetFocus
End Sub

Private Sub dtpFechaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPorcentaje.SetFocus
End Sub

Private Sub txtTelefono1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelFax.SetFocus
End Sub

Private Sub txtApartadoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail_02.SetFocus
End Sub

Private Sub txtEmail_02_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtWebSite.SetFocus
End Sub


Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedula.SetFocus
End Sub


Private Sub txtTelFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub

Private Sub txtTelMovil_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono1.SetFocus
End Sub

Private Sub txtWebSite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApartado.SetFocus

End Sub
