VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmRadar_Directorio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Radar: Directorio"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   13056
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   13056
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   12360
      Top             =   120
   End
   Begin XtremeSuiteControls.ComboBox cboCentro 
      Height          =   312
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5892
      _Version        =   1245185
      _ExtentX        =   10393
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7572
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   12972
      _Version        =   1245185
      _ExtentX        =   22881
      _ExtentY        =   13356
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
      Color           =   2048
      ItemCount       =   3
      Item(0).Caption =   "General"
      Item(0).ControlCount=   31
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "cboTipo"
      Item(0).Control(2)=   "Label1(13)"
      Item(0).Control(3)=   "cboProvincia"
      Item(0).Control(4)=   "cboCanton"
      Item(0).Control(5)=   "cboDistrito"
      Item(0).Control(6)=   "Label1(4)"
      Item(0).Control(7)=   "Label1(5)"
      Item(0).Control(8)=   "Label1(6)"
      Item(0).Control(9)=   "txtTelefono1"
      Item(0).Control(10)=   "Label1(0)"
      Item(0).Control(11)=   "txtTelefono1Ext"
      Item(0).Control(12)=   "txtTelefono2"
      Item(0).Control(13)=   "Label1(1)"
      Item(0).Control(14)=   "txtTelefono2Ext"
      Item(0).Control(15)=   "txtEmail1"
      Item(0).Control(16)=   "Label1(2)"
      Item(0).Control(17)=   "txtEmail2"
      Item(0).Control(18)=   "Label1(3)"
      Item(0).Control(19)=   "txtWebSite"
      Item(0).Control(20)=   "Label1(7)"
      Item(0).Control(21)=   "txtFacebook"
      Item(0).Control(22)=   "Label1(8)"
      Item(0).Control(23)=   "txtDireccion"
      Item(0).Control(24)=   "GroupBox1(0)"
      Item(0).Control(25)=   "GroupBox1(1)"
      Item(0).Control(26)=   "txtContactoVisitas"
      Item(0).Control(27)=   "Label1(17)"
      Item(0).Control(28)=   "lswSec"
      Item(0).Control(29)=   "txtCriterioDept"
      Item(0).Control(30)=   "txtCriterioSec"
      Item(1).Caption =   "Contactos"
      Item(1).ControlCount=   21
      Item(1).Control(0)=   "Label1(18)"
      Item(1).Control(1)=   "Label1(19)"
      Item(1).Control(2)=   "Label1(20)"
      Item(1).Control(3)=   "Label1(21)"
      Item(1).Control(4)=   "Label1(22)"
      Item(1).Control(5)=   "Label1(23)"
      Item(1).Control(6)=   "lswContactos"
      Item(1).Control(7)=   "cboContacto"
      Item(1).Control(8)=   "txtC_Telefono1"
      Item(1).Control(9)=   "txtC_Telefono2"
      Item(1).Control(10)=   "txtC_Email1"
      Item(1).Control(11)=   "txtC_Email2"
      Item(1).Control(12)=   "Label1(24)"
      Item(1).Control(13)=   "chkContacto"
      Item(1).Control(14)=   "Label1(25)"
      Item(1).Control(15)=   "txtContactoId"
      Item(1).Control(16)=   "btnContacto_Nuevo"
      Item(1).Control(17)=   "txtC_Nombre"
      Item(1).Control(18)=   "txtC_Identificacion"
      Item(1).Control(19)=   "txtC_Movil"
      Item(1).Control(20)=   "Label1(26)"
      Item(2).Caption =   "Tipos de Centros"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "lswCentrosTipos"
      Item(2).Control(1)=   "Label1(27)"
      Begin XtremeSuiteControls.PushButton btnContacto_Nuevo 
         Height          =   492
         Left            =   -58600
         TabIndex        =   64
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1245185
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Nuevo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmRadar_Directorio.frx":0000
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1932
         Index           =   0
         Left            =   9720
         TabIndex        =   27
         Top             =   5400
         Width           =   3132
         _Version        =   1245185
         _ExtentX        =   5524
         _ExtentY        =   3408
         _StockProps     =   79
         Caption         =   "Localización:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtLatitud 
            Height          =   312
            Left            =   1560
            TabIndex        =   36
            Top             =   360
            Width           =   1332
            _Version        =   1245185
            _ExtentX        =   2350
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
         Begin XtremeSuiteControls.FlatEdit txtLongitud 
            Height          =   312
            Left            =   1560
            TabIndex        =   38
            Top             =   720
            Width           =   1332
            _Version        =   1245185
            _ExtentX        =   2350
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
         Begin XtremeSuiteControls.FlatEdit txtLatitudM 
            Height          =   312
            Left            =   1560
            TabIndex        =   40
            Top             =   1200
            Width           =   1332
            _Version        =   1245185
            _ExtentX        =   2350
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
         Begin XtremeSuiteControls.FlatEdit txtLongitudM 
            Height          =   312
            Left            =   1560
            TabIndex        =   42
            Top             =   1560
            Width           =   1332
            _Version        =   1245185
            _ExtentX        =   2350
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Longitud (m)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   16
            Left            =   600
            TabIndex        =   43
            Top             =   1560
            Width           =   972
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Latitud (m)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   15
            Left            =   600
            TabIndex        =   41
            Top             =   1200
            Width           =   972
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Longitud"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   14
            Left            =   600
            TabIndex        =   39
            Top             =   720
            Width           =   972
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Latitud"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   12
            Left            =   600
            TabIndex        =   37
            Top             =   360
            Width           =   972
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono1 
         Height          =   312
         Left            =   7920
         TabIndex        =   12
         Top             =   1320
         Width           =   2052
         _Version        =   1245185
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3012
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   5892
         _Version        =   1245185
         _ExtentX        =   10393
         _ExtentY        =   5313
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
         MultiSelect     =   -1  'True
         View            =   3
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   7920
         TabIndex        =   4
         Top             =   480
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboProvincia 
         Height          =   312
         Left            =   6240
         TabIndex        =   6
         Top             =   3840
         Width           =   2292
         _Version        =   1245185
         _ExtentX        =   4043
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboCanton 
         Height          =   312
         Left            =   8520
         TabIndex        =   7
         Top             =   3840
         Width           =   2172
         _Version        =   1245185
         _ExtentX        =   3831
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboDistrito 
         Height          =   312
         Left            =   10680
         TabIndex        =   8
         Top             =   3840
         Width           =   2172
         _Version        =   1245185
         _ExtentX        =   3831
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono1Ext 
         Height          =   312
         Left            =   10080
         TabIndex        =   14
         Top             =   1320
         Width           =   732
         _Version        =   1245185
         _ExtentX        =   1291
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
         Left            =   7920
         TabIndex        =   15
         Top             =   1680
         Width           =   2052
         _Version        =   1245185
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono2Ext 
         Height          =   312
         Left            =   10080
         TabIndex        =   17
         Top             =   1680
         Width           =   732
         _Version        =   1245185
         _ExtentX        =   1291
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
      Begin XtremeSuiteControls.FlatEdit txtEmail1 
         Height          =   312
         Left            =   7920
         TabIndex        =   18
         Top             =   2040
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtEmail2 
         Height          =   312
         Left            =   7920
         TabIndex        =   20
         Top             =   2400
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtWebSite 
         Height          =   312
         Left            =   7920
         TabIndex        =   22
         Top             =   2760
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtFacebook 
         Height          =   312
         Left            =   7920
         TabIndex        =   24
         Top             =   3120
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
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
         Height          =   1032
         Left            =   6240
         TabIndex        =   26
         Top             =   4200
         Width           =   6612
         _Version        =   1245185
         _ExtentX        =   11663
         _ExtentY        =   1820
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1932
         Index           =   1
         Left            =   6240
         TabIndex        =   28
         Top             =   5400
         Width           =   3132
         _Version        =   1245185
         _ExtentX        =   5524
         _ExtentY        =   3408
         _StockProps     =   79
         Caption         =   "Visitas y otras referencias: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkVisitaAutorizacion 
            Height          =   252
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   2892
            _Version        =   1245185
            _ExtentX        =   5101
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Visita Requiere Autorización?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtPoblado 
            Height          =   312
            Left            =   1440
            TabIndex        =   29
            Top             =   840
            Width           =   1572
            _Version        =   1245185
            _ExtentX        =   2773
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
         Begin XtremeSuiteControls.FlatEdit txtZona 
            Height          =   312
            Left            =   1440
            TabIndex        =   31
            Top             =   1200
            Width           =   1572
            _Version        =   1245185
            _ExtentX        =   2773
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
         Begin XtremeSuiteControls.FlatEdit txtEmpleadosNo 
            Height          =   312
            Left            =   1440
            TabIndex        =   34
            Top             =   1560
            Width           =   1572
            _Version        =   1245185
            _ExtentX        =   2773
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Empleados"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   11
            Left            =   240
            TabIndex        =   35
            Top             =   1560
            Width           =   972
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Zona"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   10
            Left            =   240
            TabIndex        =   32
            Top             =   1200
            Width           =   972
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Poblado"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   9
            Left            =   240
            TabIndex        =   30
            Top             =   840
            Width           =   972
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtContactoVisitas 
         Height          =   312
         Left            =   7920
         TabIndex        =   44
         Top             =   840
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.ListView lswContactos 
         Height          =   7212
         Left            =   -69880
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   5892
         _Version        =   1245185
         _ExtentX        =   10393
         _ExtentY        =   12721
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
         MultiSelect     =   -1  'True
         View            =   3
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboContacto 
         Height          =   312
         Left            =   -62080
         TabIndex        =   47
         Top             =   1680
         Visible         =   0   'False
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Telefono1 
         Height          =   312
         Left            =   -62080
         TabIndex        =   49
         Top             =   4080
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1245185
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Telefono2 
         Height          =   312
         Left            =   -62080
         TabIndex        =   51
         Top             =   4440
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1245185
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Email1 
         Height          =   312
         Left            =   -62080
         TabIndex        =   53
         Top             =   4800
         Visible         =   0   'False
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtC_Email2 
         Height          =   312
         Left            =   -62080
         TabIndex        =   55
         Top             =   5160
         Visible         =   0   'False
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtC_Nombre 
         Height          =   312
         Left            =   -62080
         TabIndex        =   57
         Top             =   2520
         Visible         =   0   'False
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtC_Identificacion 
         Height          =   312
         Left            =   -62080
         TabIndex        =   59
         Top             =   2880
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1245185
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkContacto 
         Height          =   252
         Left            =   -59200
         TabIndex        =   61
         Top             =   2880
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1245185
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Contacto Activo?"
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
      Begin XtremeSuiteControls.FlatEdit txtContactoId 
         Height          =   312
         Left            =   -58600
         TabIndex        =   63
         Top             =   1200
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1245185
         _ExtentX        =   2561
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
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Movil 
         Height          =   312
         Left            =   -62080
         TabIndex        =   66
         Top             =   3480
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1245185
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswSec 
         Height          =   3372
         Left            =   120
         TabIndex        =   69
         Top             =   4200
         Width           =   5892
         _Version        =   1245185
         _ExtentX        =   10393
         _ExtentY        =   5948
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
         MultiSelect     =   -1  'True
         View            =   3
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCriterioDept 
         Height          =   312
         Left            =   120
         TabIndex        =   70
         Top             =   360
         Width           =   5892
         _Version        =   1245185
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
      Begin XtremeSuiteControls.FlatEdit txtCriterioSec 
         Height          =   312
         Left            =   120
         TabIndex        =   71
         Top             =   3840
         Width           =   5892
         _Version        =   1245185
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
      Begin XtremeSuiteControls.ListView lswCentrosTipos 
         Height          =   7212
         Left            =   -64480
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   7452
         _Version        =   1245185
         _ExtentX        =   13144
         _ExtentY        =   12721
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
         MultiSelect     =   -1  'True
         View            =   3
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipos de Centros de Trabajo asociados a esta institución"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1092
         Index           =   27
         Left            =   -67840
         TabIndex        =   73
         Top             =   480
         Visible         =   0   'False
         Width           =   3372
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Móvil"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   26
         Left            =   -63760
         TabIndex        =   67
         Top             =   3480
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto Id:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   25
         Left            =   -59680
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   24
         Left            =   -63760
         TabIndex        =   60
         Top             =   2880
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Completo: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   23
         Left            =   -63760
         TabIndex        =   58
         Top             =   2520
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email (2)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   22
         Left            =   -63760
         TabIndex        =   56
         Top             =   5160
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email (1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   21
         Left            =   -63760
         TabIndex        =   54
         Top             =   4800
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono (2)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   20
         Left            =   -63760
         TabIndex        =   52
         Top             =   4440
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono (1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   19
         Left            =   -63760
         TabIndex        =   50
         Top             =   4080
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Contacto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   18
         Left            =   -63760
         TabIndex        =   48
         Top             =   1680
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto Visitas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   17
         Left            =   6240
         TabIndex        =   45
         Top             =   840
         Width           =   1572
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Facebook"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   8
         Left            =   6240
         TabIndex        =   25
         Top             =   3120
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sitio Web"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   7
         Left            =   6240
         TabIndex        =   23
         Top             =   2760
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email (2)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   6240
         TabIndex        =   21
         Top             =   2400
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email (1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   6240
         TabIndex        =   19
         Top             =   2040
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono (2)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   6240
         TabIndex        =   16
         Top             =   1680
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono (1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   6240
         TabIndex        =   13
         Top             =   1320
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   6
         Left            =   10680
         TabIndex        =   11
         Top             =   3600
         Width           =   852
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantón"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   8520
         TabIndex        =   10
         Top             =   3600
         Width           =   852
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   6240
         TabIndex        =   9
         Top             =   3600
         Width           =   852
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Centro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   13
         Left            =   6240
         TabIndex        =   5
         Top             =   480
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   492
      Left            =   11400
      TabIndex        =   65
      Top             =   480
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Guardar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmRadar_Directorio.frx":07B9
   End
   Begin VB.Label lblSeccion 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Height          =   252
      Left            =   6120
      TabIndex        =   68
      Top             =   720
      Width           =   5772
   End
   Begin VB.Label lblDepartamento 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5772
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmRadar_Directorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnContacto_Nuevo_Click()

txtContactoId.Text = "0"
txtC_Email1.Text = ""
txtC_Email2.Text = ""
txtC_Movil.Text = ""
txtC_Telefono1.Text = ""
txtC_Telefono2.Text = ""
txtC_Identificacion.Text = ""
txtC_Nombre.Text = ""

chkContacto.Value = xtpChecked

txtC_Nombre.SetFocus

End Sub

Private Sub btnGuardar_Click()
Dim strSQL As String

On Error GoTo vError

If tcMain.Selected.Index = 0 Then
    strSQL = "exec spRadar_Dir_Update " & cboCentro.ItemData(cboCentro.ListIndex) & ",'" & lblDepartamento.Tag & "','" & lblSeccion.Tag _
           & "','" & glogon.Usuario & "','" & cboTipo.ItemData(cboTipo.ListIndex) _
           & "','" & txtTelefono1.Text & "','" & txtTelefono1Ext.Text & "','" & txtTelefono2.Text & "','" & txtTelefono2Ext.Text _
           & "','" & txtEmail1.Text & "','" & txtEmail2.Text & "','" & txtWebSite.Text & "','" & txtFacebook.Text _
           & "','" & txtContactoVisitas.Text & "'," & chkVisitaAutorizacion.Value _
           & ",'" & cboProvincia.ItemData(cboProvincia.ListIndex) & "','" & cboCanton.ItemData(cboCanton.ListIndex) _
           & "','" & cboDistrito.ItemData(cboDistrito.ListIndex) & "','" & txtDireccion.Text _
           & "','" & txtPoblado.Text & "','" & txtZona.Text & "'," & txtEmpleadosNo.Text _
           & "," & txtLatitud.Text & "," & txtLatitudM.Text & "," & txtLongitud.Text & "," & txtLongitudM.Text
    Call ConectionExecute(strSQL)
   
Else
            
    strSQL = "exec spRadar_Dir_Contacto_Update " & cboCentro.ItemData(cboCentro.ListIndex) & ",'" & lblDepartamento.Tag & "','" & lblSeccion.Tag _
           & "','" & glogon.Usuario & "','" & cboContacto.ItemData(cboContacto.ListIndex) _
           & "','" & txtC_Identificacion & "','" & txtC_Nombre & "'," & chkContacto.Value _
           & ",'" & txtC_Telefono1.Text & "','" & txtC_Telefono2.Text & "','" & txtC_Movil.Text _
           & "','" & txtEmail1.Text & "','" & txtEmail2.Text & "','I'," & txtContactoId.Text
    Call ConectionExecute(strSQL)
    
    'Cargar Contactos
    Call sbContactos_Load
    
End If

MsgBox "Información Registrada Satisfactoriamente!", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboCanton_Click()
Dim strSQL As String

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


Private Sub sbLimpia()

txtContactoVisitas.Text = ""
txtTelefono1.Text = ""
txtTelefono1Ext.Text = ""
txtTelefono2.Text = ""
txtTelefono2Ext.Text = ""

txtEmail1.Text = ""
txtEmail2.Text = ""

txtWebSite.Text = ""
txtFacebook.Text = ""

txtDireccion.Text = ""

chkVisitaAutorizacion.Value = xtpChecked

txtZona.Text = ""
txtPoblado.Text = ""
txtEmpleadosNo.Text = 1

txtLatitud.Text = 0
txtLatitudM.Text = 0
txtLongitud.Text = 0
txtLongitudM.Text = 0


End Sub

Private Sub cboCentro_Click()
Dim strSQL As String

If vPaso Then Exit Sub

tcMain.Item(0).Selected = True

lblDepartamento.Tag = "-Main-"
lblDepartamento.Caption = "[Principal]"

lblSeccion.Tag = "-Main-"
lblSeccion.Caption = "[Principal]"

strSQL = "select rtrim(Ct.CENTRO_TIPO) as 'IdX', rtrim(Ct.Descripcion) as 'ItmX'" _
       & " from RADAR_CENTROS_TIPOS Ct inner join RADAR_INSTITUCION_TIPOS It on Ct.CENTRO_TIPO = It.CENTRO_TIPO" _
       & " and It.COD_INSTITUCION = " & cboCentro.ItemData(cboCentro.ListIndex) _
       & " where Ct.activo = 1"
Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

Call sbCentroConsulta

'Cargado de Departamentos
txtCriterioDept.Text = ""
Call sbDepartamentosCarga

End Sub

Private Sub cboProvincia_Click()
Dim strSQL As String

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click
End Sub

Private Sub Form_Activate()
vModulo = 37
End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

vModulo = 37

lsw.ColumnHeaders.Add , , "Código", 1200
lsw.ColumnHeaders.Add , , "Centro de Trabajo", 4200

lswSec.ColumnHeaders.Add , , "Código", 1200
lswSec.ColumnHeaders.Add , , "Centro de Trabajo", 4200

lswContactos.ColumnHeaders.Add , , "Contacto Id", 1200
lswContactos.ColumnHeaders.Add , , "Nombre", 4200

lswCentrosTipos.ColumnHeaders.Add , , "Tipo", 1200
lswCentrosTipos.ColumnHeaders.Add , , "Descripción", 4200


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbSeccionesCarga()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

lswSec.ListItems.Clear

strSQL = "select S.cod_Seccion, S.descripcion, isnull(D.cod_Institucion,0) as 'Registro' " _
       & " from afSecciones S left join RADAR_DIRECTORIO D on S.cod_Institucion = D.cod_Institucion" _
       & " and S.cod_Departamento = D.cod_Departamento and S.cod_Seccion = D.cod_Seccion" _
       & " where S.cod_institucion = " & cboCentro.ItemData(cboCentro.ListIndex) _
       & "   and S.cod_Departamento = '" & lblDepartamento.Tag _
       & "'  and S.descripcion like '%" & Trim(txtCriterioSec.Text) & "%'"
Call OpenRecordSet(rs, strSQL)

lswSec.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lswSec.ListItems.Add(, , rs!cod_Seccion)
      itmX.SubItems(1) = rs!Descripcion
      If rs!Registro > 0 Then
        itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbDepartamentosCarga()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

lblSeccion.Tag = "-Main-"
lblSeccion.Caption = "[Principal]"

lsw.ListItems.Clear
lswSec.ListItems.Clear

strSQL = "select S.cod_Departamento, S.descripcion, isnull(D.cod_Institucion,0) as 'Registro' " _
       & " from afDepartamentos S left join RADAR_DIRECTORIO D on S.cod_Institucion = D.cod_Institucion" _
       & " and S.cod_Departamento = D.cod_Departamento" _
       & " where S.cod_institucion = " & cboCentro.ItemData(cboCentro.ListIndex) _
       & "  and S.descripcion like '%" & Trim(txtCriterioDept.Text) & "%'" _
       & " group by S.cod_Departamento, S.descripcion, isnull(D.cod_Institucion,0)"
       
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!cod_Departamento)
      itmX.SubItems(1) = rs!Descripcion
      If rs!Registro > 0 Then
        itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
tcMain.Item(0).Selected = True

lblDepartamento.Tag = Item.Text
lblDepartamento.Caption = Item.SubItems(1)

lblSeccion.Tag = "-Main-"
lblSeccion.Caption = "[Principal]"

Call sbCentroConsulta

'Carga Secciones
txtCriterioSec.Text = ""
Call sbSeccionesCarga

End Sub


Private Sub sbCentroConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


Call sbLimpia

strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & ",rtrim(B.Descripcion) as 'Tipo'" _
       & " from RADAR_DIRECTORIO P inner Join RADAR_CENTROS_TIPOS B on P.CENTRO_TIPO = B.CENTRO_TIPO" _
       & " left join Provincias Prov on P.Cod_Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Cod_Provincia = Cant.Provincia and P.Cod_Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Cod_Provincia = Dist.Provincia and P.Cod_Canton = Dist.Canton and P.Cod_distrito = Dist.distrito" _
       & " where P.COD_DEPARTAMENTO = '" & lblDepartamento.Tag & "' and P.COD_INSTITUCION = " & cboCentro.ItemData(cboCentro.ListIndex) _
       & " and P.COD_SECCION = '" & lblSeccion.Tag & "'"
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 
  txtTelefono1.Text = rs!Telefono_01 & ""
  txtTelefono2.Text = rs!Telefono_02 & ""
  txtTelefono1Ext.Text = rs!Telefono_01_Ext & ""
  txtTelefono2Ext.Text = rs!Telefono_02_Ext & ""

  txtEmail1.Text = rs!Email_01 & ""
  txtEmail2.Text = rs!Email_02 & ""
  
  txtFacebook.Text = rs!Facebook & ""
  txtWebSite.Text = rs!Sitio_Web & ""

  txtContactoVisitas.Text = rs!Visita_Contacto & ""
  chkVisitaAutorizacion.Value = rs!visita_autorizada
  
  
  txtPoblado.Text = rs!Poblado & ""
  txtZona.Text = rs!Zona & ""
  
  txtEmpleadosNo.Text = CStr(rs!Empleados_Numero)
  
  txtLatitud.Text = CStr(rs!Mapa_Latitud)
  txtLatitudM.Text = CStr(rs!Mapa_Latitud_M)
  txtLongitud.Text = CStr(rs!Mapa_Longitud)
  txtLongitudM.Text = CStr(rs!Mapa_Longitud_M)

  Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
  Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
  Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
  
  Call sbCboAsignaDato(cboTipo, rs!Tipo, True, rs!Centro_Tipo)
     
  cboDistrito.ToolTipText = Trim(rs!cod_distrito) & ""
  txtDireccion.Text = rs!Direccion


End If
rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswCentrosTipos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

If Item.Checked Then
   strSQL = "insert RADAR_INSTITUCION_TIPOS(COD_INSTITUCION,CENTRO_TIPO,REGISTRO_FECHA,REGISTRO_USUARIO)" _
          & " values(" & cboCentro.ItemData(cboCentro.ListIndex) & ",'" & Item.Text & "',dbo.mygetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete RADAR_INSTITUCION_TIPOS where COD_INSTITUCION = " & cboCentro.ItemData(cboCentro.ListIndex) _
          & " and CENTRO_TIPO = '" & Item.Text & "'"
End If

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswContactos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

Call btnContacto_Nuevo_Click


txtContactoId.Text = Item.Text

strSQL = "select P.*" _
       & ",rtrim(B.Descripcion) as 'Tipo'" _
       & " from RADAR_DIRECTORIO_CONTACTOS P inner Join RADAR_CONTACTOS_TIPOS B on P.CONTACTO_TIPO = B.CONTACTO_TIPO" _
       & " where P.COD_INSTITUCION = " & cboCentro.ItemData(cboCentro.ListIndex) _
       & ",P.COD_DEPARTAMENTO = '" & lblDepartamento.Tag & "' and P.COD_SECCION = '" & lblSeccion.Tag & "'" _
       & " and P.CONTACTO_LINEA = " & Item.Text
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  
  txtContactoId.Text = rs!Contacto_Linea
  
  txtC_Identificacion = rs!Identificacion & ""
  txtC_Nombre = rs!Nombre & ""
  
  chkContacto.Value = rs!Activo
  txtC_Movil.Text = rs!Tel_Movil & ""
  txtC_Telefono1.Text = rs!Telefono_01 & ""
  txtC_Telefono2.Text = rs!Telefono_02 & ""

  txtC_Email1.Text = rs!Email_01 & ""
  txtC_Email2.Text = rs!Email_02 & ""
  
  Call sbCboAsignaDato(cboContacto, rs!Tipo, True, rs!Contacto_Tipo)
     

End If
rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbContactos_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select Contacto_Linea, Nombre" _
       & " from RADAR_DIRECTORIO_CONTACTOS" _
       & " where cod_institucion = " & cboCentro.ItemData(cboCentro.ListIndex) _
       & " and cod_Departamento = '" & lblDepartamento.Tag & "'"
Call OpenRecordSet(rs, strSQL)

lswContactos.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lswContactos.ListItems.Add(, , rs!Contacto_Linea)
      itmX.SubItems(1) = rs!Nombre
  rs.MoveNext
Loop
rs.Close

'Limpia
Call btnContacto_Nuevo_Click

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCentrosTipos_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select rtrim(Ct.CENTRO_TIPO) as 'IdX', rtrim(Ct.Descripcion) as 'ItmX'" _
       & ", case when isnull(It.COD_INSTITUCION,0) = 0 THEN 0 ELSE 1 END AS 'Check'" _
       & " from RADAR_CENTROS_TIPOS Ct left join RADAR_INSTITUCION_TIPOS It on Ct.CENTRO_TIPO = It.CENTRO_TIPO" _
       & " and It.COD_INSTITUCION = " & cboCentro.ItemData(cboCentro.ListIndex) _
       & " where Ct.activo = 1"

Call OpenRecordSet(rs, strSQL)

vPaso = True
lswCentrosTipos.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lswCentrosTipos.ListItems.Add(, , rs!IdX)
      itmX.SubItems(1) = rs!itmX
      If rs!Check = 1 Then
          itmX.Checked = True
      End If
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswSec_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
tcMain.Item(0).Selected = True


lblSeccion.Tag = Item.Text
lblSeccion.Caption = Item.SubItems(1)

Call sbCentroConsulta
End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If vPaso Then Exit Sub

Select Case Item.Caption
  Case "General"
  Case "Tipos de Centros"
        Call sbCentrosTipos_Load
  Case "Contactos"
        Call sbContactos_Load
End Select


End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

On Error GoTo vError

TimerX.Interval = 0
TimerX.Enabled = False

strSQL = "select rtrim(CONTACTO_TIPO) as 'IdX', rtrim(Descripcion) as 'ItmX' from RADAR_CONTACTOS_TIPOS where activo = 1"
Call sbCbo_Llena_New(cboContacto, strSQL, False, True)

vPaso = True
    strSQL = "select rtrim(cod_institucion) as 'IdX', rtrim(Descripcion) as 'ItmX' from INSTITUCIONES where ACTIVA = 1"
    Call sbCbo_Llena_New(cboCentro, strSQL, False, True)

    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False

Call cboCentro_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub txtCriterioDept_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbDepartamentosCarga
End Sub

Private Sub txtCriterioSec_KeyUp(KeyCode As Integer, Shift As Integer)
Call sbSeccionesCarga
End Sub
