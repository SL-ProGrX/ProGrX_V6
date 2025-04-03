VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_PolizasAseguradoras 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Aseguradoras"
   ClientHeight    =   7020
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   10935
      _Version        =   1572864
      _ExtentX        =   19283
      _ExtentY        =   10393
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
      Item(0).ControlCount=   21
      Item(0).Control(0)=   "txtCedJuridica"
      Item(0).Control(1)=   "txtTelefono"
      Item(0).Control(2)=   "txtTelefono2"
      Item(0).Control(3)=   "txtTelFax"
      Item(0).Control(4)=   "txtWebSite"
      Item(0).Control(5)=   "txtEmail"
      Item(0).Control(6)=   "txtEmail2"
      Item(0).Control(7)=   "txtAptoPostal"
      Item(0).Control(8)=   "chkActivo"
      Item(0).Control(9)=   "Label4(0)"
      Item(0).Control(10)=   "Label4(1)"
      Item(0).Control(11)=   "Label4(2)"
      Item(0).Control(12)=   "Label4(3)"
      Item(0).Control(13)=   "Label4(4)"
      Item(0).Control(14)=   "Label4(5)"
      Item(0).Control(15)=   "Label4(6)"
      Item(0).Control(16)=   "Label4(7)"
      Item(0).Control(17)=   "gbDireccion"
      Item(0).Control(18)=   "gbContacto"
      Item(0).Control(19)=   "Label4(12)"
      Item(0).Control(20)=   "txtProveedor"
      Item(1).Caption =   "Adicionales"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "gbCuentasBancarias"
      Item(1).Control(1)=   "gbRecaudo"
      Item(1).Control(2)=   "gbCuentasContabbles"
      Begin XtremeSuiteControls.FlatEdit txtCedJuridica 
         Height          =   312
         Left            =   1680
         TabIndex        =   5
         Top             =   960
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono 
         Height          =   312
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono2 
         Height          =   312
         Left            =   1680
         TabIndex        =   7
         Top             =   1680
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelFax 
         Height          =   312
         Left            =   1680
         TabIndex        =   8
         Top             =   2040
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtWebSite 
         Height          =   312
         Left            =   5520
         TabIndex        =   9
         Top             =   960
         Width           =   5292
         _Version        =   1572864
         _ExtentX        =   9334
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   312
         Left            =   5520
         TabIndex        =   10
         Top             =   1320
         Width           =   5292
         _Version        =   1572864
         _ExtentX        =   9334
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail2 
         Height          =   312
         Left            =   5520
         TabIndex        =   11
         Top             =   1680
         Width           =   5292
         _Version        =   1572864
         _ExtentX        =   9334
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAptoPostal 
         Height          =   312
         Left            =   5520
         TabIndex        =   12
         Top             =   2040
         Width           =   5292
         _Version        =   1572864
         _ExtentX        =   9334
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   252
         Left            =   2760
         TabIndex        =   13
         Top             =   480
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activa?"
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
      Begin XtremeSuiteControls.GroupBox gbDireccion 
         Height          =   2052
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   10692
         _Version        =   1572864
         _ExtentX        =   18860
         _ExtentY        =   3619
         _StockProps     =   79
         Caption         =   "Dirección"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   312
            Left            =   1560
            TabIndex        =   23
            Top             =   480
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3836
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
            Left            =   1560
            TabIndex        =   24
            Top             =   840
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3836
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
            Left            =   1560
            TabIndex        =   25
            Top             =   1200
            Width           =   2172
            _Version        =   1572864
            _ExtentX        =   3836
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
            TabIndex        =   26
            Top             =   480
            Width           =   6852
            _Version        =   1572864
            _ExtentX        =   12086
            _ExtentY        =   1926
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   10
            Left            =   120
            TabIndex        =   49
            Top             =   1200
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Distrito"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   9
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cantón"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   8
            Left            =   120
            TabIndex        =   47
            Top             =   480
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Provincia"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox gbContacto 
         Height          =   972
         Left            =   120
         TabIndex        =   27
         Top             =   4560
         Width           =   10572
         _Version        =   1572864
         _ExtentX        =   18648
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   "Contacto:"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtContacto 
            Height          =   312
            Left            =   1560
            TabIndex        =   28
            Top             =   480
            Width           =   9012
            _Version        =   1572864
            _ExtentX        =   15896
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   11
            Left            =   120
            TabIndex        =   50
            Top             =   480
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Nombre"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox gbCuentasBancarias 
         Height          =   2412
         Left            =   -69640
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   4254
         _StockProps     =   79
         Caption         =   "Cuentas Bancarias"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswCuentas 
            Height          =   1452
            Left            =   120
            TabIndex        =   30
            Top             =   756
            Width           =   10092
            _Version        =   1572864
            _ExtentX        =   17801
            _ExtentY        =   2561
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
            Appearance      =   16
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCuentas 
            Height          =   372
            Left            =   8520
            TabIndex        =   31
            Top             =   360
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cuentas Bancarias"
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
         End
         Begin XtremeSuiteControls.ComboBox cboBancos 
            Height          =   312
            Left            =   2040
            TabIndex        =   32
            Top             =   396
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
            Left            =   120
            TabIndex        =   33
            Top             =   396
            Width           =   1812
         End
      End
      Begin XtremeSuiteControls.GroupBox gbRecaudo 
         Height          =   1212
         Left            =   -69640
         TabIndex        =   34
         Top             =   2880
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   2138
         _StockProps     =   79
         Caption         =   "Recaudación y Formatos"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboFormato 
            Height          =   312
            Left            =   1920
            TabIndex        =   35
            Top             =   840
            Width           =   2292
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtRetDesc 
            Height          =   312
            Left            =   3120
            TabIndex        =   36
            Top             =   480
            Width           =   7092
            _Version        =   1572864
            _ExtentX        =   12509
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
         Begin XtremeSuiteControls.FlatEdit txtRetCod 
            Height          =   312
            Left            =   1920
            TabIndex        =   37
            Top             =   480
            Width           =   1212
            _Version        =   1572864
            _ExtentX        =   2138
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
            Height          =   252
            Index           =   13
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   1092
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
            Height          =   252
            Index           =   18
            Left            =   120
            TabIndex        =   38
            Top             =   840
            Width           =   1212
         End
      End
      Begin XtremeSuiteControls.GroupBox gbCuentasContabbles 
         Height          =   1212
         Left            =   -69640
         TabIndex        =   40
         Top             =   4320
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1572864
         _ExtentX        =   18224
         _ExtentY        =   2138
         _StockProps     =   79
         Caption         =   "Cuentas Contables:"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   312
            Left            =   4320
            TabIndex        =   41
            Top             =   480
            Width           =   5892
            _Version        =   1572864
            _ExtentX        =   10393
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
         Begin XtremeSuiteControls.FlatEdit txtComisionCtaDesc 
            Height          =   312
            Left            =   4320
            TabIndex        =   42
            Top             =   840
            Width           =   5892
            _Version        =   1572864
            _ExtentX        =   10393
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaCod 
            Height          =   312
            Left            =   1920
            TabIndex        =   43
            Top             =   480
            Width           =   2412
            _Version        =   1572864
            _ExtentX        =   4254
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
         Begin XtremeSuiteControls.FlatEdit txtComisionCta 
            Height          =   312
            Left            =   1920
            TabIndex        =   44
            Top             =   840
            Width           =   2412
            _Version        =   1572864
            _ExtentX        =   4254
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
         Begin VB.Label Label1 
            Caption         =   "Cuenta por Pagar"
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
            Left            =   120
            TabIndex        =   46
            Top             =   480
            Width           =   1572
         End
         Begin VB.Label Label1 
            Caption         =   "Comisiones"
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
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1572
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtProveedor 
         Height          =   315
         Left            =   5520
         TabIndex        =   53
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   480
         Width           =   5295
         _Version        =   1572864
         _ExtentX        =   9334
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   495
         Index           =   12
         Left            =   3960
         TabIndex        =   52
         Top             =   360
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Proveedor Relacionado"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   7
         Left            =   3960
         TabIndex        =   21
         Top             =   2040
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Apto. Postal"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   6
         Left            =   3960
         TabIndex        =   20
         Top             =   1680
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Email (2)"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   5
         Left            =   3960
         TabIndex        =   19
         Top             =   1320
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Email (1)"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   4
         Left            =   3960
         TabIndex        =   18
         Top             =   960
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Web Site"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tel. Fax"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Teléfono (2)"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Teléfono (1)"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ced. Jurídica"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8760
      Top             =   360
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   6735
      _Version        =   1572864
      _ExtentX        =   11874
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   9840
      TabIndex        =   51
      Top             =   600
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Aseguradora"
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
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmCR_PolizasAseguradoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vCantonMascara As String, vDistritoMascara As String, vFechaActual As Date
Dim vScroll As Boolean, vPaso As Boolean



Private Sub sbCuentas_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lswCuentas.ListItems.Clear
    strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
           & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
           & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
           & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
           & " where C.Identificacion = '" & Trim(txtCedJuridica.Text) & "' and C.Modulo = 'Pol'"
    
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
GLOBALES.gTag2 = "Pol"

frmCC_Cuentas_Bancarias.Show vbModal

Call sbCuentas_Load
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

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cod_aseguradora from CRD_POLIZAS_ASEGURADORAS"
    
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
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 11
End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

 vModulo = 11

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 

cboFormato.Clear
cboFormato.AddItem "ISABS"
cboFormato.AddItem "SISEPO"

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

tcMain.Item(0).Selected = True

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

cboFormato.Text = "ISABS"

txtProveedor.Text = ""
txtProveedor.Tag = ""

End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 1 Then
  Call sbCuentas_Load
End If
End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

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
Dim strSQL As String

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
       gBusquedas.Consulta = "select cod_aseguradora,nombre from CRD_POLIZAS_ASEGURADORAS"
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
Dim rs As New ADODB.Recordset, strSQL As String


On Error GoTo vError

If Not fxSIFValidaCadena(txtCodigo.Text) Then
   Exit Sub
End If

Me.MousePointer = vbHourglass
'       & ",case when P.Tipo_Emision = 'TE' then 'Transferencia' when P.Tipo_Emision = 'CK' then 'Cheque' else 'Transferencia' end as 'Tipo_Pago' " _

strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & ", isnull(Cat.Descripcion,'') as 'RetDesc', isnull(Ban.Descripcion, '') as 'BancoDesc'" _
       & ", isnull(Cta.Cod_Cuenta_Mask,P.cod_Cuenta) as 'COD_CUENTA_MASK', isnull(Cta.Descripcion,'') as 'CtaDesc'" _
       & ", isnull(Ctc.Cod_Cuenta_Mask,P.cod_Cuenta_Comision) as 'COD_CUENTA_COM', isnull(Ctc.Descripcion,'') as 'CtaComisionDesc'" _
       & ", isnull(Cpp.descripcion,'') as 'ProveedorDesc'" _
       & " from CRD_POLIZAS_ASEGURADORAS P " _
       & " left join Provincias Prov on P.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Provincia = Cant.Provincia and P.Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Provincia = Dist.Provincia and convert(int,P.Canton) = convert(int,Dist.Canton) and P.distrito = Dist.distrito" _
       & " left join Catalogo Cat on P.codigo_Retencion = Cat.Codigo" _
       & " left join Tes_Bancos Ban on P.Cod_Banco = Ban.Id_Banco" _
       & " left join vCNTX_CUENTAS_LOCAL Cta on P.cod_Cuenta = Cta.Cod_Cuenta" _
       & " left join vCNTX_CUENTAS_LOCAL Ctc on P.cod_Cuenta_Comision = Ctc.Cod_Cuenta" _
       & " left join CxP_Proveedores Cpp on P.Cod_Proveedor = Cpp.Cod_Proveedor" _
       & " where P.cod_aseguradora = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCodigo = rs!cod_Aseguradora
  txtCodigo = rs!cod_Aseguradora

  txtNombre = rs!Nombre & ""
  
  txtCedJuridica.Text = rs!Cedula_Juridica & ""
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
     
  cboDistrito.ToolTipText = Trim(rs!Distrito) & ""
  txtDireccion.Text = rs!direccion

  txtProveedor.Tag = rs!Cod_Proveedor & ""
  txtProveedor.Text = rs!ProveedorDesc


'  txtBancoId.Text = rs!Cod_Banco & ""
'  txtBancoDesc.Text = rs!BancoDesc & ""
'  txtCuentaCliente.Text = rs!Cuenta_Cliente & ""
    
  Call sbCboAsignaDato(cboBancos, rs!BancoDesc, True, rs!Cod_Banco & "")

 
  txtRetCod.Text = rs!Codigo_Retencion & ""
  txtRetDesc.Text = rs!RetDesc

  txtCuentaCod.Text = rs!COD_CUENTA_MASK & ""
  txtCuentaDesc.Text = rs!CtaDesc & ""
  
  txtComisionCta.Text = rs!COD_CUENTA_COM & ""
  txtComisionCtaDesc.Text = rs!CtaComisionDesc & ""
  
  
  
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
If Trim(txtComisionCtaDesc.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Cuenta Contable para Comisión no es válida!..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vEstadoCivil As String, vProveedorId As Long

On Error GoTo vError


If IsNumeric(txtProveedor.Tag) Then
  vProveedorId = txtProveedor.Tag
Else
  vProveedorId = 0
End If

If vEdita Then
  strSQL = "update CRD_POLIZAS_ASEGURADORAS set nombre = '" & Trim(txtNombre.Text) & "', Cedula_Juridica = '" & txtCedJuridica.Text _
         & "',Telefono_01 = '" & txtTelefono.Text & "',Telefono_02 = '" & txtTelefono2.Text _
         & "',Tel_Fax = '" & txtTelFax.Text & "',Sitio_Web = '" & txtWebSite.Text & "',apto_postal = '" & txtAptoPostal _
         & "',email_01 = '" & txtEmail & "', email_02 = '" & txtEmail2.Text & "',direccion = '" & txtDireccion _
         & "',distrito = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "',canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
         & "',provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
         & "',Nombre_Contacto = '" & txtContacto.Text & "',Activo = " & chkActivo.Value _
         & ",Codigo_Retencion = '" & txtRetCod.Text & "',formato_Tramas = '" & cboFormato.Text _
         & "', cod_Banco = " & cboBancos.ItemData(cboBancos.ListIndex) _
         & " , Cod_Cuenta = '" & fxgCntCuentaFormato(False, txtCuentaCod.Text, 0) & "'" _
         & " , COD_CUENTA_COMISION = '" & fxgCntCuentaFormato(False, txtComisionCta.Text, 0) & "'" _
         & " , cod_Proveedor = " & vProveedorId _
         & " where cod_aseguradora = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Aseguradora: " & vCodigo)

Else
  vCodigo = txtCodigo

   strSQL = "insert into CRD_POLIZAS_ASEGURADORAS(cod_aseguradora,nombre,cedula_Juridica,Telefono_01,Telefono_02,Tel_fax,Activo,registro_fecha,registro_usuario" _
          & ",apto_postal,email_01,email_02,Sitio_Web,direccion,distrito,provincia,canton,nombre_contacto,cod_banco,cod_Cuenta,COD_CUENTA_COMISION" _
          & " , formato_Tramas, codigo_Retencion, cod_Proveedor)" _
          & " values('" & vCodigo & "','" & txtNombre.Text & "','" & txtCedJuridica.Text & "','" & txtTelefono.Text & "','" & txtTelefono2.Text & "','" & txtTelFax.Text _
          & "'," & chkActivo.Value & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtAptoPostal.Text & "','" & txtEmail.Text & "','" & txtEmail2.Text & "','" & txtWebSite.Text _
          & "','" & txtDireccion.Text & "','" & cboDistrito.ItemData(cboDistrito.ListIndex) & "','" _
          & cboProvincia.ItemData(cboProvincia.ListIndex) & "','" & cboCanton.ItemData(cboCanton.ListIndex) _
          & "','" & txtContacto.Text & "'," & cboBancos.ItemData(cboBancos.ListIndex) _
          & ",'" & fxgCntCuentaFormato(False, txtCuentaCod.Text) _
          & "','" & fxgCntCuentaFormato(False, txtComisionCta.Text) _
          & "','" & cboFormato.Text & "','" & txtRetCod.Text & "', " & vProveedorId & ")"
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
  strSQL = "delete CRD_POLIZAS_ASEGURADORAS where cod_aseguradora = '" & vCodigo & "'"
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
  gBusquedas.Consulta = "select cod_aseguradora,nombre from CRD_POLIZAS_ASEGURADORAS"
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



Private Sub txtCuentaCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRetCod.SetFocus

End Sub


Private Sub txtComisionCta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComisionCtaDesc.SetFocus

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    gBusquedas.Resultado = gCuenta
    txtComisionCtaDesc.Text = fxgCntCuentaDesc(gCuenta)
    txtComisionCta.Text = fxgCntCuentaFormato(True, gCuenta, 0)
End If
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
  gBusquedas.Consulta = "select cod_aseguradora,nombre from CRD_POLIZAS_ASEGURADORAS"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub




Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedJuridica.SetFocus

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
  txtProveedor.Text = gBusquedas.Resultado3
  txtProveedor.Tag = gBusquedas.Resultado
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








