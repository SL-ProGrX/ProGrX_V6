VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmARF_Acreedores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrendamientos: Registro de Arrendadores"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10980
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8760
      Top             =   360
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5892
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   10932
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
      Item(0).Control(0)=   "txtTelefono"
      Item(0).Control(1)=   "txtTelefono2"
      Item(0).Control(2)=   "txtWebSite"
      Item(0).Control(3)=   "txtEmail"
      Item(0).Control(4)=   "txtEmail2"
      Item(0).Control(5)=   "txtAptoPostal"
      Item(0).Control(6)=   "chkActivo"
      Item(0).Control(7)=   "Label4(0)"
      Item(0).Control(8)=   "Label4(1)"
      Item(0).Control(9)=   "Label4(2)"
      Item(0).Control(10)=   "Label4(3)"
      Item(0).Control(11)=   "Label4(4)"
      Item(0).Control(12)=   "Label4(5)"
      Item(0).Control(13)=   "Label4(6)"
      Item(0).Control(14)=   "Label4(7)"
      Item(0).Control(15)=   "gbDireccion"
      Item(0).Control(16)=   "gbContacto"
      Item(0).Control(17)=   "Label4(12)"
      Item(0).Control(18)=   "txtProveedor"
      Item(0).Control(19)=   "cboTipoId"
      Item(0).Control(20)=   "txtIdentificacion"
      Item(1).Caption =   "Adicionales"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "gbCuentasBancarias"
      Item(1).Control(1)=   "gbCuentasContabbles"
      Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   1320
         Width           =   2055
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
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   1680
         Width           =   2055
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
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   2040
         Width           =   2055
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
            TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            Index           =   8
            Left            =   120
            TabIndex        =   16
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   9
            Left            =   120
            TabIndex        =   15
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
            Index           =   10
            Left            =   120
            TabIndex        =   14
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
      End
      Begin XtremeSuiteControls.GroupBox gbContacto 
         Height          =   972
         Left            =   120
         TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   19
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
         Height          =   3735
         Left            =   -69640
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1572864
         _ExtentX        =   18230
         _ExtentY        =   6588
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
            Height          =   2775
            Left            =   120
            TabIndex        =   21
            Top             =   750
            Width           =   10095
            _Version        =   1572864
            _ExtentX        =   17806
            _ExtentY        =   4895
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
            TabIndex        =   22
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
            TabIndex        =   23
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
            TabIndex        =   24
            Top             =   396
            Width           =   1812
         End
      End
      Begin XtremeSuiteControls.GroupBox gbCuentasContabbles 
         Height          =   1212
         Left            =   -69640
         TabIndex        =   25
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
            TabIndex        =   26
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
            TabIndex        =   27
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
            TabIndex        =   28
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
            TabIndex        =   29
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
            TabIndex        =   31
            Top             =   840
            Width           =   1572
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
            TabIndex        =   30
            Top             =   480
            Width           =   1572
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtProveedor 
         Height          =   315
         Left            =   5520
         TabIndex        =   32
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
      Begin XtremeSuiteControls.ComboBox cboTipoId 
         Height          =   330
         Left            =   1680
         TabIndex        =   46
         Top             =   960
         Width           =   2055
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   47
         Top             =   960
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo Id"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   1320
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Identificacion"
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
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   39
         Top             =   1680
         Width           =   1335
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
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   38
         Top             =   2040
         Width           =   1335
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
         Index           =   4
         Left            =   3960
         TabIndex        =   37
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
         Index           =   5
         Left            =   3960
         TabIndex        =   36
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
         Index           =   6
         Left            =   3960
         TabIndex        =   35
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
         Index           =   7
         Left            =   3960
         TabIndex        =   34
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
         Height          =   495
         Index           =   12
         Left            =   3960
         TabIndex        =   33
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
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
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
      TabIndex        =   42
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
      TabIndex        =   43
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
      TabIndex        =   44
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
      TabIndex        =   45
      Top             =   600
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Arrendador:"
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
Attribute VB_Name = "frmARF_Acreedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vEdita As Boolean, vCodigo As Long
Dim vCantonMascara As String, vDistritoMascara As String, vFechaActual As Date
Dim vScroll As Boolean, vPaso As Boolean



Private Sub sbCuentas_Load()

On Error GoTo vError

lswCuentas.ListItems.Clear
    strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
           & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
           & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
           & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
           & " where C.Identificacion = '" & Trim(txtIdentificacion.Text) & "' and C.Modulo = 'ARF'"
    
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
   MsgBox "Consulte una Arrendador Primero...", vbExclamation
   tcMain.Item(0).Selected = True
   Exit Sub
End If

GLOBALES.gTag = Trim(txtIdentificacion)
GLOBALES.gTag2 = "ARF"

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


If txtCodigo = "" And FlatScrollBar.Value = 1 Then txtCodigo.Text = "0"
If txtCodigo = "" And FlatScrollBar.Value = 0 Then txtCodigo.Text = "999999999"

If vScroll Then
    strSQL = "select Top 1 COD_ACREEDOR from ARF_ACREEDORES"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_ACREEDOR > " & txtCodigo.Text & " order by COD_ACREEDOR asc"
    Else
       strSQL = strSQL & " where COD_ACREEDOR < " & txtCodigo.Text & " order by COD_ACREEDOR desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!COD_ACREEDOR
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
vModulo = 20
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 20

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


vCodigo = 0
txtCodigo.Text = ""

tcMain.Item(0).Selected = True

txtIdentificacion.Text = ""

txtNombre = ""
txtTelefono.Text = ""
txtTelefono2.Text = ""

txtWebSite.Text = ""
txtEmail.Text = ""
txtEmail2.Text = ""
txtAptoPostal.Text = ""

txtDireccion = ""
txtContacto.Text = ""

txtCuentaCod.Text = ""
txtCuentaDesc.Text = ""

chkActivo.Value = vbChecked

txtProveedor.Text = ""
txtProveedor.Tag = ""

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

strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)


End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtNombre.SetFocus
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
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "Descripcion"
       gBusquedas.Consulta = "select COD_ACREEDOR,Descripcion from ARF_ACREEDORES"
       frmBusquedas.Show vbModal
       
       If IsNumeric(gBusquedas.Resultado) Then
            txtCodigo.SetFocus
            txtCodigo = gBusquedas.Resultado
            txtNombre.SetFocus
       End If
    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(xCodigo As Long)

On Error GoTo vError

If Not fxSIFValidaCadena(txtCodigo.Text) Then
   Exit Sub
End If

Me.MousePointer = vbHourglass
'       & ",case when P.Tipo_Emision = 'TE' then 'Transferencia' when P.Tipo_Emision = 'CK' then 'Cheque' else 'Transferencia' end as 'Tipo_Pago' " _

strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & ", isnull(Ban.Descripcion, '') as 'BancoDesc', Tid.Descripcion as 'Tipo_Id_Desc'" _
       & ", isnull(Cta.Cod_Cuenta_Mask,P.cod_Cuenta) as 'COD_CUENTA_MASK', isnull(Cta.Descripcion,'') as 'CtaDesc'" _
       & ", isnull(Cpp.descripcion,'') as 'ProveedorDesc'" _
       & " from ARF_ACREEDORES P " _
       & " left join Provincias Prov on P.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Provincia = Cant.Provincia and P.Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Provincia = Dist.Provincia and convert(int,P.Canton) = convert(int,Dist.Canton) and P.distrito = Dist.distrito" _
       & " left join Tes_Bancos Ban on P.Cod_Banco = Ban.Id_Banco" _
       & " left join vCNTX_CUENTAS_LOCAL Cta on P.cod_Cuenta = Cta.Cod_Cuenta" _
       & " left join CxP_Proveedores Cpp on P.Cod_Proveedor = Cpp.Cod_Proveedor" _
       & " inner join AFI_Tipos_Ids Tid on P.Tipo_Id = Tid.Tipo_Id" _
       & " where P.COD_ACREEDOR = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCodigo = rs!COD_ACREEDOR
  txtCodigo = rs!COD_ACREEDOR

  txtNombre = rs!Descripcion & ""
  
  txtIdentificacion.Text = rs!Identificacion & ""
  txtTelefono.Text = rs!Telefono_01 & ""
  txtTelefono2.Text = rs!Telefono_02 & ""


  txtWebSite.Text = rs!WebSite & ""
  txtEmail.Text = rs!email_01 & ""
  txtEmail2.Text = rs!Email_02 & ""
  txtAptoPostal.Text = rs!apto_postal & ""

  txtContacto.Text = rs!Contacto_Nombre & ""

  Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")
  Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")
  Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
     
  cboDistrito.ToolTipText = Trim(rs!Distrito) & ""
  txtDireccion.Text = rs!direccion

  txtProveedor.Tag = rs!Cod_Proveedor & ""
  txtProveedor.Text = rs!ProveedorDesc
    
  txtCuentaCod.Text = rs!COD_CUENTA_MASK & ""
  txtCuentaDesc.Text = rs!CtaDesc & ""
  
  Call sbCboAsignaDato(cboBancos, rs!BancoDesc, True, rs!Cod_Banco)
  Call sbCboAsignaDato(cboTipoId, rs!Tipo_Id_Desc, True, rs!Tipo_Id)


Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If


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
If txtIdentificacion.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Número de Identificación no es válida! ..."

If Trim(txtCuentaDesc.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Cuenta Contable no es válida!..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim vProveedorId As Long

On Error GoTo vError

If IsNumeric(txtProveedor.Tag) Then
  vProveedorId = txtProveedor.Tag
Else
  vProveedorId = 0
End If

If vEdita Then
  strSQL = "update ARF_ACREEDORES set Descripcion = '" & Trim(txtNombre.Text) & "', Identificacion = '" & txtIdentificacion.Text _
         & "',Telefono_01 = '" & txtTelefono.Text & "',Telefono_02 = '" & txtTelefono2.Text _
         & "', WebSite = '" & txtWebSite.Text & "', apto_postal = '" & txtAptoPostal _
         & "', email_01 = '" & txtEmail & "', email_02 = '" & txtEmail2.Text & "',direccion = '" & txtDireccion _
         & "', distrito = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "',canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
         & "', provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
         & "', Contacto_Nombre = '" & txtContacto.Text & "', Activo = " & chkActivo.Value _
         & " , cod_Banco = " & cboBancos.ItemData(cboBancos.ListIndex) _
         & " , Cod_Cuenta = '" & fxgCntCuentaFormato(False, txtCuentaCod.Text, 0) & "'" _
         & " , cod_Proveedor = " & vProveedorId & ", Tipo_Id = " & cboTipoId.ItemData(cboTipoId.ListIndex) _
         & " , Modifica_Fecha = getdate(), Modifica_Usuario = '" & glogon.Usuario & "'" _
         & " where COD_ACREEDOR = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Arrendador: " & vCodigo)

Else

   strSQL = "select isnull(max(COD_ACREEDOR),0) as ultimo from ARF_ACREEDORES"
   Call OpenRecordSet(rs, strSQL)
     txtCodigo.Text = rs!ultimo + 1
     vCodigo = txtCodigo
   rs.Close
   

   strSQL = "insert into ARF_ACREEDORES(COD_ACREEDOR, Descripcion, Tipo_Id, Identificacion, Telefono_01, Telefono_02, Activo" _
          & ", apto_postal, email_01, email_02, WebSite, provincia, canton, distrito, direccion" _
          & ", Contacto_Nombre, cod_banco, cod_Cuenta, cod_Proveedor, registro_fecha, registro_usuario)" _
          & " values(" & vCodigo & ", '" & txtNombre.Text & "', " & cboTipoId.ItemData(cboTipoId.ListIndex) & ", '" & txtIdentificacion.Text _
          & "', '" & txtTelefono.Text & "', '" & txtTelefono2.Text & "', " & chkActivo.Value _
          & ",'" & txtAptoPostal.Text & "', '" & txtEmail.Text & "', '" & txtEmail2.Text & "', '" & txtWebSite.Text _
          & "', '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "', '" & cboCanton.ItemData(cboCanton.ListIndex) _
          & "', '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "', '" & txtDireccion.Text _
          & "', '" & txtContacto.Text & "', " & cboBancos.ItemData(cboBancos.ListIndex) _
          & ", '" & fxgCntCuentaFormato(False, txtCuentaCod.Text) _
          & "', " & vProveedorId & ", Getdate(), '" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Arrendador: " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(txtCodigo.Text)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete ARF_ACREEDORES where COD_ACREEDOR = " & vCodigo
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Arrendador: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COD_ACREEDOR"
  gBusquedas.Orden = "COD_ACREEDOR"
  gBusquedas.Consulta = "select COD_ACREEDOR, Descripcion from ARF_ACREEDORES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If IsNumeric(gBusquedas.Resultado) Then
    vCodigo = gBusquedas.Resultado
    Call sbConsulta(vCodigo)
  End If

End If

End Sub


Private Sub txtCodigo_LostFocus()
If IsNumeric(txtCodigo) Then
  Call sbConsulta(txtCodigo.Text)
End If
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
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtIdentificacion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"
  gBusquedas.Consulta = "select COD_ACREEDOR, Descripcion from ARF_ACREEDORES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  
  If IsNumeric(gBusquedas.Resultado) Then
        vCodigo = gBusquedas.Resultado
        Call sbConsulta(vCodigo)
  End If
End If

End Sub

Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtIdentificacion.SetFocus

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
  If IsNumeric(gBusquedas.Resultado) Then
    txtProveedor.Text = gBusquedas.Resultado3
    txtProveedor.Tag = gBusquedas.Resultado
  End If
End If
End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtWebSite.SetFocus
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

Private Sub txtWebSite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub



