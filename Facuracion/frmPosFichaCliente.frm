VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmPosFichaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha del Cliente"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   11265
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   3480
      Top             =   0
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   10440
      TabIndex        =   1
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   6972
      _Version        =   1310723
      _ExtentX        =   12298
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   2052
      _Version        =   1310723
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6612
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   11292
      _Version        =   1310723
      _ExtentX        =   19918
      _ExtentY        =   11663
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
      ItemCount       =   1
      Item(0).Caption =   "General"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "txtRazon"
      Item(0).Control(1)=   "cboTipoId"
      Item(0).Control(2)=   "GroupBox2"
      Item(0).Control(3)=   "GroupBox3"
      Item(0).Control(4)=   "chkActivo"
      Item(0).Control(5)=   "Label4(8)"
      Item(0).Control(6)=   "Label4(9)"
      Item(0).Control(7)=   "gbDireccion"
      Item(0).Control(8)=   "cboPlanilla"
      Item(0).Control(9)=   "Label4(3)"
      Item(0).Control(10)=   "chkExcento"
      Item(0).Control(11)=   "chkCredito"
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   252
         Left            =   2640
         TabIndex        =   6
         Top             =   1080
         Width           =   1092
         _Version        =   1310723
         _ExtentX        =   1926
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
         Appearance      =   2
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1572
         Left            =   240
         TabIndex        =   7
         Top             =   5160
         Width           =   10812
         _Version        =   1310723
         _ExtentX        =   19071
         _ExtentY        =   2773
         _StockProps     =   79
         Caption         =   "Datos Adicionales"
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
         Begin XtremeSuiteControls.ComboBox cboSexo 
            Height          =   312
            Left            =   1440
            TabIndex        =   8
            Top             =   360
            Width           =   2052
            _Version        =   1310723
            _ExtentX        =   3625
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   312
            Left            =   1440
            TabIndex        =   9
            Top             =   720
            Width           =   2052
            _Version        =   1310723
            _ExtentX        =   3625
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.DateTimePicker dtpFecNac 
            Height          =   312
            Left            =   1440
            TabIndex        =   10
            Top             =   1080
            Width           =   2052
            _Version        =   1310723
            _ExtentX        =   3619
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
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   1092
            Left            =   3840
            TabIndex        =   11
            Top             =   360
            Width           =   6852
            _Version        =   1310723
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   16
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1332
            _Version        =   1310723
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Sexo"
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
            Index           =   15
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   1332
            _Version        =   1310723
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Est. Civil"
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
            Index           =   14
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   1332
            _Version        =   1310723
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fec. Nac."
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
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1452
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   10812
         _Version        =   1310723
         _ExtentX        =   19071
         _ExtentY        =   2561
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
            Top             =   720
            Width           =   5292
            _Version        =   1310723
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail 
            Height          =   312
            Left            =   5400
            TabIndex        =   17
            Top             =   360
            Width           =   5292
            _Version        =   1310723
            _ExtentX        =   9334
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAptoPostal 
            Height          =   312
            Left            =   5400
            TabIndex        =   18
            Top             =   1080
            Width           =   5292
            _Version        =   1310723
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono 
            Height          =   312
            Left            =   1440
            TabIndex        =   19
            Top             =   720
            Width           =   2052
            _Version        =   1310723
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
         Begin XtremeSuiteControls.FlatEdit txtTelefono2 
            Height          =   312
            Left            =   1440
            TabIndex        =   20
            Top             =   1080
            Width           =   2052
            _Version        =   1310723
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
         Begin XtremeSuiteControls.FlatEdit txtCelular 
            Height          =   312
            Left            =   1440
            TabIndex        =   21
            Top             =   360
            Width           =   2052
            _Version        =   1310723
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   360
            Width           =   1332
            _Version        =   1310723
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tel. Móvil"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   1
            Left            =   0
            TabIndex        =   26
            Top             =   720
            Width           =   1332
            _Version        =   1310723
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   2
            Left            =   0
            TabIndex        =   25
            Top             =   1080
            Width           =   1332
            _Version        =   1310723
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
            Index           =   4
            Left            =   3840
            TabIndex        =   24
            Top             =   720
            Width           =   1332
            _Version        =   1310723
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
            Index           =   5
            Left            =   3840
            TabIndex        =   23
            Top             =   360
            Width           =   1332
            _Version        =   1310723
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Email "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   7
            Left            =   3840
            TabIndex        =   22
            Top             =   1080
            Width           =   1332
            _Version        =   1310723
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
      End
      Begin XtremeSuiteControls.ComboBox cboTipoId 
         Height          =   312
         Left            =   1680
         TabIndex        =   28
         Top             =   480
         Width           =   2052
         _Version        =   1310723
         _ExtentX        =   3625
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtRazon 
         Height          =   552
         Left            =   5640
         TabIndex        =   29
         Top             =   480
         Width           =   5292
         _Version        =   1310723
         _ExtentX        =   9334
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
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbDireccion 
         Height          =   2052
         Left            =   240
         TabIndex        =   30
         Top             =   3480
         Width           =   10812
         _Version        =   1310723
         _ExtentX        =   19071
         _ExtentY        =   3619
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
            TabIndex        =   31
            Top             =   480
            Width           =   2052
            _Version        =   1310723
            _ExtentX        =   3625
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
         End
         Begin XtremeSuiteControls.ComboBox cboCanton 
            Height          =   312
            Left            =   1440
            TabIndex        =   32
            Top             =   840
            Width           =   2052
            _Version        =   1310723
            _ExtentX        =   3625
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
         End
         Begin XtremeSuiteControls.ComboBox cboDistrito 
            Height          =   312
            Left            =   1440
            TabIndex        =   33
            Top             =   1200
            Width           =   2052
            _Version        =   1310723
            _ExtentX        =   3625
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtDireccion 
            Height          =   1092
            Left            =   3840
            TabIndex        =   34
            Top             =   480
            Width           =   6852
            _Version        =   1310723
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   13
            Left            =   120
            TabIndex        =   37
            Top             =   480
            Width           =   1332
            _Version        =   1310723
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   252
            Index           =   12
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   1332
            _Version        =   1310723
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
            Index           =   11
            Left            =   120
            TabIndex        =   35
            Top             =   1200
            Width           =   1332
            _Version        =   1310723
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
      End
      Begin XtremeSuiteControls.ComboBox cboPlanilla 
         Height          =   312
         Left            =   5640
         TabIndex        =   40
         Top             =   1080
         Width           =   5292
         _Version        =   1310723
         _ExtentX        =   9340
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.CheckBox chkExcento 
         Height          =   252
         Left            =   5640
         TabIndex        =   42
         Top             =   1440
         Width           =   1812
         _Version        =   1310723
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cliente Exento?"
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
      End
      Begin XtremeSuiteControls.CheckBox chkCredito 
         Height          =   252
         Left            =   7800
         TabIndex        =   43
         Top             =   1440
         Width           =   1932
         _Version        =   1310723
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Credito Abierto?"
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
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   252
         Index           =   3
         Left            =   4080
         TabIndex        =   41
         Top             =   1080
         Width           =   1452
         _Version        =   1310723
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Entidad/Empresa"
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
         Left            =   240
         TabIndex        =   39
         Top             =   480
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo Id"
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
         Left            =   4080
         TabIndex        =   38
         Top             =   480
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Razón Social"
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   252
      Index           =   17
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cliente"
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
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPosFichaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean, vTipoJuridica As Integer, vPaso As Boolean


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

Private Sub cboTipoId_Click()
If vPaso Then Exit Sub

If cboTipoId.ItemData(cboTipoId.ListIndex) = vTipoJuridica Then
   dtpFecNac.Enabled = False
   cboSexo.Enabled = False
   cboEstado.Enabled = False

   txtRazon.Enabled = True
   txtRazon.BackColor = vbWhite
  
Else
   dtpFecNac.Enabled = True
   cboSexo.Enabled = True
   cboEstado.Enabled = True
   txtRazon.Enabled = False
   txtRazon.BackColor = RGB(84, 153, 199)
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 cedula from pv_clientes"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cedula > '" & txtCedula.Text & "' order by cedula asc"
    Else
       strSQL = strSQL & " where cedula < '" & txtCedula.Text & "' order by cedula desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCedula.Text = rs!Cedula
      Call txtCedula_LostFocus
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
vModulo = 33
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 33
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla


 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()
vCodigo = ""
txtCedula = ""

txtNombre = ""
txtRazon.Text = ""

txtTelefono = ""
txtTelefono2 = ""
txtCelular = ""

txtWebSite.Text = ""

dtpFecNac.Value = fxFechaServidor
cboEstado.Text = "Soltero"
cboSexo.Text = "Masculino"

txtDireccion = ""

txtNotas = ""
txtAptoPostal = ""
txtEmail = ""

chkActivo.Value = xtpChecked

End Sub


Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

TimerX.Enabled = False
TimerX.Interval = 0

'vFechaActual = Format(fxFechaServidor, "dd/mm/yyyy")


'Revisa cual Tipo de Identificacion es Juridica (Solo es Valido la Primera)
vTipoJuridica = 0
strSQL = "select TIPO_ID from AFI_TIPOS_IDS where Tipo_Personeria = 'J'"
rs.Open strSQL, glogon.Conection
If Not rs.EOF And Not rs.BOF Then
    vTipoJuridica = rs!Tipo_id
End If
rs.Close

'Carga Tipos de Identificacion
vPaso = True
strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
vPaso = False
Call cboTipoId_Click

''Carga Clasificaciones de Clientes
'strSQL = "select rtrim(cod_categoria) as 'IdX' , rtrim(descripcion) as ItmX from CxC_Categoria_Clientes"
'Call sbCbo_Llena_New(cboClasificacion, strSQL, False, True)

'Carga combo de Provincias
vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False
 
strSQL = "select cod_institucion as 'IdX', rtrim(descripcion) as 'ItmX' from instituciones where Activa = 1"
Call sbCbo_Llena_New(cboPlanilla, strSQL, False, True)



'Sexo
cboSexo.Clear
cboSexo.AddItem "Masculino"
cboSexo.ItemData(cboSexo.ListCount - 1) = "M"
cboSexo.AddItem "Femenido"
cboSexo.ItemData(cboSexo.ListCount - 1) = "F"
Call sbCboAsignaDato(cboSexo, "Femenino", True, "F")
 
strSQL = "select rtrim(Estado_Civil) as 'IdX', rtrim(DESCRIPCION) as 'ItmX'" _
       & " from SYS_ESTADO_CIVIL where ACTIVO = 1"
Call sbCbo_Llena_New(cboEstado, strSQL, False, True)

Call sbLimpiaPantalla

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCedula.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCedula.SetFocus
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
       gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
       frmBusquedas.Show vbModal
       txtCedula.SetFocus
       txtCedula = gBusquedas.Resultado
       txtNombre.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

'Verifica el Enlaces con Bases de Clientes
Call sbXFichaCliente(xCodigo)

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & ",Tid.Descripcion as TipoIdDesc,Tid.Tipo_Personeria" _
       & ",isnull(Ec.Descripcion,'No. Identificado') as 'EstadoCivilDesc' " _
       & ",isnull(Inst.Descripcion,'No. Identificado') as 'InstitucionDesc' " _
       & " from pv_clientes P " _
       & " left join Provincias Prov on P.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Provincia = Cant.Provincia and P.Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Provincia = Dist.Provincia and P.Canton = Dist.Canton and P.distrito = Dist.distrito" _
       & " left join AFI_TIPOS_IDS Tid on P.tipo_id = Tid.tipo_id" _
       & " left join SYS_ESTADO_CIVIL Ec on P.EstadoCivil = Ec.Estado_Civil" _
       & " left join INSTITUCIONES Inst on P.cod_institucion = Inst.COD_INSTITUCION" _
       & " where P.cedula = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)


If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!Cedula
  txtCedula.Text = rs!Cedula
 
  txtNombre.Text = rs!Nombre & ""
     
  txtTelefono.Text = rs!telefono1 & ""
  txtTelefono2.Text = rs!telefono2 & ""
  txtCelular.Text = rs!Celular & ""
  txtEmail.Text = rs!Email & ""
  
  
  Call sbCboAsignaDato(cboPlanilla, rs!InstitucionDesc, True, rs!cod_institucion)
  
  Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
  Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
  Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
  cboDistrito.ToolTipText = Trim(rs!distrito) & ""
  
  txtDireccion.Text = rs!Direccion

  Call sbCboAsignaDato(cboTipoId, rs!TipoIdDesc, True, rs!Tipo_id)
  Call sbCboAsignaDato(cboSexo, IIf((rs!sexo = "M"), "Masculino", "Femenino"), True, rs!sexo)
  Call sbCboAsignaDato(cboEstado, rs!EstadoCivilDesc, True, rs!EstadoCivil)

  txtRazon.Text = rs!Razon_Social & ""
  txtWebSite.Text = rs!WebSite & ""
  txtAptoPostal = rs!apto_postal & ""
  
  dtpFecNac.Value = rs!fecha_nacimiento
  
  txtNotas = rs!Notas & ""
  
  chkCredito.Value = rs!Credito_Cerrado
  chkExcento.Value = rs!Cliente_Excento
  chkActivo.Value = rs!activo
  
  Call cboTipoId_Click
  
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

If txtCedula.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Identificación no es válida ..."
If txtNombre.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre no es válido ..."
If txtEmail.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No se ha indicado Correo Electrónico..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vEstadoCivil As String

On Error GoTo vError

If cboTipoId.ItemData(cboTipoId.ListIndex) = vTipoJuridica Then
    vEstadoCivil = "O"
Else
    vEstadoCivil = cboEstado.ItemData(cboEstado.ListIndex)
End If


If vEdita Then
  strSQL = "update pv_clientes set nombre = '" & Trim(txtNombre.Text) & "', Razon_Social = '" & Trim(txtRazon.Text) & "'" _
         & ",telefono1 = '" & Trim(txtTelefono.Text) & "',telefono2 = '" & Trim(txtTelefono2.Text) _
         & "',celular = '" & Trim(txtCelular.Text) & "', WebSite = '" & Trim(txtWebSite.Text) _
         & "',sexo = '" & Mid(cboSexo.Text, 1, 1) _
         & "',EstadoCivil = '" & vEstadoCivil _
         & "',Fecha_nacimiento = '" & Format(dtpFecNac.Value, "yyyy/mm/dd") _
         & "',notas = '" & txtNotas & "',apto_postal = '" & txtAptoPostal _
         & "',email = '" & txtEmail & "',direccion = '" & txtDireccion _
         & "',distrito = '" & cboDistrito.ItemData(cboDistrito.ListIndex) _
         & "',canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
         & "',provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) _
         & " ,cod_institucion = " & cboPlanilla.ItemData(cboPlanilla.ListIndex) _
         & " ,Tipo_Id = " & cboTipoId.ItemData(cboTipoId.ListIndex) _
         & " ,credito_cerrado = " & chkCredito.Value & ",cliente_excento = " & chkExcento.Value _
         & " ,Activo = " & chkActivo.Value _
         & " where cedula = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Cliente: " & vCodigo)

Else
  vCodigo = txtCedula
   
   strSQL = "insert into pv_clientes(tipo_id,cedula,nombre,razon_social, celular,telefono1,telefono2,WebSite, sexo,estadoCivil,fecha_nacimiento" _
          & ",apto_postal,email,notas,direccion,distrito,provincia,canton,cod_institucion,credito_cerrado,cliente_excento, activo)" _
          & " values(" & cboTipoId.ItemData(cboTipoId.ListIndex) & ",'" & vCodigo & "','" & Trim(txtNombre.Text) & "','" & Trim(txtRazon.Text) _
          & "','" & Trim(txtCelular.Text) & "','" & Trim(txtTelefono.Text) & "','" & Trim(txtTelefono2.Text) & "','" & Trim(txtWebSite.Text) & "','" _
          & Mid(cboSexo.Text, 1, 1) & "','" & vEstadoCivil & "','" _
          & Format(dtpFecNac.Value, "yyyy/mm/dd") & "','" & Trim(txtAptoPostal.Text) & "','" & Trim(txtEmail.Text) & "','" & Trim(txtNotas.Text) _
          & "','" & Trim(txtDireccion.Text) & "','" & cboDistrito.ItemData(cboDistrito.ListIndex) _
          & "','" & cboProvincia.ItemData(cboProvincia.ListIndex) & "','" & cboCanton.ItemData(cboCanton.ListIndex) _
          & "'," & cboPlanilla.ItemData(cboPlanilla.ListIndex) & "," & chkCredito.Value & "," & chkExcento.Value _
          & "," & chkActivo.Value & ")"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Cliente: " & vCodigo)
 
End If

Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete pv_clientes where cedula = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Cliente : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
  
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCelular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCedula <> "" And vEdita Then Call sbConsulta(txtCedula)
  txtNombre.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  If txtCedula <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCedula_LostFocus()
txtNombre = fxSIFCCodigos("D", txtCedula, "Clientes")

If txtCedula.Text <> "" Then
    Call sbConsulta(txtCedula.Text)
End If


End Sub


Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboSexo.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  If txtCedula <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCelular.SetFocus
End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub txtDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboSexo.SetFocus
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFecNac.SetFocus
End Sub

Private Sub dtpFecNac_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboPlanilla.SetFocus
End Sub

