VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_Prendas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informacion de la Garantia Prendaria"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   330
      Left            =   6360
      TabIndex        =   0
      Top             =   1560
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cerrar"
            Object.ToolTipText     =   "Cierra esta ventana"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   4815
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
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtExpediente 
      Height          =   315
      Left            =   7680
      TabIndex        =   5
      Top             =   240
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      Top             =   600
      Width           =   4935
      _Version        =   1441793
      _ExtentX        =   8705
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6375
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   10335
      _Version        =   1441793
      _ExtentX        =   18230
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
      Color           =   128
      ItemCount       =   4
      SelectedItem    =   1
      Item(0).Caption =   "General"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Garantía"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "tcAux"
      Item(1).Control(1)=   "tcGarantia"
      Item(2).Caption =   "Trámites"
      Item(2).ControlCount=   0
      Item(3).Caption =   "Seguimiento del Trámite"
      Item(3).ControlCount=   11
      Item(3).Control(0)=   "Label3(36)"
      Item(3).Control(1)=   "Label3(37)"
      Item(3).Control(2)=   "Label3(49)"
      Item(3).Control(3)=   "Label3(50)"
      Item(3).Control(4)=   "lblRegistroFecha"
      Item(3).Control(5)=   "lblRegistroFechaAbog"
      Item(3).Control(6)=   "lblRegistroUsuario"
      Item(3).Control(7)=   "lblRegistroUsuarioAbog"
      Item(3).Control(8)=   "ShortcutCaption1(2)"
      Item(3).Control(9)=   "gbInfoNotarial"
      Item(3).Control(10)=   "gbTramite"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   6735
         Left            =   -69880
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   10095
         _Version        =   1441793
         _ExtentX        =   17806
         _ExtentY        =   11880
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
      End
      Begin XtremeSuiteControls.GroupBox gbTramite 
         Height          =   4335
         Left            =   -70000
         TabIndex        =   110
         Top             =   2160
         Visible         =   0   'False
         Width           =   10455
         _Version        =   1441793
         _ExtentX        =   18441
         _ExtentY        =   7646
         _StockProps     =   79
         Caption         =   "Trámites de Seguimiento"
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswTramite 
            Height          =   3255
            Left            =   120
            TabIndex        =   111
            Top             =   600
            Width           =   10095
            _Version        =   1441793
            _ExtentX        =   17806
            _ExtentY        =   5741
            _StockProps     =   77
         End
         Begin XtremeSuiteControls.PushButton btnNotasTramite 
            Height          =   330
            Left            =   8880
            TabIndex        =   112
            Top             =   240
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Notas"
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
            Picture         =   "frmCR_Prendas_New.frx":0000
         End
      End
      Begin XtremeSuiteControls.GroupBox gbInfoNotarial 
         Height          =   1455
         Left            =   -70000
         TabIndex        =   95
         Top             =   735
         Visible         =   0   'False
         Width           =   10455
         _Version        =   1441793
         _ExtentX        =   18441
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Información Notarial"
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtNotario 
            Height          =   315
            Left            =   960
            TabIndex        =   103
            Top             =   480
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtR_Usuario 
            Height          =   315
            Left            =   6360
            TabIndex        =   104
            Top             =   480
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtM_Usuario 
            Height          =   315
            Left            =   6360
            TabIndex        =   105
            Top             =   840
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtR_Fecha 
            Height          =   315
            Left            =   8160
            TabIndex        =   106
            Top             =   480
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtM_Fecha 
            Height          =   315
            Left            =   8160
            TabIndex        =   107
            Top             =   840
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
            _ExtentY        =   556
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTomo 
            Height          =   315
            Left            =   960
            TabIndex        =   108
            Top             =   840
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtAsiento 
            Height          =   315
            Left            =   3480
            TabIndex        =   109
            Top             =   840
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   9
            Left            =   5160
            TabIndex        =   102
            Top             =   840
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Modificación"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   5
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   8
            Left            =   5280
            TabIndex        =   101
            Top             =   480
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Registro"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   5
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   7
            Left            =   8280
            TabIndex        =   100
            Top             =   240
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   6
            Left            =   6480
            TabIndex        =   99
            Top             =   240
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Usuario"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   5
            Left            =   2520
            TabIndex        =   98
            Top             =   840
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Asiento"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   97
            Top             =   840
            Width           =   735
            _Version        =   1441793
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tomo"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   96
            Top             =   480
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Notario"
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
      Begin XtremeSuiteControls.TabControl tcGarantia 
         Height          =   3375
         Left            =   0
         TabIndex        =   36
         Top             =   360
         Width           =   10215
         _Version        =   1441793
         _ExtentX        =   18018
         _ExtentY        =   5953
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   2
         PaintManager.Position=   2
         PaintManager.ShowTabs=   0   'False
         ItemCount       =   2
         SelectedItem    =   1
         Item(0).Caption =   "General"
         Item(0).ControlCount=   14
         Item(0).Control(0)=   "txtModelo"
         Item(0).Control(1)=   "txtSerie"
         Item(0).Control(2)=   "txtDescripcion"
         Item(0).Control(3)=   "txtCobertura"
         Item(0).Control(4)=   "txtMarca"
         Item(0).Control(5)=   "txtAvaluo"
         Item(0).Control(6)=   "txtCoberturaPorc"
         Item(0).Control(7)=   "Label1(6)"
         Item(0).Control(8)=   "Label1(2)"
         Item(0).Control(9)=   "Label2(1)"
         Item(0).Control(10)=   "Label1(4)"
         Item(0).Control(11)=   "Label1(5)"
         Item(0).Control(12)=   "Label1(0)"
         Item(0).Control(13)=   "Label1(1)"
         Item(1).Caption =   "Vehicular"
         Item(1).ControlCount=   36
         Item(1).Control(0)=   "Label4(0)"
         Item(1).Control(1)=   "Label4(1)"
         Item(1).Control(2)=   "Label4(2)"
         Item(1).Control(3)=   "Label4(3)"
         Item(1).Control(4)=   "Label4(4)"
         Item(1).Control(5)=   "Label4(6)"
         Item(1).Control(6)=   "Label4(7)"
         Item(1).Control(7)=   "Label4(8)"
         Item(1).Control(8)=   "Label4(9)"
         Item(1).Control(9)=   "Label4(5)"
         Item(1).Control(10)=   "Label4(10)"
         Item(1).Control(11)=   "Label4(11)"
         Item(1).Control(12)=   "Label4(12)"
         Item(1).Control(13)=   "Label4(13)"
         Item(1).Control(14)=   "Label4(14)"
         Item(1).Control(15)=   "Label4(15)"
         Item(1).Control(16)=   "Label4(16)"
         Item(1).Control(17)=   "Label4(17)"
         Item(1).Control(18)=   "Label4(18)"
         Item(1).Control(19)=   "txtV_Modelo"
         Item(1).Control(20)=   "txtV_PlacaRegistral"
         Item(1).Control(21)=   "txtV_PlacaProvisional"
         Item(1).Control(22)=   "txtV_Color"
         Item(1).Control(23)=   "txtV_Anio"
         Item(1).Control(24)=   "cboV_Marca"
         Item(1).Control(25)=   "cboV_Combustible"
         Item(1).Control(26)=   "txtV_Chasis"
         Item(1).Control(27)=   "txtV_Capacidad"
         Item(1).Control(28)=   "txtV_CilindrajeUd"
         Item(1).Control(29)=   "txtV_Peso"
         Item(1).Control(30)=   "txtV_Puertas"
         Item(1).Control(31)=   "txtV_Cilindraje"
         Item(1).Control(32)=   "cboV_Uso"
         Item(1).Control(33)=   "txtV_VIN"
         Item(1).Control(34)=   "cboV_Presentacion"
         Item(1).Control(35)=   "cboV_Comercializa"
         Begin XtremeSuiteControls.FlatEdit txtModelo 
            Height          =   315
            Left            =   -68560
            TabIndex        =   37
            Top             =   2040
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3831
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
         Begin XtremeSuiteControls.FlatEdit txtSerie 
            Height          =   315
            Left            =   -68560
            TabIndex        =   38
            Top             =   2400
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3831
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
         Begin XtremeSuiteControls.FlatEdit txtDescripcion 
            Height          =   915
            Left            =   -68560
            TabIndex        =   39
            Top             =   600
            Visible         =   0   'False
            Width           =   7815
            _Version        =   1441793
            _ExtentX        =   13785
            _ExtentY        =   1614
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
         Begin XtremeSuiteControls.FlatEdit txtCobertura 
            Height          =   330
            Left            =   -62920
            TabIndex        =   40
            Top             =   2400
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3831
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMarca 
            Height          =   315
            Left            =   -68560
            TabIndex        =   41
            Top             =   1680
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3831
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
         Begin XtremeSuiteControls.FlatEdit txtAvaluo 
            Height          =   330
            Left            =   -62920
            TabIndex        =   42
            Top             =   1680
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3831
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
         Begin XtremeSuiteControls.FlatEdit txtCoberturaPorc 
            Height          =   315
            Left            =   -62920
            TabIndex        =   43
            Top             =   2040
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3831
            _ExtentY        =   550
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtV_Modelo 
            Height          =   315
            Left            =   6600
            TabIndex        =   70
            Top             =   120
            Width           =   2895
            _Version        =   1441793
            _ExtentX        =   5106
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtV_PlacaRegistral 
            Height          =   315
            Left            =   1920
            TabIndex        =   71
            Top             =   120
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtV_PlacaProvisional 
            Height          =   315
            Left            =   1920
            TabIndex        =   72
            Top             =   480
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtV_Color 
            Height          =   315
            Left            =   1920
            TabIndex        =   73
            Top             =   840
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtV_Anio 
            Height          =   315
            Left            =   1920
            TabIndex        =   74
            Top             =   1200
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.ComboBox cboV_Marca 
            Height          =   330
            Left            =   1920
            TabIndex        =   75
            Top             =   1560
            Width           =   3255
            _Version        =   1441793
            _ExtentX        =   5741
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
         Begin XtremeSuiteControls.ComboBox cboV_Presentacion 
            Height          =   330
            Left            =   1920
            TabIndex        =   76
            Top             =   1920
            Width           =   3255
            _Version        =   1441793
            _ExtentX        =   5741
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
         Begin XtremeSuiteControls.ComboBox cboV_Combustible 
            Height          =   330
            Left            =   1920
            TabIndex        =   77
            Top             =   2280
            Width           =   3255
            _Version        =   1441793
            _ExtentX        =   5741
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
         Begin XtremeSuiteControls.ComboBox cboV_Comercializa 
            Height          =   330
            Left            =   1920
            TabIndex        =   78
            Top             =   2760
            Width           =   3255
            _Version        =   1441793
            _ExtentX        =   5741
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
         Begin XtremeSuiteControls.FlatEdit txtV_Chasis 
            Height          =   315
            Left            =   6600
            TabIndex        =   79
            Top             =   480
            Width           =   2895
            _Version        =   1441793
            _ExtentX        =   5106
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtV_CilindrajeUd 
            Height          =   315
            Left            =   8520
            TabIndex        =   81
            Top             =   1920
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtV_Capacidad 
            Height          =   315
            Left            =   6600
            TabIndex        =   80
            Top             =   840
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtV_Puertas 
            Height          =   315
            Left            =   6600
            TabIndex        =   83
            Top             =   1560
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtV_Cilindraje 
            Height          =   315
            Left            =   6600
            TabIndex        =   84
            Top             =   1920
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.ComboBox cboV_Uso 
            Height          =   330
            Left            =   6600
            TabIndex        =   85
            Top             =   2280
            Width           =   2895
            _Version        =   1441793
            _ExtentX        =   5106
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
         Begin XtremeSuiteControls.FlatEdit txtV_VIN 
            Height          =   315
            Left            =   6600
            TabIndex        =   86
            Top             =   2760
            Width           =   2895
            _Version        =   1441793
            _ExtentX        =   5106
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtV_Peso 
            Height          =   315
            Left            =   6600
            TabIndex        =   82
            Top             =   1200
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   18
            Left            =   5520
            TabIndex        =   69
            Top             =   2760
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "VIN: "
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
            Index           =   17
            Left            =   5520
            TabIndex        =   68
            Top             =   2280
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Uso: "
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
            Index           =   16
            Left            =   7800
            TabIndex        =   67
            Top             =   1920
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Unidad: "
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
            Index           =   15
            Left            =   5520
            TabIndex        =   66
            Top             =   1920
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cilindraje: "
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
            Index           =   14
            Left            =   5520
            TabIndex        =   65
            Top             =   1560
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Puertas: "
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
            Index           =   13
            Left            =   7800
            TabIndex        =   64
            Top             =   1200
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Kg"
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
            Index           =   12
            Left            =   5520
            TabIndex        =   63
            Top             =   1200
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Peso: "
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
            Index           =   11
            Left            =   7800
            TabIndex        =   62
            Top             =   840
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Personas"
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
            Index           =   10
            Left            =   5520
            TabIndex        =   61
            Top             =   840
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Capacidad: "
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
            Index           =   5
            Left            =   5520
            TabIndex        =   60
            Top             =   480
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Chasís: "
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
            Index           =   9
            Left            =   5520
            TabIndex        =   59
            Top             =   120
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Modelo: "
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
            Index           =   8
            Left            =   240
            TabIndex        =   58
            Top             =   2760
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Comercializa: "
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
            Index           =   7
            Left            =   240
            TabIndex        =   57
            Top             =   2280
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo Combustible: "
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
            Index           =   6
            Left            =   240
            TabIndex        =   56
            Top             =   1920
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Presentación: "
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
            Index           =   4
            Left            =   240
            TabIndex        =   55
            Top             =   1560
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Marca: "
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
            Index           =   3
            Left            =   240
            TabIndex        =   54
            Top             =   1200
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Año Fabricación: "
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
            TabIndex        =   53
            Top             =   840
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Color: "
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
            TabIndex        =   52
            Top             =   480
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Placa Provisional:"
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
            TabIndex        =   51
            Top             =   120
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Placa Registral: "
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Avalúo"
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
            Left            =   -64240
            TabIndex        =   50
            Top             =   1680
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Serie/Año"
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
            Left            =   -69880
            TabIndex        =   49
            Top             =   2400
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
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
            Index           =   5
            Left            =   -69880
            TabIndex        =   48
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Marca"
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
            Index           =   4
            Left            =   -69880
            TabIndex        =   47
            Top             =   1680
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo"
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
            Left            =   -69880
            TabIndex        =   46
            Top             =   2040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cobertura"
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
            Left            =   -64240
            TabIndex        =   45
            Top             =   2400
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "% Cobertura"
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
            Index           =   6
            Left            =   -64240
            TabIndex        =   44
            Top             =   2040
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   2535
         Left            =   0
         TabIndex        =   13
         Top             =   3840
         Width           =   10335
         _Version        =   1441793
         _ExtentX        =   18230
         _ExtentY        =   4471
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "Anotaciones"
         Item(0).ControlCount=   4
         Item(0).Control(0)=   "txtObservaciones"
         Item(0).Control(1)=   "lswPolizas"
         Item(0).Control(2)=   "Label4(19)"
         Item(0).Control(3)=   "Label4(20)"
         Item(1).Caption =   "Avalúo"
         Item(1).ControlCount=   16
         Item(1).Control(0)=   "Label10(16)"
         Item(1).Control(1)=   "Label10(15)"
         Item(1).Control(2)=   "Label10(13)"
         Item(1).Control(3)=   "dtpFechaInspeccion"
         Item(1).Control(4)=   "txtAvaluo_Notas"
         Item(1).Control(5)=   "Label10(14)"
         Item(1).Control(6)=   "Label10(18)"
         Item(1).Control(7)=   "optPoliza(0)"
         Item(1).Control(8)=   "optPoliza(1)"
         Item(1).Control(9)=   "Label10(0)"
         Item(1).Control(10)=   "txtValorTotal"
         Item(1).Control(11)=   "txtExtras"
         Item(1).Control(12)=   "txtValorSExtras"
         Item(1).Control(13)=   "lblValorPrenda"
         Item(1).Control(14)=   "txtValorFiscal"
         Item(1).Control(15)=   "btnExtras"
         Begin XtremeSuiteControls.PushButton btnExtras 
            Height          =   330
            Left            =   -66040
            TabIndex        =   90
            Top             =   1560
            Visible         =   0   'False
            Width           =   370
            _Version        =   1441793
            _ExtentX        =   653
            _ExtentY        =   582
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_Prendas_New.frx":0720
         End
         Begin XtremeSuiteControls.ListView lswPolizas 
            Height          =   1695
            Left            =   5160
            TabIndex        =   87
            Top             =   720
            Width           =   5055
            _Version        =   1441793
            _ExtentX        =   8916
            _ExtentY        =   2990
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
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.FlatEdit txtObservaciones 
            Height          =   1695
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   5055
            _Version        =   1441793
            _ExtentX        =   8916
            _ExtentY        =   2990
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
         Begin XtremeSuiteControls.FlatEdit txtValorTotal 
            Height          =   315
            Left            =   -67960
            TabIndex        =   15
            Top             =   2040
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Text            =   "0.00"
            BackColor       =   16777152
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtExtras 
            Height          =   315
            Left            =   -67960
            TabIndex        =   16
            Top             =   1560
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Text            =   "0.00"
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtValorSExtras 
            Height          =   315
            Left            =   -67960
            TabIndex        =   17
            Top             =   1200
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Text            =   "0.00"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaInspeccion 
            Height          =   315
            Left            =   -67960
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   550
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.FlatEdit txtAvaluo_Notas 
            Height          =   975
            Left            =   -65440
            TabIndex        =   19
            Top             =   840
            Visible         =   0   'False
            Width           =   5535
            _Version        =   1441793
            _ExtentX        =   9763
            _ExtentY        =   1720
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
         Begin XtremeSuiteControls.RadioButton optPoliza 
            Height          =   255
            Index           =   0
            Left            =   -63760
            TabIndex        =   20
            Top             =   2040
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Factor"
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
         End
         Begin XtremeSuiteControls.RadioButton optPoliza 
            Height          =   255
            Index           =   1
            Left            =   -62080
            TabIndex        =   21
            Top             =   2040
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Personalizada"
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
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtValorFiscal 
            Height          =   315
            Left            =   -67960
            TabIndex        =   88
            Top             =   840
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Text            =   "0.00"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   20
            Left            =   5160
            TabIndex        =   92
            Top             =   480
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Coberturas de Pólizas: "
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
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   91
            Top             =   480
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anotaciones: "
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
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   0
            Left            =   -69880
            TabIndex        =   89
            Top             =   840
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valor Fiscal"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblValorPrenda 
            Height          =   255
            Left            =   -69880
            TabIndex        =   27
            Top             =   2040
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valor del Vehiculo"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   16
            Left            =   -69880
            TabIndex        =   26
            Top             =   1560
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Total Extras"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   15
            Left            =   -69880
            TabIndex        =   25
            Top             =   1200
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valor sin Extras"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   13
            Left            =   -69880
            TabIndex        =   24
            Top             =   480
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha de Avalúo"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   14
            Left            =   -65440
            TabIndex        =   23
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observaciones:"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   18
            Left            =   -65440
            TabIndex        =   22
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo de Póliza"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   2
         Left            =   -70000
         TabIndex        =   94
         Top             =   360
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1441793
         _ExtentX        =   18230
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seguimiento Notarial de Garantías Prendarias"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblRegistroFecha 
         Height          =   315
         Left            =   -67600
         TabIndex        =   35
         Top             =   6480
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblRegistroFechaAbog 
         Height          =   315
         Left            =   -62320
         TabIndex        =   34
         Top             =   6480
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblRegistroUsuario 
         Height          =   315
         Left            =   -67600
         TabIndex        =   33
         Top             =   6840
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblRegistroUsuarioAbog 
         Height          =   315
         Left            =   -62320
         TabIndex        =   32
         Top             =   6840
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   36
         Left            =   -69040
         TabIndex        =   31
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
         Height          =   255
         Index           =   37
         Left            =   -69040
         TabIndex        =   30
         Top             =   6840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   450
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
         Height          =   255
         Index           =   49
         Left            =   -63760
         TabIndex        =   29
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha"
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
         Height          =   255
         Index           =   50
         Left            =   -63760
         TabIndex        =   28
         Top             =   6840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   450
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
   End
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   9720
      TabIndex        =   93
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   1560
      Width           =   495
      _Version        =   1441793
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
      Picture         =   "frmCR_Prendas_New.frx":0E40
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   10335
      _Version        =   1441793
      _ExtentX        =   18230
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Registro de Prendas "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   240
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Operación"
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
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   600
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   450
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   7
      Top             =   240
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Expediente"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "frmCR_Prendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mPrendaId As Long

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean
Dim vEdita As Integer, mFecha As Date



Private Sub btnAdjuntos_Click()
 gGA.Modulo = "CR_01"
 gGA.Llave_01 = txtCedula.Text
 gGA.Llave_02 = txtOperacion.Text
 gGA.Llave_03 = "1" 'txtCodigo.Text
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End Sub

Private Function fxPrenda_Tipo(pTipo As String) As String
Dim vResult As String

strSQL = "select "

fxPrenda_Tipo = ""
End Function


Private Sub cboTipo_Click()
If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "select Formulario " _
       & " from crd_prendas_tipos Where Tipo_Prenda = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Formulario = "Vehículo" Then
    lblValorPrenda.Caption = "Valor del Vehículo"
    tcGarantia.Item(1).Selected = True
Else
    tcGarantia.Item(0).Selected = True
    lblValorPrenda.Caption = "Valor del Bien"
End If

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

scMain.Caption = "Prendas para la Operación: " & Operacion.Operacion

With lsw.ColumnHeaders
    .Clear
    .Add , , "Tipo", 1000
    .Add , , "Categoria", 2500
    .Add , , "Avaluo", 1800, vbRightJustify
    .Add , , "%", 1800, vbRightJustify
    .Add , , "Cobertura", 1800, vbRightJustify
    .Add , , "Descripción", 2500
    .Add , , "Modelo", 2500
    .Add , , "Serie", 2500
    .Add , , "Marca", 2500
End With


With lswTramite.ColumnHeaders
    .Clear
    .Add , , "Tramite", 3000
    .Add , , "Notas", 3000
    .Add , , "Usuario", 1900, vbCenter
    .Add , , "Fecha", 1900, vbCenter
End With


With lswPolizas.ColumnHeaders
  .Clear
  .Add , , "Coberturas de Pólizas", lswPolizas.Width - 250
End With

Dim itmX As ListViewItem
Set itmX = lswPolizas.ListItems.Add(, , "A) Obligatoria: Responsabilidad")
Set itmX = lswPolizas.ListItems.Add(, , "D) Obligatoria: Colisión y/o vuelco")
Set itmX = lswPolizas.ListItems.Add(, , "F) Obligatoria: Robo y Hurto")
Set itmX = lswPolizas.ListItems.Add(, , "H) Obligatoria: Riesgos adicional")
Set itmX = lswPolizas.ListItems.Add(, , "RC Obligatoria: (Responsabilidad...")

Call sbToolBarIconos(tlbPrincipal, False)

With tlbPrincipal
    .Buttons(1).Enabled = True
    .Buttons(2).Enabled = False
    .Buttons(3).Enabled = False
    .Buttons(4).Enabled = False
    .Buttons(5).Enabled = False
End With

'fra.Enabled = False

End Sub

Private Sub sbLimpia()
 txtDescripcion.Text = ""
 txtModelo.Text = ""
 txtSerie.Text = ""
 txtMarca.Text = ""
 
 txtAvaluo.Text = "0"
 txtCoberturaPorc.Text = "0"
 txtCobertura.Text = "0"
 
 mPrendaId = 0
 
End Sub

Private Sub sbPrendas_Load()

Dim curTotal As Currency

curTotal = 0

strSQL = "exec spCrd_Operacion_Prenda_Consulta " & Operacion.Operacion

Call OpenRecordSet(rs, strSQL)
lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Tipo_Prenda)
      
      itmX.SubItems(1) = rs!PrendaDesc
      
      
      itmX.SubItems(2) = Format(rs!Avaluo, "Standard")
      itmX.SubItems(3) = Format(rs!Porc_Cobertura, "Standard")
      itmX.SubItems(4) = Format(rs!Cobertura, "Standard")
      
      itmX.SubItems(5) = rs!Descripcion
      itmX.SubItems(6) = rs!Modelo
      itmX.SubItems(7) = rs!Serie
      itmX.SubItems(8) = rs!Marca
      
      itmX.Tag = rs!Prenda_Id
      
      curTotal = curTotal + rs!Cobertura
 rs.MoveNext
Loop
rs.Close

'txtCoberturaTotal.Text = Format(curTotal, "Standard")

End Sub




Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

 mPrendaId = Item.Tag
 
 strSQL = "exec spCrd_Operacion_Prenda_Consulta " & Operacion.Operacion & ", " & mPrendaId
 
 Call OpenRecordSet(rs, strSQL)
 
 
 txtDescripcion.Text = rs!Descripcion & ""
 txtModelo.Text = rs!Modelo & ""
 txtSerie.Text = rs!Serie & ""
 txtMarca.Text = rs!Marca & ""
 
 txtAvaluo.Text = Format(rs!Avaluo, "Standard")
 txtCoberturaPorc.Text = Format(rs!Porc_Cobertura, "Standard")
 txtCobertura.Text = Format(rs!Cobertura, "Standard")
 

 Call sbCboAsignaDato(cboTipo, rs!PrendaDesc, True, Trim(rs!Tipo_Prenda))
 

With tlbPrincipal
   .Buttons(1).Enabled = False
   .Buttons(2).Enabled = True
   .Buttons(3).Enabled = True
   .Buttons(4).Enabled = False
   .Buttons(5).Enabled = False
End With


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbInicializa()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select rtrim(tipo_prenda) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from crd_prendas_tipos where Activa = 1 order by descripcion "
vPaso = True
    Call sbCbo_Llena_New(cboTipo, strSQL, False, True)
vPaso = False
 
strSQL = "exec spCrd_Prendas_Cat_List_Cbo 'MAR'"
Call sbCbo_Llena_New(cboV_Marca, strSQL, False, True)
 
strSQL = "exec spCrd_Prendas_Cat_List_Cbo 'MOD'"
Call sbCbo_Llena_New(cboV_Presentacion, strSQL, False, True)
 
strSQL = "exec spCrd_Prendas_Cat_List_Cbo 'COB'"
Call sbCbo_Llena_New(cboV_Combustible, strSQL, False, True)
 
strSQL = "exec spCrd_Prendas_Cat_List_Cbo 'COM'"
Call sbCbo_Llena_New(cboV_Comercializa, strSQL, False, True)
  
cboV_Uso.AddItem "PERSONAL"
cboV_Uso.AddItem "TRABAJO"
cboV_Uso.Text = "PERSONAL"
  
Call cboTipo_Click
  
Me.MousePointer = vbDefault
 
Call sbPrendas_Load

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub


Private Sub txtAvaluo_GotFocus()
On Error GoTo vError
    txtAvaluo.Text = CCur(txtAvaluo.Text)
vError:

End Sub

Private Sub txtAvaluo_LostFocus()
On Error GoTo vError
    txtAvaluo.Text = Format(CCur(txtAvaluo.Text), "Standard")
vError:
End Sub


Private Sub txtAvaluo_KeyPress(KeyAscii As Integer)
On Error GoTo vError
  If KeyAscii = vbKeyReturn Then txtCoberturaPorc.SetFocus
vError:
End Sub

Private Function fxVerificaDatos() As Boolean
Dim vMensaje As String

fxVerificaDatos = True
vMensaje = ""

'Revision de Inyección

txtDescripcion.Text = fxSysCleanTxtInject(txtDescripcion.Text)
txtModelo.Text = fxSysCleanTxtInject(txtModelo.Text)
txtSerie.Text = fxSysCleanTxtInject(txtSerie.Text)
txtMarca.Text = fxSysCleanTxtInject(txtMarca.Text)



If Len(Trim(txtDescripcion)) < 10 Then vMensaje = vMensaje & vbCrLf & "- La descripción no es válida"
If Not IsNumeric(txtAvaluo.Text) Then
   vMensaje = vMensaje & vbCrLf & "- El dato del Avalúo es erroneo!"
End If


If Len(vMensaje) > 0 Then
  fxVerificaDatos = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuarda()

On Error GoTo vError
        
 strSQL = "exec spCrd_Operacion_Prenda_Registro " & mPrendaId & ", " & Operacion.Operacion & ", '" & Operacion.Codigo & "', '" _
        & cboTipo.ItemData(cboTipo.ListIndex) & "', " & CCur(txtAvaluo.Text) & ", '" & txtDescripcion.Text & "', '" & txtModelo.Text & "', '" & txtSerie.Text _
        & "', '" & txtMarca.Text & "', '" & glogon.Usuario & "', 'A'"

 Call ConectionExecute(strSQL)
 
 MsgBox "Prenda registrada satisfactoriamente...", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iRespuesta As Integer

Select Case Button.Key
  Case "insertar", "nuevo"
   vEdita = 0
   Call sbLimpia
    
    With tlbPrincipal
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = True
       .Buttons(5).Enabled = True
    End With
'    fra.Enabled = True
    cboTipo.SetFocus
  
  Case "editar", "modificar"
   vEdita = 1
    With tlbPrincipal
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = True
       .Buttons(5).Enabled = True
    End With
'    fra.Enabled = True
    cboTipo.SetFocus
  
  Case "borrar"
   
   If mPrendaId > 0 Then
    iRespuesta = MsgBox("Esta seguro que desea eliminar esta prenda?", vbYesNo)
    
    If iRespuesta = vbYes Then
       strSQL = "exec spCrd_Operacion_Prenda_Registro " & mPrendaId & ", " & Operacion.Operacion & ", '" & Operacion.Codigo & "', '" _
        & cboTipo.ItemData(cboTipo.ListIndex) & "', " & CCur(txtAvaluo.Text) & ",'" & txtDescripcion.Text & "', '" & txtModelo.Text & "', '" & txtSerie.Text _
        & "', '" & txtMarca.Text & "', '" & glogon.Usuario & "', 'E'"


      Call ConectionExecute(strSQL)
      Call sbPrendas_Load
      Call sbLimpia
    Else
      Call sbLimpia
    End If
    
    With tlbPrincipal
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = False
       .Buttons(5).Enabled = False
     End With
    
   End If
  
  Case "salvar", "guardar"
    If fxVerificaDatos Then
      Call sbGuarda
      
      Call sbPrendas_Load
      
      With tlbPrincipal
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        .Buttons(4).Enabled = False
        .Buttons(5).Enabled = False
      End With
      
      Call sbLimpia
    
    Else
      MsgBox "Información Ingresada es Incorrecta por favor verifique...", vbInformation
    End If
  
  Case "deshacer"
    Call sbLimpia
    
    With tlbPrincipal
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = False
       .Buttons(5).Enabled = False
    End With
  
  Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
        
  Case "salir", "cerrar"
    Unload Me

End Select

End Sub






