VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmTES_Reclasificacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reclasificacion de Solicitudes y Documentos"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1010
   Icon            =   "frmTES_Reclasificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   9255
      _Version        =   1310723
      _ExtentX        =   16325
      _ExtentY        =   4048
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtCuenta 
         Height          =   315
         Left            =   1680
         TabIndex        =   22
         Top             =   360
         Width           =   2655
         _Version        =   1310723
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   315
         Left            =   6120
         TabIndex        =   23
         Top             =   360
         Width           =   2655
         _Version        =   1310723
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   315
         Left            =   6120
         TabIndex        =   24
         Top             =   1560
         Width           =   2655
         _Version        =   1310723
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtFecha 
         Height          =   315
         Left            =   6120
         TabIndex        =   25
         Top             =   1920
         Width           =   2655
         _Version        =   1310723
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   315
         Left            =   1680
         TabIndex        =   26
         Top             =   1560
         Width           =   2655
         _Version        =   1310723
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtTipo 
         Height          =   315
         Left            =   1680
         TabIndex        =   27
         Top             =   1920
         Width           =   2655
         _Version        =   1310723
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtBanco 
         Height          =   315
         Left            =   2400
         TabIndex        =   28
         Top             =   720
         Width           =   6375
         _Version        =   1310723
         _ExtentX        =   11245
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBeneficiario 
         Height          =   315
         Left            =   1680
         TabIndex        =   29
         Top             =   1200
         Width           =   7095
         _Version        =   1310723
         _ExtentX        =   12515
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtID_Banco 
         Height          =   315
         Left            =   1680
         TabIndex        =   30
         Top             =   720
         Width           =   735
         _Version        =   1310723
         _ExtentX        =   1296
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   21
         Top             =   1920
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   6
         Left            =   5160
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   19
         Top             =   1920
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Emite"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Código"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Beneficiario"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   16
         Top             =   360
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Banco/Cta"
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
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cuenta"
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
   End
   Begin XtremeSuiteControls.TabControl ssTab 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   9375
      _Version        =   1310723
      _ExtentX        =   16536
      _ExtentY        =   3201
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
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "Solicitud"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "fraSolicitud"
      Item(1).Caption =   "Documento Emitido"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "fraDocumento"
      Item(2).Caption =   "Cuenta Bancaria"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "cmdCambiaBanco"
      Item(2).Control(1)=   "cboBancoDestino"
      Item(2).Control(2)=   "Label7"
      Begin XtremeSuiteControls.GroupBox fraDocumento 
         Height          =   1335
         Left            =   -69880
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1310723
         _ExtentX        =   15901
         _ExtentY        =   2355
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton cmdCambiaDocumentos 
            Height          =   765
            Left            =   7320
            TabIndex        =   10
            Top             =   480
            Width           =   1695
            _Version        =   1310723
            _ExtentX        =   2990
            _ExtentY        =   1349
            _StockProps     =   79
            Caption         =   "Cambiar"
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
            Picture         =   "frmTES_Reclasificacion.frx":6852
         End
         Begin XtremeSuiteControls.FlatEdit txtDocumentoActual 
            Height          =   315
            Left            =   2520
            TabIndex        =   11
            Top             =   360
            Width           =   2655
            _Version        =   1310723
            _ExtentX        =   4683
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
            Enabled         =   0   'False
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDocumentoNuevo 
            Height          =   315
            Left            =   2520
            TabIndex        =   12
            Top             =   720
            Width           =   2655
            _Version        =   1310723
            _ExtentX        =   4683
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
         Begin VB.Label Label5 
            Caption         =   "Documento Nuevo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   9
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "Documento Actual"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
      End
      Begin XtremeSuiteControls.GroupBox fraSolicitud 
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   9135
         _Version        =   1310723
         _ExtentX        =   16113
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Cambios: "
         BackColor       =   -2147483633
         Appearance      =   6
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton cmdCambiaSolicitud 
            Height          =   765
            Left            =   7320
            TabIndex        =   6
            Top             =   480
            Width           =   1695
            _Version        =   1310723
            _ExtentX        =   2990
            _ExtentY        =   1349
            _StockProps     =   79
            Caption         =   "Cambiar"
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
            Picture         =   "frmTES_Reclasificacion.frx":702A
         End
         Begin XtremeSuiteControls.ComboBox cboBancos 
            Height          =   330
            Left            =   1680
            TabIndex        =   31
            Top             =   360
            Width           =   5055
            _Version        =   1310723
            _ExtentX        =   8916
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
            Height          =   330
            Left            =   1680
            TabIndex        =   32
            Top             =   720
            Width           =   5055
            _Version        =   1310723
            _ExtentX        =   8916
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   1452
         End
         Begin VB.Label Label4 
            Caption         =   "Cuenta"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   852
         End
      End
      Begin XtremeSuiteControls.PushButton cmdCambiaBanco 
         Height          =   765
         Left            =   -62560
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   1349
         _StockProps     =   79
         Caption         =   "Cambiar"
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
         Picture         =   "frmTES_Reclasificacion.frx":7802
      End
      Begin XtremeSuiteControls.ComboBox cboBancoDestino 
         Height          =   330
         Left            =   -69520
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   5055
         _Version        =   1310723
         _ExtentX        =   8916
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label7 
         Caption         =   "Cuenta Destino"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -69520
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1812
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Index           =   3
      Left            =   120
      TabIndex        =   34
      Top             =   3240
      Width           =   9375
      _Version        =   1310723
      _ExtentX        =   16536
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Notas del cambio: "
      ForeColor       =   4210752
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   975
         Left            =   1680
         TabIndex        =   35
         Top             =   360
         Width           =   7095
         _Version        =   1310723
         _ExtentX        =   12515
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
   End
   Begin XtremeSuiteControls.FlatEdit txtNumeroSolicitud 
      Height          =   435
      Left            =   6240
      TabIndex        =   36
      Top             =   120
      Width           =   2175
      _Version        =   1310723
      _ExtentX        =   3831
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Solicitud"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   37
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmTES_Reclasificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbConsulta()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Despliega en pantalla los datos pertinentes a la solicitud digitada por el
'               usuario.
'REFERENCIAS:   LimpiaObjetos - (Limpia los objetos que muestran informacion pertinente a
'               la solicitud por reclasificar)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, rs As New ADODB.Recordset
Dim strEstado As String

Me.MousePointer = vbHourglass


On Error GoTo vError

If Trim(txtNumeroSolicitud) = "" Then
   Me.MousePointer = vbDefault
   MsgBox "Indique el Numero de la Solicitud", vbExclamation
   Exit Sub
End If

strSQL = "Select T.*,B.Descripcion as 'BancoDesc',B.CtaConta as 'BancoCta', Td.descripcion as 'TipoDesc'" _
       & " from Tes_Transacciones T " _
       & " inner join Tes_Bancos B on T.id_Banco = B.id_Banco" _
       & " inner join tes_tipos_doc Td on T.Tipo = Td.Tipo" _
       & " Where T.Nsolicitud=" & Trim(txtNumeroSolicitud)
Call OpenRecordSet(rs, strSQL)
       
ssTab.Item(0).Enabled = True
ssTab.Item(0).Selected = True
ssTab.Item(1).Enabled = False
ssTab.Item(2).Enabled = False
        
If Not rs.EOF And Not rs.BOF Then
   Select Case rs!Estado
      Case "A"
        txtEstado = "Anulado"
        ssTab.Item(1).Selected = True
        ssTab.Item(0).Enabled = False
        ssTab.Item(1).Enabled = True
        ssTab.Item(2).Enabled = True
      Case "T"
        txtEstado = "Transferido"
        ssTab.Item(1).Selected = True
        ssTab.Item(0).Enabled = False
        ssTab.Item(1).Enabled = True
        ssTab.Item(2).Enabled = True
      Case "I"
        txtEstado = "Impreso"
        ssTab.Item(1).Selected = True
        ssTab.Item(0).Enabled = False
        ssTab.Item(1).Enabled = True
        ssTab.Item(2).Enabled = True
      Case "P"
        txtEstado = "Pendiente"
        ssTab.Item(0).Enabled = True
        ssTab.Item(1).Enabled = False
        ssTab.Item(2).Enabled = False
   End Select
   
   txtBeneficiario = rs!Beneficiario
   txtCodigo = rs!Codigo
   txtMonto = Format(rs!Monto, "Standard")
   
   txtTipo.Text = rs!TipoDesc
   txtTipo.Tag = rs!Tipo
   
   txtFecha = Format(rs!fecha_solicitud, "dd/mm/yyyy")
   txtID_Banco = rs!ID_BANCO
   txtBanco = Trim(rs!BancoDesc)
   txtCuenta.Text = Trim(rs!Cta_Ahorros)
   txtDocumentoActual = IIf(IsNull(rs!nDocumento), "", Trim(rs!nDocumento))

Else
    Call LimpiaObjetos
    MsgBox "Numero de Solicitud No existe", vbExclamation, "Atención"
    ssTab.Item(1).Enabled = False
'    ssTab.Item(1).Enabled = False
End If

Me.MousePointer = vbDefault

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Sub LimpiaObjetos()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Limpia los objetos que muestran informacion pertinente a la solicitud por
'               reclasificar.
'REFERENCIAS:   Ninguna.
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

txtNumeroSolicitud = ""
txtEstado = ""
txtBeneficiario = ""
txtCodigo = ""
txtMonto = ""
txtTipo = ""
txtFecha = ""
txtID_Banco = ""
txtBanco = ""
txtCuenta = ""
End Sub

Private Function fxCuentaBanco(vBanco As Integer) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select ctaconta as Cuenta from Tes_Bancos where id_banco = " & vBanco
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 fxCuentaBanco = Trim(rs!Cuenta)
Else
 fxCuentaBanco = ""
End If
rs.Close

End Function


Private Sub cboBancos_Click()
If vPaso Then
 Call sbTesTiposDocsCargaCboAcceso(cboTipoDocumento, glogon.Usuario, cboBancos.ItemData(cboBancos.ListIndex))
End If
End Sub


Private Sub cmdCambiaBanco_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pNota As String

On Error GoTo vError

If Val(txtID_Banco.Text) = cboBancoDestino.ItemData(cboBancoDestino.ListIndex) Then Exit Sub

strSQL = "select estado_asiento from Tes_Transacciones where nsolicitud = " & txtNumeroSolicitud.Text
Call OpenRecordSet(rs, strSQL)
If rs!estado_asiento = "G" Then
  rs.Close
  MsgBox "El asiento de esta solicitud ya fue generado, no se puede reclasificar...", vbInformation
  Exit Sub
End If
rs.Close

If Len(txtNotas.Text) = 0 Then
   MsgBox "Identifique una Nota válida para realizar el movimiento!", vbExclamation
   Exit Sub
End If



On Error GoTo vError

Me.MousePointer = vbHourglass


pNota = Mid(fxSysCleanTxtInject(txtNotas.Text), 1, 500)


strSQL = "exec spTes_Reclasificacion " & txtNumeroSolicitud & ", " & cboBancoDestino.ItemData(cboBancoDestino.ListIndex) _
        & ", '" & txtTipo.Tag & "','" & glogon.Usuario & "','" & pNota & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Reclasifica Solicitud " & Trim(txtNumeroSolicitud))

Me.MousePointer = vbDefault

Call LimpiaObjetos

MsgBox "Cambio de Banco Realizado Satisfactoriamente...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub cmdCambiaDocumentos_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Modifica la solicitud en cuanto al # de Documento.
'REFERENCIAS:   Bitacora - (Registra movimientos sobre la Base de Datos)
'               LimpiaObjetos - (Limpia los objetos que muestran informacion pertinente a
'               la solicitud por reclasificar)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, rs As New ADODB.Recordset, vTipo As String

If Trim(txtDocumentoNuevo) = "" Then
   MsgBox "Escriba el Numero De Documento", vbExclamation, "No Se Puede Reclasificar"
   Exit Sub
End If

If Trim(txtNumeroSolicitud) = "" Then
   MsgBox "Suministre Numero de Solicitud", vbExclamation, "No Se Puede Reclasificar"
   Exit Sub
End If

If Len(txtNotas.Text) = 0 Then
   MsgBox "Identifique una Nota válida para realizar el movimiento!", vbExclamation
   Exit Sub
End If


vTipo = txtTipo.Tag

strSQL = "Select Nsolicitud from Tes_Transacciones where id_banco= " & Trim(txtID_Banco) _
       & " And Tipo='" & vTipo & "' and Ndocumento='" & Trim(Me.txtDocumentoNuevo) & "'"
With rs
 .Open strSQL, glogon.Conection, adOpenStatic
   If .EOF = False Then
     MsgBox "# Documento Ya Existe", vbExclamation, "No Se Puede Reclasificar"
     Exit Sub
   End If
 .Close
End With

On Error GoTo vError
Me.MousePointer = vbHourglass
'glogon.Conection.BeginTrans

strSQL = "Update Tes_Transacciones Set Ndocumento='" & Trim(txtDocumentoNuevo) _
       & "' Where NSolicitud=" & Trim(txtNumeroSolicitud)
Call ConectionExecute(strSQL)


strSQL = "Cambio N.Documento de " & Trim(txtDocumentoActual.Text) & " a " & Trim(txtDocumentoNuevo.Text)
Call sbTesBitacoraEspecial(txtNumeroSolicitud, "09", strSQL)

Call Bitacora("Modifica", "Modifica Documento a Solicitud " & Trim(txtNumeroSolicitud))

'glogon.Conection.CommitTrans

Call LimpiaObjetos

MsgBox "Reclasificacion Realizada", vbExclamation

fraDocumento.Enabled = False
fraSolicitud.Enabled = False

Me.MousePointer = vbDefault

Exit Sub
vError:
 '   glogon.Conection.RollbackTrans
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdCambiaSolicitud_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Reclasifica la solicitud en cuanto al Banco y Tipo de Documento. Ademas
'               actualiza para el detalle de la solicitud el # Cuenta del Banco.
'REFERENCIAS:   Bitacora - (Registra movimientos sobre la Base de Datos)
'               LimpiaObjetos - (Limpia los objetos que muestran informacion pertinente a
'               la solicitud por reclasificar)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipo As String, pNota As String

If Not ssTab.Item(0).Enabled Then Exit Sub


If Trim(cboTipoDocumento) = "" Then
   MsgBox "Elija El Tipo De Documento", vbExclamation, "No Se Puede Reclasificar"
   Exit Sub
End If

If Trim(cboBancos) = "" Then
   MsgBox "Elija El Banco", vbExclamation, "No Se Puede Reclasificar"
   Exit Sub
End If

If Trim(txtNumeroSolicitud) = "" Then
   MsgBox "Suministre Numero de Solicitud", vbExclamation, "No Se Puede Reclasificar"
   Exit Sub
End If

If Len(txtNotas.Text) = 0 Then
   MsgBox "Identifique una Nota válida para realizar el movimiento!", vbExclamation
   Exit Sub
End If


On Error GoTo vError

Me.MousePointer = vbHourglass

vTipo = cboTipoDocumento.ItemData(cboTipoDocumento.ListIndex)

pNota = Mid(fxSysCleanTxtInject(txtNotas.Text), 1, 500)


strSQL = "exec spTes_Reclasificacion " & txtNumeroSolicitud.Text & ", " & cboBancos.ItemData(cboBancos.ListIndex) _
        & ", '" & vTipo & "','" & glogon.Usuario & "','" & pNota & "'"
Call ConectionExecute(strSQL)



Call Bitacora("Modifica", "Reclasifica Solicitud " & Trim(txtNumeroSolicitud))

Me.MousePointer = vbDefault

Call LimpiaObjetos

MsgBox "Reclasificacion Realizada", vbExclamation

fraDocumento.Enabled = False
fraSolicitud.Enabled = False


Exit Sub
vError:
 '   glogon.Conection.RollbackTrans
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Carga el combo de Tes_Bancos y el combo de tipos de Documentos.
'REFERENCIAS:   CentrarFrm - (Centra el formulario dentro del formulario MDI)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String

vModulo = 9

vPaso = False
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

strSQL = "select id_banco as Idx,rtrim(descripcion) as ItmX from Tes_Bancos where estado = 'A'"
Call sbCbo_Llena_New(cboBancos, strSQL, False, True)

vPaso = True

Call sbTesTiposDocsCargaCboAcceso(cboTipoDocumento, glogon.Usuario, cboBancos.ItemData(cboBancos.ListIndex))
Call sbCbo_Llena_New(cboBancoDestino, strSQL, False, True)



ssTab.Item(0).Selected = True
ssTab.Item(1).Enabled = False
ssTab.Item(2).Enabled = False


Call Formularios(Me)
Call RefrescaTags(Me)

If IsNumeric(GLOBALES.gTag) Then
   Call sbGReclasifica(CLng(GLOBALES.gTag))
End If

End Sub

Public Sub sbGReclasifica(vSolicitud As Long)

txtNumeroSolicitud = vSolicitud
Call sbConsulta

End Sub


Private Sub txtNumeroSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsulta
End Sub
