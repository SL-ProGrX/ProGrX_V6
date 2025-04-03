VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmTES_TransferenciaReversa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reversión de Transferencia"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12330
   Icon            =   "frmTES_TransferenciaReversa.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   12330
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7692
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12132
      _Version        =   1441793
      _ExtentX        =   21399
      _ExtentY        =   13568
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
      Item(0).Caption =   "Reversión"
      Item(0).ControlCount=   13
      Item(0).Control(0)=   "txtNdocumento"
      Item(0).Control(1)=   "cbo"
      Item(0).Control(2)=   "lsw"
      Item(0).Control(3)=   "Label1(0)"
      Item(0).Control(4)=   "Label1(1)"
      Item(0).Control(5)=   "btnBuscar"
      Item(0).Control(6)=   "cboTipo"
      Item(0).Control(7)=   "gbReversion"
      Item(0).Control(8)=   "Label1(2)"
      Item(0).Control(9)=   "gbFiltros"
      Item(0).Control(10)=   "btnExport"
      Item(0).Control(11)=   "cboPlan"
      Item(0).Control(12)=   "Label1(6)"
      Item(1).Caption =   "Consulta"
      Item(1).ControlCount=   15
      Item(1).Control(0)=   "cboBancos"
      Item(1).Control(1)=   "dtpInicio"
      Item(1).Control(2)=   "dtpCorte"
      Item(1).Control(3)=   "lswConsulta"
      Item(1).Control(4)=   "lswDetalle"
      Item(1).Control(5)=   "Label1(3)"
      Item(1).Control(6)=   "Label1(4)"
      Item(1).Control(7)=   "Label2"
      Item(1).Control(8)=   "Label4"
      Item(1).Control(9)=   "Label1(5)"
      Item(1).Control(10)=   "Label1(7)"
      Item(1).Control(11)=   "txtCasosDet"
      Item(1).Control(12)=   "txtMontoDet"
      Item(1).Control(13)=   "btnConsulta(0)"
      Item(1).Control(14)=   "btnConsulta(1)"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3135
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20976
         _ExtentY        =   5524
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
      Begin XtremeSuiteControls.ListView lswDetalle 
         Height          =   3255
         Left            =   -69880
         TabIndex        =   24
         Top             =   4320
         Visible         =   0   'False
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20981
         _ExtentY        =   5741
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
      Begin XtremeSuiteControls.ListView lswConsulta 
         Height          =   2055
         Left            =   -69880
         TabIndex        =   23
         Top             =   1680
         Visible         =   0   'False
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20981
         _ExtentY        =   3625
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
      Begin XtremeSuiteControls.GroupBox gbFiltros 
         Height          =   855
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20981
         _ExtentY        =   1508
         _StockProps     =   79
         BackColor       =   -2147483643
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtDocTran 
            Height          =   315
            Left            =   7560
            TabIndex        =   41
            Top             =   360
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
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
         Begin XtremeSuiteControls.FlatEdit txtCodigo 
            Height          =   315
            Left            =   0
            TabIndex        =   37
            Top             =   360
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtBeneficiario 
            Height          =   315
            Left            =   1920
            TabIndex        =   39
            Top             =   360
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtIBAN 
            Height          =   315
            Left            =   9360
            TabIndex        =   43
            Top             =   360
            Width           =   2535
            _Version        =   1441793
            _ExtentX        =   4471
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
         Begin XtremeSuiteControls.Label Label6 
            Height          =   255
            Index           =   3
            Left            =   9360
            TabIndex        =   44
            Top             =   120
            Width           =   2535
            _Version        =   1441793
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No.Cuenta  / IBAN"
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   255
            Index           =   2
            Left            =   7560
            TabIndex        =   42
            Top             =   120
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No.Documento"
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   40
            Top             =   120
            Width           =   5655
            _Version        =   1441793
            _ExtentX        =   9975
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Beneficiario"
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   38
            Top             =   120
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Código"
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin XtremeSuiteControls.GroupBox gbReversion 
         Height          =   2292
         Left            =   120
         TabIndex        =   2
         Top             =   5400
         Width           =   11772
         _Version        =   1441793
         _ExtentX        =   20764
         _ExtentY        =   4043
         _StockProps     =   79
         Caption         =   "Reversión:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton cmdAplicar 
            Height          =   552
            Left            =   9480
            TabIndex        =   3
            Top             =   1680
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
            _ExtentY        =   974
            _StockProps     =   79
            Caption         =   "&Reversar"
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmTES_TransferenciaReversa.frx":6852
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtObservaciones 
            Height          =   915
            Left            =   4680
            TabIndex        =   17
            Top             =   600
            Width           =   6495
            _Version        =   1441793
            _ExtentX        =   11451
            _ExtentY        =   1609
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtContraseña 
            Height          =   315
            Left            =   4680
            TabIndex        =   19
            Top             =   1560
            Width           =   2055
            _Version        =   1441793
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
            PasswordChar    =   "*"
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCasos 
            Height          =   315
            Left            =   6480
            TabIndex        =   31
            Top             =   240
            Width           =   1095
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   315
            Left            =   8520
            TabIndex        =   32
            Top             =   240
            Width           =   2655
            _Version        =   1441793
            _ExtentX        =   4678
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label lblEnd1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Casos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5760
            TabIndex        =   34
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblEnd2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   7680
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Notas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   18
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Contraseña de Autorizador"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   4
            Top             =   1560
            Width           =   2295
         End
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   435
         Left            =   9840
         TabIndex        =   7
         Top             =   600
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "&Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_TransferenciaReversa.frx":6F6B
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   330
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
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
      Begin XtremeSuiteControls.FlatEdit txtNdocumento 
         Height          =   330
         Left            =   5760
         TabIndex        =   15
         Top             =   720
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboBancos 
         Height          =   312
         Left            =   -68200
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   5892
         _Version        =   1441793
         _ExtentX        =   10398
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -68200
         TabIndex        =   21
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -66880
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.FlatEdit txtCasosDet 
         Height          =   315
         Left            =   -62680
         TabIndex        =   25
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMontoDet 
         Height          =   315
         Left            =   -60640
         TabIndex        =   26
         Top             =   3960
         Visible         =   0   'False
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4678
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   432
         Index           =   0
         Left            =   -65200
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "&Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_TransferenciaReversa.frx":766B
      End
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   432
         Index           =   1
         Left            =   -63880
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "&Reporte"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_TransferenciaReversa.frx":7D6B
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   7800
         TabIndex        =   30
         Top             =   720
         Width           =   1935
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
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   435
         Left            =   11160
         TabIndex        =   45
         ToolTipText     =   "Exportar a Excel"
         Top             =   600
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   767
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_TransferenciaReversa.frx":8472
      End
      Begin XtremeSuiteControls.ComboBox cboPlan 
         Height          =   330
         Left            =   4200
         TabIndex        =   46
         Top             =   720
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
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
      Begin VB.Label Label1 
         Caption         =   "Plan"
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
         Index           =   6
         Left            =   4200
         TabIndex        =   47
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Reversión"
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
         Index           =   2
         Left            =   7800
         TabIndex        =   35
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Detalle Reversión"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   -69880
         TabIndex        =   13
         Top             =   3960
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Reversión"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   -69880
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Casos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -63400
         TabIndex        =   11
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -61600
         TabIndex        =   10
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fechas Reversión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   -69760
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   -69760
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2052
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
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
         Index           =   0
         Left            =   5760
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   156
      Left            =   0
      TabIndex        =   0
      Top             =   9012
      Width           =   12324
      _ExtentX        =   21749
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   372
      Left            =   2640
      TabIndex        =   29
      Top             =   360
      Width           =   6372
      _Version        =   1441793
      _ExtentX        =   11239
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Reversión de Transferencias Electrónicas"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_TransferenciaReversa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, fFechaEmision As Date, mReversionId As Long

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnConsulta_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If vPaso Then Exit Sub
  
    
lswConsulta.ListItems.Clear
lswDetalle.ListItems.Clear
    
Select Case Index
  Case 0 '"Buscar"
        strSQL = "select * from tes_te_reversion where id_banco = " & cboBancos.ItemData(cboBancos.ListIndex) _
              & " and fecha_genera between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
              & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
        Call OpenRecordSet(rs, strSQL)
        
        Do While Not rs.EOF
            Set itmX = lswConsulta.ListItems.Add(, , rs!id_reversion)
                itmX.SubItems(1) = rs!Documento & ""
                itmX.SubItems(2) = rs!autorizado
                itmX.SubItems(3) = rs!user_genera
                itmX.SubItems(4) = Format(rs!fecha_genera, "dd/mm/yyyy")
                itmX.SubItems(5) = rs!observaciones
            rs.MoveNext
        Loop
        rs.Close
        
  Case 1 '"Reporte"
        Call sbBoleta_Reversion(mReversionId)

End Select

End Sub

Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

PrgBar.Visible = True

Call Excel_Exportar_Lsw(lsw, PrgBar)

PrgBar.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cbo_Click()

If vPaso Then Exit Sub

txtCasos.Text = 0
txtMonto.Text = 0
lsw.ListItems.Clear


If cbo.ListCount = 0 Then Exit Sub

Dim strSQL As String

strSQL = "select Bp.COD_PLAN as 'IdX', Bp.COD_PLAN as 'ItmX'" _
       & " from TES_BANCOS B inner join TES_BANCO_PLANES_TE Bp on B.ID_BANCO = Bp.ID_BANCO" _
       & " Where B.ID_BANCO = " & cbo.ItemData(cbo.ListIndex) & " And B.UTILIZA_PLAN = 1" _
       & " order by Bp.COD_PLAN  asc"
Call sbCbo_Llena_New(cboPlan, strSQL, False, True)
If cboPlan.ListCount = 0 Then
   cboPlan.AddItem "Sin Plan"
   cboPlan.ItemData(cboPlan.ListCount - 1) = "-sp-"
   cboPlan.Text = "Sin Plan"
End If


'txtNdocumento.Text = fxTesTipoDocConsec(cbo.ItemData(cbo.ListIndex), "TE", "/")
Call sbNTrasnferencia

vPaso = False
End Sub

Private Sub sbNTrasnferencia()
Dim strSQL As String, rs As New ADODB.Recordset

txtNdocumento = fxTesTipoDocConsec(cbo.ItemData(cbo.ListIndex), "TE", "/", cboPlan.ItemData(cboPlan.ListIndex))

End Sub


Private Sub cboBancos_Click()
If vPaso Then Exit Sub
lswConsulta.ListItems.Clear
lswDetalle.ListItems.Clear
End Sub

Private Sub cboTipo_Click()

If cboTipo.ListCount = 0 Then Exit Sub
If vPaso Then Exit Sub

txtCodigo.Text = ""
txtBeneficiario.Text = ""
txtDocTran.Text = ""
txtIBAN.Text = ""

If cboTipo.ItemData(cboTipo.ListIndex) = "T" Then
    lsw.Checkboxes = False
    gbFiltros.Enabled = False
    
Else
    lsw.Checkboxes = True
    gbFiltros.Enabled = True
End If

Call sbBuscar

End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Long, iConsecutivo As Long
Dim pDias As Long

On Error GoTo vError


If Trim(txtContraseña.Text) = "" Then
   MsgBox "No se puede Autorizar" & vbCrLf & "Suministre La Contraseña De Autorización", vbExclamation
   Exit Sub
End If

If Trim(txtObservaciones.Text) = "" Then
   MsgBox "Digite la Nota para justificar la reversión de esta transferencia", vbExclamation
   Exit Sub
End If

strSQL = "select valor   from TES_PARAMETROS where COD_PARAMETRO = '11'"
Call OpenRecordSet(rs, strSQL)
If IsNumeric(rs!Valor) Then
    pDias = rs!Valor
Else
    pDias = 5
End If


If DateDiff("d", fFechaEmision, fxFechaServidor) > pDias Then
   MsgBox "Esta intentando reversar una transferencia con mas de " & pDias & " días de emisión: " & fFechaEmision
   Exit Sub
End If




'If UCase(txtUSuarioAutoriza) = UCase(glogon.Usuario) Then
'   MsgBox "No puede ser autorizado por el usuarui actual"
'   Call Form_Load
'   Exit Sub
'End If


Me.MousePointer = vbHourglass
 
strSQL = "Select * From Tes_Autorizaciones Where Clave='" _
       & fxTESCifrado(Trim(txtContraseña)) & "' and nombre = '" & glogon.Usuario _
       & "' and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
     Me.MousePointer = vbDefault
     MsgBox "Contraseña Incorrecta, o no Existe Nivel de Autorización", vbExclamation
     txtContraseña.Text = ""
     Exit Sub
End If
rs.Close


strSQL = "select count(*) as Existe from tes_te_reversion where isnull(Tipo,'T') = 'T'" _
       & " and id_Banco = " & cbo.ItemData(cbo.ListIndex) & " and Documento = '" & txtNdocumento.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 1 Then
    Me.MousePointer = vbDefault
    MsgBox "La transferencia No." & txtNdocumento.Text & ", ya fue reversada anteriormente!", vbExclamation
    Exit Sub
End If
rs.Close

strSQL = "exec spTES_TE_Reversion_Main " & cbo.ItemData(cbo.ListIndex) & ", '" & cboTipo.ItemData(cboTipo.ListIndex) _
        & "', '" & txtNdocumento.Text & "', '" & txtObservaciones.Text & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

iConsecutivo = rs!ReversionId


PrgBar.Max = lsw.ListItems.Count + 1
PrgBar.Value = 1

PrgBar.Visible = True

        
With lsw.ListItems
strSQL = ""
For i = 1 To .Count
    
    If cboTipo.ItemData(cboTipo.ListIndex) = "T" Then
        strSQL = strSQL & Space(10) & "EXEC spTES_TE_Reversion_Transaccion " & iConsecutivo & ", " & .Item(i).Text & ", '" & glogon.Usuario & "'"
    
    Else
        If .Item(i).Checked Then
        
            strSQL = strSQL & Space(10) & "EXEC spTES_TE_Reversion_Transaccion " & iConsecutivo & ", " & .Item(i).Text & ", '" & glogon.Usuario & "'"
        
        End If
    End If
    
    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If
   
   PrgBar.Value = PrgBar.Value + 1
   
Next i
    
End With
    
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If
  

PrgBar.Visible = False

Call Bitacora("Aplica", "Reversion Transferencia = " & txtNdocumento & " Id.Cuenta:" & cbo.ItemData(cbo.ListIndex) & ", Tipo: " & cboTipo.Text)

Call sbBoleta_Reversion(iConsecutivo)

Call sbBuscar

Me.MousePointer = vbDefault


Exit Sub

vError:
  glogon.Conection.RollbackTrans
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
vModulo = 9

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

vPaso = True

With lsw.ColumnHeaders
  .Clear
  .Add , , "No. Solicitud", 1400
  .Add , , "Código", 1400
  .Add , , "Beneficiario", 4000
  .Add , , "Monto", 1600, vbRightJustify
  .Add , , "Fecha", 1200, vbCenter
  .Add , , "Cuenta", 2200
  .Add , , "No.Documento", 2200
End With

With lswDetalle.ColumnHeaders
  .Clear
  .Add , , "No. Solicitud", 1400
  .Add , , "Código", 1400
  .Add , , "Beneficiario", 4000
  .Add , , "Monto", 1600, vbRightJustify
  .Add , , "Fecha", 1200, vbCenter
  .Add , , "Cuenta", 2200
  .Add , , "No.Documento", 2200
  .Add , , "Tipo Apl.", 2200, vbCenter
  .Add , , "Id Tesoreria New", 2200, vbCenter
  
End With


With lswConsulta.ColumnHeaders
  .Clear
  .Add , , "No. Reversión", 1600
  .Add , , "No. Documento", 1600
  .Add , , "Autorizado por", 2000
  .Add , , "Aplicada por", 2000
  .Add , , "Fecha", 2100
  .Add , , "Notas", 3200
End With


cboTipo.Clear
cboTipo.AddItem "TOTAL"
cboTipo.ItemData(cboTipo.ListCount - 1) = "T"
cboTipo.AddItem "PARCIAL"
cboTipo.ItemData(cboTipo.ListCount - 1) = "P"
cboTipo.Text = "TOTAL"

lsw.ListItems.Clear


txtNdocumento.Text = ""
txtMonto.Text = ""
txtCasos.Text = ""

tcMain.Item(0).Selected = True

Call sbTesBancoCargaCboAccesoGestion(cbo, glogon.Usuario, "Autoriza")

vPaso = False

Call cbo_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Long, curMonto As Currency, itmX As ListViewItem

If vPaso Then Exit Sub

    On Error GoTo vError
    
    
    If txtNdocumento = "" Then
        MsgBox "No ha seleccionado la transferencia a Reversar", vbExclamation
        Exit Sub
    End If
    
    
    Me.MousePointer = vbHourglass
    
    
    'Revisa Filtros
    txtCodigo.Text = fxSysCleanTxtInject(txtCodigo.Text)
    txtBeneficiario.Text = fxSysCleanTxtInject(txtBeneficiario.Text)
    txtDocTran.Text = fxSysCleanTxtInject(txtDocTran.Text)
    txtIBAN.Text = fxSysCleanTxtInject(txtIBAN.Text)
    txtNdocumento.Text = fxSysCleanTxtInject(txtNdocumento.Text)
        
        
    lsw.ListItems.Clear
    i = 0
    curMonto = 0
    
    strSQL = "select nsolicitud,codigo,beneficiario,monto,fecha_emision,cta_ahorros,Ndocumento" _
           & " from Tes_Transacciones " _
           & " where documento_base = '" & txtNdocumento & "' and id_banco = " & cbo.ItemData(cbo.ListIndex)
    
    If Len(Trim(txtCodigo.Text)) > 0 Then
          strSQL = strSQL & " and Codigo like '%" & txtCodigo.Text & "%'"
    End If
    
    If Len(Trim(txtBeneficiario.Text)) > 0 Then
          strSQL = strSQL & " and Beneficiario like '%" & txtBeneficiario.Text & "%'"
    End If
    
    If Len(Trim(txtDocTran.Text)) > 0 Then
          strSQL = strSQL & " and NDocumento like '%" & txtDocTran.Text & "%'"
    End If
    
    If Len(Trim(txtIBAN.Text)) > 0 Then
          strSQL = strSQL & " and cta_ahorros like '%" & txtIBAN.Text & "%'"
    End If
    
    If cboPlan.ItemData(cboPlan.ListIndex) <> "-sp-" Then
          strSQL = strSQL & " and isnull(cod_Plan,'-sp-') = '" & cboPlan.ItemData(cboPlan.ListIndex) & "'"
        
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
        Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = rs!Beneficiario
        itmX.SubItems(3) = Format(rs!Monto, "Standard")
        itmX.SubItems(4) = Format(rs!Fecha_Emision, "yyyy/mm/dd hh:mm:ss")
        itmX.SubItems(5) = rs!Cta_Ahorros
        itmX.SubItems(6) = rs!nDocumento & ""
        
        fFechaEmision = rs!Fecha_Emision
        
        
        If lsw.Checkboxes = False Then
            curMonto = curMonto + rs!Monto
            i = i + 1
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    
    txtCasos.Text = Format(i, "###,###,###,##0")
    txtMonto.Text = Format(curMonto, "Standard")
    
    Me.MousePointer = vbDefault
    Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub


Private Sub sbBoleta_Reversion(vConsec As Long)
Dim strSQL As String

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Banking"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(3) = "Usuario='" & glogon.Usuario & "'"

    .ReportFileName = SIFGlobal.fxPathReportes("Banking_TransferenciaReversion.rpt")
    
    strSQL = "{Tes_Te_Reversion.id_reversion} = " & vConsec
    .SelectionFormula = strSQL
    .PrintReport

End With

End Sub


Private Sub Form_Resize()
On Error Resume Next


tcMain.Width = Me.Width - 350
tcMain.Height = Me.Height - (tcMain.top + 650)

lsw.Width = tcMain.Width - 100
gbFiltros.Width = lsw.Width
gbReversion.WhatsThisHelpID = lsw.Width

lsw.Height = tcMain.Height - (lsw.top + gbReversion.Height + 200)
gbReversion.top = lsw.top + lsw.Height + 100

lswConsulta.Width = lsw.Width
lswDetalle.Width = lsw.Width

lswDetalle.Height = tcMain.Height - (lswDetalle.top + 200)


End Sub




Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim i As Long, curMonto As Currency

 i = CLng(txtCasos.Text)
 curMonto = CCur(txtMonto.Text)
 
 If Item.Checked Then
      i = i + 1
     curMonto = curMonto + CCur(Item.SubItems(3))
      
 Else
      i = i - 1
     curMonto = curMonto - CCur(Item.SubItems(3))
 End If
    
    txtCasos.Text = Format(i, "###,###,###,##0")
    txtMonto.Text = Format(curMonto, "Standard")

End Sub


Private Sub lswConsulta_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Long, curMonto As Currency, itmX As ListViewItem

On Error GoTo vError

    Me.MousePointer = vbHourglass
    
    lswDetalle.ListItems.Clear
    
    i = 0
    curMonto = 0
    
    mReversionId = Item.Text
    
    strSQL = "select * from vTes_TE_Reversion_Det where id_reversion = " & lswConsulta.SelectedItem.Text
    
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
     Set itmX = lswDetalle.ListItems.Add(, , rs!NSolicitud)
         itmX.SubItems(1) = rs!Cedula
         itmX.SubItems(2) = rs!Nombre
         itmX.SubItems(3) = Format(rs!Monto, "Standard")
         itmX.SubItems(4) = Format(rs!Fecha_Emision & "", "dd/MM/yyyy")
         itmX.SubItems(5) = rs!Cta_Ahorros & ""
         itmX.SubItems(6) = rs!nDocumento & ""
         
         itmX.SubItems(7) = rs!ESTADO_TRANSAC_DESC
         itmX.SubItems(8) = rs!TESORERIA_ID_NEW & ""
         
         
         curMonto = curMonto + rs!Monto
         i = i + 1
      
         
     rs.MoveNext
    Loop
    
    rs.Close

    txtCasosDet.Text = Format(i, "###,###,###,##0")
    txtMontoDet.Text = Format(curMonto, "Standard")
    
    Me.MousePointer = vbDefault
    Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 1 Then
  vPaso = True
     Call sbTesBancoCargaCboAccesoGestion(cboBancos, glogon.Usuario, "Autoriza")
  vPaso = False
   
  dtpCorte.Value = fxFechaServidor
  dtpInicio.Value = dtpCorte.Value
  Call cboBancos_Click
End If
End Sub



Private Sub txtBeneficiario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbBuscar
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbBuscar
End Sub


Private Sub txtDocTran_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbBuscar
End Sub


Private Sub txtIBAN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbBuscar
End Sub



Private Sub txtNdocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbBuscar

End Sub
