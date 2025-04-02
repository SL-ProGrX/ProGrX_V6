VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCC_ReportesAlCorte 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes Operativos de Auxiliares al Corte"
   ClientHeight    =   8295
   ClientLeft      =   210
   ClientTop       =   510
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   9420
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   1812
      Left            =   0
      TabIndex        =   59
      Top             =   2880
      Width           =   9372
      _Version        =   1441793
      _ExtentX        =   16531
      _ExtentY        =   3196
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   252
      Index           =   0
      Left            =   1200
      TabIndex        =   53
      Top             =   2160
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Afiliación"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3372
      Left            =   0
      TabIndex        =   5
      Top             =   4800
      Width           =   9372
      _Version        =   1441793
      _ExtentX        =   16531
      _ExtentY        =   5948
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
      ItemCount       =   4
      Item(0).Caption =   "General"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "Label1(11)"
      Item(0).Control(1)=   "Label1(12)"
      Item(0).Control(2)=   "Label1(13)"
      Item(0).Control(3)=   "cboTipo"
      Item(0).Control(4)=   "cboEstado"
      Item(0).Control(5)=   "cboZonas"
      Item(0).Control(6)=   "cboProfesion"
      Item(0).Control(7)=   "cboSector"
      Item(0).Control(8)=   "Label1(7)"
      Item(0).Control(9)=   "Label1(8)"
      Item(0).Control(10)=   "Label1(15)"
      Item(0).Control(11)=   "Label1(16)"
      Item(0).Control(12)=   "cboCondicion"
      Item(0).Control(13)=   "cboSexo"
      Item(0).Control(14)=   "cboEstadoCivil"
      Item(0).Control(15)=   "btnInforme"
      Item(1).Caption =   "Crédito/Cobro"
      Item(1).ControlCount=   14
      Item(1).Control(0)=   "Label9(1)"
      Item(1).Control(1)=   "Label10(1)"
      Item(1).Control(2)=   "chkLineas"
      Item(1).Control(3)=   "txtCodigo"
      Item(1).Control(4)=   "cboGarantia"
      Item(1).Control(5)=   "cboCartera"
      Item(1).Control(6)=   "Label9(2)"
      Item(1).Control(7)=   "Label10(2)"
      Item(1).Control(8)=   "cboDestino"
      Item(1).Control(9)=   "cboRecurso"
      Item(1).Control(10)=   "Label9(3)"
      Item(1).Control(11)=   "cboOficina"
      Item(1).Control(12)=   "txtDescripcion"
      Item(1).Control(13)=   "Label9(4)"
      Item(2).Caption =   "Adicionales"
      Item(2).ControlCount=   17
      Item(2).Control(0)=   "chkDistritos"
      Item(2).Control(1)=   "chkCantones"
      Item(2).Control(2)=   "chkProvincias"
      Item(2).Control(3)=   "chkDepartamento"
      Item(2).Control(4)=   "chkSeccion"
      Item(2).Control(5)=   "lblSeccion"
      Item(2).Control(6)=   "lblDepartamento"
      Item(2).Control(7)=   "Label9(0)"
      Item(2).Control(8)=   "Label10(0)"
      Item(2).Control(9)=   "Label18(0)"
      Item(2).Control(10)=   "txtDeptCodigo"
      Item(2).Control(11)=   "txtDeptDesc"
      Item(2).Control(12)=   "txtSecCodigo"
      Item(2).Control(13)=   "txtSecDesc"
      Item(2).Control(14)=   "cboProvincia"
      Item(2).Control(15)=   "cboCanton"
      Item(2).Control(16)=   "cboDistrito"
      Item(3).Caption =   "Cubo"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "lblStatus"
      Item(3).Control(1)=   "Image3"
      Item(3).Control(2)=   "Label2"
      Item(3).Control(3)=   "btnCubo_Procesa"
      Begin XtremeSuiteControls.CheckBox chkProvincias 
         Height          =   252
         Left            =   -64720
         TabIndex        =   44
         Top             =   600
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   7200
         TabIndex        =   9
         Top             =   2280
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   2280
         TabIndex        =   10
         Top             =   600
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8281
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboZonas 
         Height          =   312
         Left            =   2280
         TabIndex        =   11
         Top             =   960
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8281
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboProfesion 
         Height          =   312
         Left            =   2280
         TabIndex        =   12
         Top             =   1440
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8281
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboSector 
         Height          =   312
         Left            =   2280
         TabIndex        =   13
         Top             =   1800
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8281
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboCondicion 
         Height          =   312
         Left            =   5400
         TabIndex        =   18
         Top             =   2760
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2778
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
      End
      Begin XtremeSuiteControls.ComboBox cboSexo 
         Height          =   312
         Left            =   2280
         TabIndex        =   19
         Top             =   2760
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2778
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
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoCivil 
         Height          =   312
         Left            =   3840
         TabIndex        =   20
         Top             =   2760
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2778
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
      End
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   612
         Left            =   7200
         TabIndex        =   21
         Top             =   2640
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Informe"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCC_ReportesAlCorte.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtDeptCodigo 
         Height          =   312
         Left            =   -68320
         TabIndex        =   27
         Top             =   2280
         Visible         =   0   'False
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
      Begin XtremeSuiteControls.FlatEdit txtDeptDesc 
         Height          =   312
         Left            =   -67600
         TabIndex        =   28
         Top             =   2280
         Visible         =   0   'False
         Width           =   4932
         _Version        =   1441793
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.FlatEdit txtSecCodigo 
         Height          =   312
         Left            =   -68320
         TabIndex        =   29
         Top             =   2640
         Visible         =   0   'False
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
      Begin XtremeSuiteControls.FlatEdit txtSecDesc 
         Height          =   312
         Left            =   -67600
         TabIndex        =   30
         Top             =   2640
         Visible         =   0   'False
         Width           =   4932
         _Version        =   1441793
         _ExtentX        =   8700
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
      Begin XtremeSuiteControls.ComboBox cboProvincia 
         Height          =   312
         Left            =   -67600
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4683
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
      End
      Begin XtremeSuiteControls.ComboBox cboCanton 
         Height          =   312
         Left            =   -67600
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4683
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
      End
      Begin XtremeSuiteControls.ComboBox cboDistrito 
         Height          =   312
         Left            =   -67600
         TabIndex        =   33
         Top             =   1320
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4683
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
      End
      Begin XtremeSuiteControls.ComboBox cboGarantia 
         Height          =   312
         Left            =   -68080
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9340
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
      End
      Begin XtremeSuiteControls.ComboBox cboCartera 
         Height          =   312
         Left            =   -68080
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9340
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
      End
      Begin XtremeSuiteControls.ComboBox cboDestino 
         Height          =   312
         Left            =   -68080
         TabIndex        =   40
         Top             =   2160
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9340
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
      End
      Begin XtremeSuiteControls.ComboBox cboRecurso 
         Height          =   312
         Left            =   -68080
         TabIndex        =   41
         Top             =   2520
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9340
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
      End
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   312
         Left            =   -68080
         TabIndex        =   43
         Top             =   2880
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9340
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
      End
      Begin XtremeSuiteControls.CheckBox chkCantones 
         Height          =   252
         Left            =   -64720
         TabIndex        =   45
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkDistritos 
         Height          =   252
         Left            =   -64720
         TabIndex        =   46
         Top             =   1320
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkDepartamento 
         Height          =   252
         Left            =   -62440
         TabIndex        =   47
         Top             =   2280
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkSeccion 
         Height          =   252
         Left            =   -62440
         TabIndex        =   48
         Top             =   2640
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkLineas 
         Height          =   252
         Left            =   -62560
         TabIndex        =   49
         Top             =   1560
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   -68080
         TabIndex        =   50
         Top             =   1560
         Visible         =   0   'False
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   -67240
         TabIndex        =   51
         Top             =   1560
         Visible         =   0   'False
         Width           =   4452
         _Version        =   1441793
         _ExtentX        =   7853
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
      Begin XtremeSuiteControls.PushButton btnCubo_Procesa 
         Height          =   612
         Left            =   -63400
         TabIndex        =   62
         Top             =   2400
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Procesar Resultados (Cubo)"
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
      Begin VB.Label Label2 
         Caption         =   "Proceso para cargar información de los cambios en la antiguedad de mora, del periodo seleccionado.  (Cubos)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   -68680
         TabIndex        =   61
         Top             =   840
         Visible         =   0   'False
         Width           =   6612
      End
      Begin VB.Image Image3 
         Height          =   630
         Left            =   -69645
         Picture         =   "frmCC_ReportesAlCorte.frx":07BC
         Top             =   600
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatus 
         Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   732
         Left            =   -68680
         TabIndex        =   60
         Top             =   1680
         Visible         =   0   'False
         Width           =   4692
      End
      Begin VB.Label Label9 
         Caption         =   "Línea"
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
         Index           =   4
         Left            =   -69280
         TabIndex        =   52
         Top             =   1560
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label9 
         Caption         =   "Oficina"
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
         Index           =   3
         Left            =   -69280
         TabIndex        =   42
         Top             =   2880
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label10 
         Caption         =   "Destino"
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
         Left            =   -69280
         TabIndex        =   39
         Top             =   2160
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label9 
         Caption         =   "Recurso"
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
         Left            =   -69280
         TabIndex        =   38
         Top             =   2520
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label10 
         Caption         =   "Garantía"
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
         Left            =   -69280
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label9 
         Caption         =   "Cartera"
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
         Left            =   -69280
         TabIndex        =   34
         Top             =   960
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label18 
         Caption         =   "Distrito"
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
         Left            =   -68560
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label10 
         Caption         =   "Provincia"
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
         Left            =   -68560
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label9 
         Caption         =   "Cantón"
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
         Left            =   -68560
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label lblDepartamento 
         Caption         =   "Departamento"
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
         TabIndex        =   23
         Top             =   2280
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Label lblSeccion 
         Caption         =   "Sección"
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
         TabIndex        =   22
         Top             =   2640
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Label Label1 
         Caption         =   "Sector"
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
         Left            =   1440
         TabIndex        =   17
         Top             =   1860
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Profesión"
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
         Index           =   15
         Left            =   1440
         TabIndex        =   16
         Top             =   1500
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Estados"
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
         Index           =   8
         Left            =   1440
         TabIndex        =   15
         Top             =   576
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Zonas"
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
         Index           =   7
         Left            =   1440
         TabIndex        =   14
         Top             =   1056
         Width           =   732
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Condición Laboral"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   13
         Left            =   5400
         TabIndex        =   8
         Top             =   2496
         Width           =   1572
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   12
         Left            =   3840
         TabIndex        =   7
         Top             =   2496
         Width           =   1452
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   11
         Left            =   2280
         TabIndex        =   6
         Top             =   2496
         Width           =   1452
      End
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   312
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   7092
      _Version        =   1441793
      _ExtentX        =   12515
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   312
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   7092
      _Version        =   1441793
      _ExtentX        =   12515
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   252
      Index           =   1
      Left            =   2880
      TabIndex        =   54
      Top             =   2160
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Patrimonio"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   252
      Index           =   2
      Left            =   4560
      TabIndex        =   55
      Top             =   2160
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fondos"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   252
      Index           =   3
      Left            =   6240
      TabIndex        =   56
      Top             =   2160
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Créditos"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton optX 
      Height          =   252
      Index           =   4
      Left            =   7920
      TabIndex        =   57
      Top             =   2160
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cobros"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5295
      Left            =   9480
      TabIndex        =   63
      Top             =   0
      Width           =   9135
      _Version        =   524288
      _ExtentX        =   16113
      _ExtentY        =   9340
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCC_ReportesAlCorte.frx":0C8B
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption lblReporte 
      Height          =   372
      Left            =   0
      TabIndex        =   58
      Top             =   2520
      Width           =   9372
      _Version        =   1441793
      _ExtentX        =   16531
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes al Cierre para Operaciones"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   7332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Institución"
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
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   11532
   End
End
Attribute VB_Name = "frmCC_ReportesAlCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mAnio As Integer, mMes As Byte, vPaso As Boolean, mModoSif As Boolean
Dim mFiltro As String, mSQL As String, mSubTitulo As String, vHeaders As vGridHeaders


Private Sub btnCubo_Procesa_Click()
Call sbProcesa
End Sub

Private Sub sbSQL(Optional pTipo As String = "GEN")
Dim strSQL As String, rs As New ADODB.Recordset

mFiltro = ""
mSQL = ""
mSubTitulo = ""
If pTipo = "GEN" Then
 
    strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
    Call OpenRecordSet(rs, strSQL)
        mAnio = rs!Anio
        mMes = rs!Mes
    rs.Close
    
     Select Case Mid(lblReporte.Tag, 1, 1)
        Case "C" 'Creditos
         mSQL = mSQL & "{vSIFAuxCorteRepCredito.anio} = " & mAnio & " AND {vSIFAuxCorteRepCredito.mes} = " & mMes
        
        Case "J" 'Cobro Antiguedad
         mSQL = mSQL & "{vSIFAuxCorteAntiguedadSaldos.anio} = " & mAnio & " AND {vSIFAuxCorteAntiguedadSaldos.mes} = " & mMes
        
        Case Else
         mSQL = mSQL & "{vSIFAuxCorteRepMain.anio} = " & mAnio & " AND {vSIFAuxCorteRepMain.mes} = " & mMes
     End Select
End If

Select Case pTipo
    Case "GEN" 'General
    
         If cboEstado.Text <> "TODOS" Then
               If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
               mSQL = mSQL & "{vSIFAuxCorteRepMain.estadoactual} = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
        
         End If
         mFiltro = mFiltro & "¦ ESTADO : " & cboEstado.Text
            
         If Mid(cboSexo.Text, 1, 1) <> "T" Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
           mSQL = mSQL & "{vSIFAuxCorteRepMain.sexo} = '" & Mid(cboSexo.Text, 1, 1) & "'"
         End If
         mSubTitulo = mSubTitulo & " ¦ GENERO: " & cboSexo.Text
          
          If Mid(cboEstadoCivil.Text, 1, 1) <> "T" Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
           mSQL = mSQL & "{vSIFAuxCorteRepMain.EstadoCivil} = '" & cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex) & "'"
         End If
         mSubTitulo = mSubTitulo & " / Estado Civil : " & cboEstadoCivil.Text
         
         If cboCondicion.Text <> "TODOS" Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
             mSQL = mSQL & "{vSIFAuxCorteRepMain.EstadoLaboral} = '" & cboCondicion.ItemData(cboCondicion.ListIndex) & "'"
         End If
         mFiltro = mFiltro & "¦ LABORAL : " & cboCondicion.Text
         
        
         If cboInstitucion.Text <> "TODOS" Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
           mSQL = mSQL & "{vSIFAuxCorteRepMain.cod_institucion} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ""
         End If
         mFiltro = mFiltro & "¦ INSTITUCION : " & cboInstitucion.Text
          
        
         If cboSector.Text <> "TODOS" Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
           mSQL = mSQL & "{vSIFAuxCorteRepMain.cod_sector} = " & cboSector.ItemData(cboSector.ListIndex) & ""
         End If
         mFiltro = mFiltro & "¦ SECTOR : " & cboSector.Text
         
         
         If cboProfesion.Text <> "TODOS" Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
           mSQL = mSQL & "{vSIFAuxCorteRepMain.cod_profesion} = " & cboProfesion.ItemData(cboProfesion.ListIndex) & ""
         End If
         mFiltro = mFiltro & "¦ PROFESION : " & cboProfesion.Text
         
         
         If cboZonas.Text <> "TODOS" Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
           mSQL = mSQL & "{vSIFAuxCorteRepMain.cod_zona} = '" & cboZonas.ItemData(cboZonas.ListIndex) & "'"
         End If
         mFiltro = mFiltro & "¦ ZONA: " & cboZonas.Text
         
         
         'Filtros Adicionales
        If chkProvincias.Value = vbUnchecked And cboProvincia.ListCount > 0 Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
           mSQL = mSQL & "{vSIFAuxCorteRepMain.provincia} = " & cboProvincia.ItemData(cboProvincia.ListIndex)
           mFiltro = mFiltro & "¦ PROVINCIA: " & cboProvincia.Text
        Else
           mFiltro = mFiltro & "¦ PROVINCIA: Todas"
        End If
         
        If chkCantones.Value = vbUnchecked And cboCanton.ListCount > 0 Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
           mSQL = mSQL & "{vSIFAuxCorteRepMain.canton} = " & cboCanton.ItemData(cboCanton.ListIndex)
           mFiltro = mFiltro & " [" & cboCanton.Text & "]"
        End If
         
        If chkDistritos.Value = vbUnchecked And cboDistrito.ListCount > 0 Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
           mSQL = mSQL & "{vSIFAuxCorteRepMain.distrito} = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "'"
           mFiltro = mFiltro & " [" & cboDistrito.Text & "]"
        End If
         
        If chkDepartamento.Value = vbUnchecked And txtDeptCodigo.Text <> "" Then
           If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
           mSQL = mSQL & "{vSIFAuxCorteRepMain.DeptCod} = '" & txtDeptCodigo.Text & "'"
           mFiltro = mFiltro & " / Dept.: " & txtDeptCodigo.Text
            
           If chkSeccion.Value = vbUnchecked And txtSecCodigo.Text <> "" Then
               If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
               mSQL = mSQL & "{vSIFAuxCorteRepMain.SecCod} = '" & txtSecCodigo.Text & "'"
               mFiltro = mFiltro & " [" & txtSecCodigo.Text & "]"
           End If
        End If
        
   Case "CRD"
   
        ' Filtro de garatías
        If cboGarantia.Text <> "TODOS" Then
            If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
            mSQL = mSQL & "{vSIFAuxCorteRepCredito.Garantia} = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
            mFiltro = mFiltro & "¦ Garantía : " & cboGarantia.Text
        End If
        'Linea de crédito
        If Len(txtCodigo.Text) > 0 Then
            If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
            mSQL = mSQL & "{vSIFAuxCorteRepCredito.codigo} = '" & Trim(txtCodigo.Text) & "'"
            mFiltro = mFiltro & "¦ Línea : " & Trim(txtCodigo.Text)
        End If
        ' Filtro de destinos
        If cboDestino.Text <> "TODOS" Then
            If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
            mSQL = mSQL & "{vSIFAuxCorteRepCredito.COD_DESTINO} = '" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
            mFiltro = mFiltro & "¦ Destino : " & cboDestino.Text
        End If
        ' Filtro de oficina
        If cboOficina.Text <> "TODOS" Then
            If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
            mSQL = mSQL & "{vSIFAuxCorteRepCredito.COD_OFICINA_R} = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
            mFiltro = mFiltro & "¦ Oficina : " & cboOficina.Text
        End If
        ' Filtro de recurso
        If cboRecurso.Text <> "TODOS" Then
            If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
            mSQL = mSQL & "{vSIFAuxCorteRepCredito.COD_GRUPO} = '" & cboRecurso.ItemData(cboRecurso.ListIndex) & "'"
            mFiltro = mFiltro & "¦ Recurso : " & cboRecurso.Text
        End If
   
   Case "CBR" 'Cobros
   
   Case "ADS" 'Antiguedad de Saldos
        
        ' Filtro de garatías
        If cboGarantia.Text <> "TODOS" Then
            If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
            mSQL = mSQL & "{vSIFAuxCorteAntiguedadSaldos.Garantia} = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
            mFiltro = mFiltro & "¦ Garantía : " & cboGarantia.Text
        End If
        'Linea de crédito
        If Len(txtCodigo.Text) > 0 Then
            If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
            mSQL = mSQL & "{vSIFAuxCorteAntiguedadSaldos.codigo} = '" & Trim(txtCodigo.Text) & "'"
            mFiltro = mFiltro & "¦ Línea : " & Trim(txtCodigo.Text)
        End If
        ' Filtro de destinos
        If cboDestino.Text <> "TODOS" Then
            If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
            mSQL = mSQL & "{vSIFAuxCorteAntiguedadSaldos.COD_DESTINO} = '" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
            mFiltro = mFiltro & "¦ Destino : " & cboDestino.Text
        End If
        ' Filtro de oficina
        If cboOficina.Text <> "TODOS" Then
            If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
            mSQL = mSQL & "{vSIFAuxCorteAntiguedadSaldos.COD_OFICINA_R} = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
            mFiltro = mFiltro & "¦ Oficina : " & cboOficina.Text
        End If
        ' Filtro de recurso
        If cboRecurso.Text <> "TODOS" Then
            If Len(mSQL) > 0 Then mSQL = mSQL & " AND "
            mSQL = mSQL & "{vSIFAuxCorteAntiguedadSaldos.COD_GRUPO} = '" & cboRecurso.ItemData(cboRecurso.ListIndex) & "'"
            mFiltro = mFiltro & "¦ Recurso : " & cboRecurso.Text
        End If
        
    
        
End Select

End Sub


Private Sub btnInforme_Click()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String, vTemp As String
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If lblReporte.Tag = "J04" Then
    Exit Sub
End If

If lblReporte.Tag = "" Then
  MsgBox "Seleccione el reporte que desea visualizar!", vbExclamation
  Exit Sub
End If

If cboPeriodo.Text = "" Then
  MsgBox "No existen periodos registrados en el sistema.!", vbExclamation
  Exit Sub
End If


'Inicia el Reporte
Me.MousePointer = vbHourglass

vTitulo = UCase(lblReporte.Caption)
vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes al Corte del Sistema"
 
 .Connect = glogon.ConectRPT
  
 vSubTitulo = "Periodo : " & cboPeriodo.Text
  
 'Filtros Generales
 Call sbSQL("GEN")
 
 vFiltro = mFiltro
 strSQL = mSQL
 
 
'Filtros Especiales x Auxiliar
Select Case True
  Case optX.Item(0).Value  'Afiliacion
  Case optX.Item(1).Value  'Patrimonio
  Case optX.Item(2).Value  'Fondos
  
  
  Case optX.Item(3).Value  'Creditos
      
        Call sbSQL("CRD")
        
        If Len(mSQL) > 0 Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & mSQL
            vFiltro = vFiltro & ", " & mFiltro
        End If
  
  Case optX.Item(4).Value  'Cobro
     
        
        
    If lblReporte.Tag = "CA01" Then
        ' Filtra solo lo que tiene mora
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{vSIFAuxCorteRepCredito.MoraCuotas} > 0"
    End If
    
    If lblReporte.Tag = "J01" Then
        ' Filtra solo lo que tiene mora
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{vSIFAuxCorteRepCredito.MoraCuotas} > 0"
    Else ' lblReporte.Tag = "J01"
        'Filtros Adicionales
        Call sbSQL("ADS")
        
        If Len(mSQL) > 0 Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & mSQL
            vFiltro = vFiltro & ", " & mFiltro
        End If
    
    End If
    
End Select
 
 
    
'    ' Filtro de cobro clasificación
'    If cboCartera.Text <> "TODOS" Then
'        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'        strSQL = strSQL & "{CBR_CLASIFICACION_DETALLE.COD_CLASIFICACION} = '" & SIFGlobal.fxCodText(cboCartera.Text) & "'"
'        vFiltro = vFiltro & "¦ Cartera : " & cboCartera.Text
'    End If
 
 
 

 .Formulas(0) = "fxFecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='" & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='" & vTitulo & "'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(5) = "fxFiltro='" & vFiltro & "'"
 

 
Select Case lblReporte.Tag
  Case "A01" 'Cantidad de Personas x Estado
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteAFIPersonasEstados.rpt")
  Case "A02" 'Cantidad de Personas x Provincias
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteAFIPersonasProvincias.rpt")
  Case "A03" 'Cantidad de Personas x Zonas
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteAFIPersonasZonas.rpt")
  Case "A04" 'Cantidad de Personas x Insitucion
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteAFIPersonasInst.rpt")
  Case "A05" 'Listado General
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteAFIListadoGeneral.rpt")
  
  Case "P01" 'Total de Aportes
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCortePATTotal.rpt")
  Case "P02" 'Aportes x Provincias
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCortePATProvincias.rpt")
  Case "P03" 'Aportes x Zonas
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCortePATZonas.rpt")
  Case "P04" 'Aportes x Instituciones
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCortePATTotalInst.rpt")
  
  Case "F01" 'Total de Fondos
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteFNDTotalx.rpt")
  Case "F02" 'Fondos x Zona
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteFNDZonas.rpt")
  Case "F03" 'Fondos x Provincia
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteFNDProvincias.rpt")
  Case "F04" 'Fondos x Institucion
      .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteFNDTotalInst.rpt")
  
  Case "C00" 'Saldo de Cartera
    If Mid(cboTipo.Text, 1, 1) = "D" Then
        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCrd_Detallado.rpt")
    Else
        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCrd_Resumen.rpt")
    End If
  
  Case "C01" 'Saldo de Cartera x Linea
    If Mid(cboTipo.Text, 1, 1) = "D" Then
        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCrd_Linea_Detallado.rpt")
    Else
        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCrd_Linea_Resumen.rpt")
    End If
  
  Case "C02" 'Saldo de Cartera x Garantía
    If Mid(cboTipo.Text, 1, 1) = "D" Then
        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCrd_Garantia_Detallado.rpt")
    Else
        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCrd_Garantia_Resumen.rpt")
    End If
  
  Case "C03" 'Saldo de Cartera y Desembolso x Destino
  Case "C04" 'Saldo de Cartera x Zona
  Case "C05" 'Colocación x Zona en el Periodo
  Case "C06" 'Tasas y Plazos Ponderados de la Cartera x Línea
  Case "C07" 'Tasas y Plazos Ponderados de la Cartera x Garantía
    
  Case "CA01" 'Reporte General de Morosidad
    If Mid(cboTipo.Text, 1, 1) = "D" Then
        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCbrListadoDetallado.rpt")
    Else
        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCbrListadoResumen.rpt")
    End If
    
    
  Case "J02", "J02.1", "J02.2" 'Antiguedad de Saldos
    
    If lblReporte.Tag = "J02" Then 'Antiguedad de Saldos (Cobro Judicial + Prod.Acumulado)
            If Mid(cboTipo.Text, 1, 1) = "D" Then
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCbrAntiguedadDet.rpt")
            Else
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCbrAntiguedadRsm.rpt")
            End If
    End If
    
    If lblReporte.Tag = "J02.1" Then 'Antiguedad de Saldos + Prod.Acumulado
            If Mid(cboTipo.Text, 1, 1) = "D" Then
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCbrAntiguedadPADet.rpt")
            Else
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCbrAntiguedadPARsm.rpt")
            End If
    End If
    
    If lblReporte.Tag = "J02.2" Then 'Antiguedad de Saldos (LEGAL)
            If Mid(cboTipo.Text, 1, 1) = "D" Then
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCbrAntiguedadLegalDet.rpt")
            Else
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCbrAntiguedadLegalRsm.rpt")
            End If
    End If
    
    
    
  Case "J03" 'Antiguedad de Saldos (Financiera)
    If Mid(cboTipo.Text, 1, 1) = "D" Then
        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCbrAntiguedadFinancieraDet.rpt")
    Else
        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCorteCbrAntiguedadFinancieraRsm.rpt")
    End If
End Select

 .SelectionFormula = strSQL
' .PrintReport
 .Action = 1

End With

Me.MousePointer = vbDefault

'Call Bitacora("Imprime", lblReporte.Caption)


Exit Sub

vError:
 Me.MousePointer = vbDefault
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

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub


Private Sub chkCantones_Click()
If chkCantones.Value = vbChecked Then
   cboCanton.Enabled = False
Else
   cboCanton.Enabled = True
End If

chkDistritos.Value = chkCantones.Value
chkDistritos_Click
End Sub

Private Sub chkDistritos_Click()

If chkDistritos.Value = vbChecked Then
   cboDistrito.Enabled = False
Else
   cboDistrito.Enabled = True
End If

End Sub

Private Sub chkLineas_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If chkLineas.Value = vbChecked Then
  
  txtCodigo.Enabled = False
  
  strSQL = "select cod_grupo as 'IdX' , rtrim(descripcion) as ItmX" _
         & " from  catalogo_grupos"
  Call sbCbo_Llena_New(cboRecurso, strSQL, True, True)
  
  strSQL = "select cod_destino as 'IdX' , rtrim(descripcion) as ItmX" _
         & " from  catalogo_destinos"
  Call sbCbo_Llena_New(cboDestino, strSQL, True, True)
  
Else
  txtCodigo.Enabled = True

  strSQL = "select (R.cod_grupo) as 'IdX' , rtrim(R.descripcion) as ItmX" _
         & " from catalogo_grupos R inner join catalogo_AsignaGrp A on R.cod_grupo = A.cod_grupo" _
         & " where A.codigo = '" & txtCodigo & "'"
  Call sbCbo_Llena_New(cboRecurso, strSQL, False, True)
  
  strSQL = "select (R.cod_destino) as 'IdX' , rtrim(R.descripcion) as ItmX" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtCodigo & "'"
  Call sbCbo_Llena_New(cboDestino, strSQL, False, True)

End If

End Sub

Private Sub chkProvincias_Click()
If chkProvincias.Value = vbChecked Then
   cboProvincia.Enabled = False
Else
   cboProvincia.Enabled = True
End If

chkCantones.Value = chkProvincias.Value
chkCantones_Click

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Informes:", 8500
End With
tcMain.Item(3).Visible = False

strSQL = "select OBJECT_ID('UPROGRAMATICA') as Resultado"
Call OpenRecordSet(rs, strSQL)
If IsNull(rs!Resultado) Then
  mModoSif = True
  lblDepartamento.Caption = "Departamento"
  lblSeccion.Caption = "Sección"
Else
  mModoSif = False
  lblDepartamento.Caption = "Unidad Programatica"
  lblSeccion.Caption = "Unidad de Trabajo"
End If
rs.Close

strSQL = "select * from ase_per_historico order by anio desc,mes desc"
Call OpenRecordSet(rs, strSQL)

cboPeriodo.Clear
Do While Not rs.EOF
 cboPeriodo.AddItem rs!Anio & " - " & fxConvierteMES(rs!Mes)
 cboPeriodo.ItemData(cboPeriodo.ListCount - 1) = CStr(rs!id_per_historico)
 
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
 rs.MoveFirst
 cboPeriodo.Text = rs!Anio & " - " & fxConvierteMES(rs!Mes)
End If
rs.Close

strSQL = "select COD_INSTITUCION as Idx,descripcion as ItmX from INSTITUCIONES"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)

strSQL = "select COD_PROFESION as Idx,descripcion as ItmX from AFI_PROFESIONES"
Call sbCbo_Llena_New(cboProfesion, strSQL, True, True)

strSQL = "select COD_SECTOR as Idx,descripcion as ItmX from AFI_SECTORES"
Call sbCbo_Llena_New(cboSector, strSQL, True, True)

strSQL = "select COD_ZONA as 'IdX', rtrim(descripcion) as 'ItmX' from AFI_ZONAS"
Call sbCbo_Llena_New(cboZonas, strSQL, True, False)


strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  afi_Estados_Persona"
Call sbCbo_Llena_New(cboEstado, strSQL, True)

strSQL = "select rtrim(GARANTIA) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from CRD_GARANTIA_TIPOS order by GARANTIA"
Call sbCbo_Llena_New(cboGarantia, strSQL, True, False)

strSQL = "select rtrim(cod_clasificacion) as 'IdX' , rtrim(descripcion) as 'ItmX'" _
        & " from CBR_CLASIFICACION_CARTERA order by cod_clasificacion"
Call sbCbo_Llena_New(cboCartera, strSQL, True, False)

strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas order by cod_oficina"
Call sbCbo_Llena_New(cboOficina, strSQL, True, False)

strSQL = "select Estado_Civil as 'IdX', Descripcion as 'ItmX' from SYS_ESTADO_CIVIL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoCivil, strSQL, True, True)

cboSexo.Clear
cboSexo.AddItem "TODOS"
cboSexo.AddItem "Femenino"
cboSexo.AddItem "Masculino"
cboSexo.Text = "TODOS"

strSQL = "select ESTADO_LABORAL as 'IdX', Descripcion as 'ItmX' from AFI_ESTADO_LABORAL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboCondicion, strSQL, True, True)

cboTipo.Clear
cboTipo.AddItem "Detalle"
cboTipo.AddItem "Resumen"
cboTipo.Text = "Detalle"

vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False

Call chkLineas_Click

Call OptX_Click(0)

chkProvincias_Click

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lsw.ListItems.Count <= 0 Then Exit Sub

lblReporte.Tag = Item.Key
lblReporte.Caption = Item.Text
    
tcMain.Item(0).Selected = True

Select Case lblReporte.Tag
    Case "J04", "J05", "J06", "J07", "J10", "J10d"
        tcMain.Item(3).Selected = True
    Case Else
End Select

End Sub


Private Sub OptX_Click(Index As Integer)
Dim itmX As ListViewItem

lsw.ListItems.Clear
lblReporte.Tag = ""
lblReporte.Caption = ""

tcMain.Item(0).Selected = True


Select Case Index
  Case 0 'Afiliacion
    Set itmX = lsw.ListItems.Add(, "A01", "Cantidad de Personas x Estado")
    Set itmX = lsw.ListItems.Add(, "A02", "Cantidad de Personas x Provincias")
    Set itmX = lsw.ListItems.Add(, "A03", "Cantidad de Personas x Zonas")
    Set itmX = lsw.ListItems.Add(, "A04", "Cantidad de Personas x Institucion")
    Set itmX = lsw.ListItems.Add(, "A05", "Listado General")
  
  Case 1 'Patrimonio
    Set itmX = lsw.ListItems.Add(, "P01", "Total de Aportes")
    Set itmX = lsw.ListItems.Add(, "P02", "Aportes x Provincias")
    Set itmX = lsw.ListItems.Add(, "P03", "Aportes x Zonas")
    Set itmX = lsw.ListItems.Add(, "P04", "Aportes x Instituciones")
  
  Case 2 'Fondos
    Set itmX = lsw.ListItems.Add(, "F01", "Total de Fondos")
    Set itmX = lsw.ListItems.Add(, "F02", "Fondos x Zona")
    Set itmX = lsw.ListItems.Add(, "F03", "Fondos x Provincia")
    Set itmX = lsw.ListItems.Add(, "F04", "Fondos x Institucion")
  
  Case 3 'Credito
    Set itmX = lsw.ListItems.Add(, "C00", "Saldo de Cartera")
    Set itmX = lsw.ListItems.Add(, "C01", "Saldo de Cartera x Linea")
    Set itmX = lsw.ListItems.Add(, "C02", "Saldo de Cartera x Garantía")
'    Set itmX = lsw.ListItems.Add(, "C03", "Saldo de Cartera y Desembolso x Destino")
'    Set itmX = lsw.ListItems.Add(, "C04", "Saldo de Cartera x Zona")
'    Set itmX = lsw.ListItems.Add(, "C05", "Colocación x Zona en el Periodo")
'    Set itmX = lsw.ListItems.Add(, "C06", "Tasas y Plazos Ponderados de la Cartera x Línea")
'    Set itmX = lsw.ListItems.Add(, "C07", "Tasas y Plazos Ponderados de la Cartera x Garantía")
    

  Case 4 'Cobro
    Set itmX = lsw.ListItems.Add(, "CA01", "Reporte General de Morosidad")
    Set itmX = lsw.ListItems.Add(, "J02", "Antiguedad de Saldos")
    Set itmX = lsw.ListItems.Add(, "J02.1", "Antiguedad de Saldos + Prod.Acumulado")
    Set itmX = lsw.ListItems.Add(, "J02.2", "Antiguedad de Saldos (Legal)")
    Set itmX = lsw.ListItems.Add(, "J03", "Antiguedad de Saldos (Financiera)")
    Set itmX = lsw.ListItems.Add(, "J04", "Cargar Comparativo de Cambios en Antiguedad Mora (Cubos)")
    Set itmX = lsw.ListItems.Add(, "J05", "Cargar Morosidad (Cubos)")
    Set itmX = lsw.ListItems.Add(, "J06", "Cargar Morosidad por Cartera (Cubos)")
    Set itmX = lsw.ListItems.Add(, "J07", "Informe de Antiguedad por Cartera (Cubos)")
    Set itmX = lsw.ListItems.Add(, "J10", "Estimación Incobrables Resumen")
    Set itmX = lsw.ListItems.Add(, "J10d", "Estimación Incobrables Detalle")

End Select

End Sub




Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then cboDestino.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCodigo_LostFocus()
 If Len(Trim(txtCodigo)) > 0 Then txtDescripcion.Text = fxDescribeCodigo(Trim(txtCodigo))
 Call chkLineas_Click
End Sub

Private Sub txtDeptCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptDesc.SetFocus
If KeyCode = vbKeyF4 Then
  
    If mModoSif Then
      gBusquedas.Columna = "cod_departamento"
      gBusquedas.Orden = "cod_departamento"
      gBusquedas.Consulta = "select cod_departamento as codigo,descripcion from afDepartamentos"
      gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    Else
      gBusquedas.Columna = "codigo"
      gBusquedas.Orden = "codigo"
      gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
      gBusquedas.Filtro = ""
    End If

  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDeptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecCodigo.SetFocus
If KeyCode = vbKeyF4 Then

    If mModoSif Then
      gBusquedas.Columna = "descripcion"
      gBusquedas.Orden = "descripcion"
      gBusquedas.Consulta = "select cod_departamento as codigo,descripcion from afDepartamentos"
      gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    Else
      gBusquedas.Columna = "codigo"
      gBusquedas.Orden = "codigo"
      gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
      gBusquedas.Filtro = ""
    End If
  
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtSecCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecDesc.SetFocus
If KeyCode = vbKeyF4 Then
  
    If mModoSif Then
        gBusquedas.Columna = "cod_seccion"
        gBusquedas.Orden = "cod_seccion"
        gBusquedas.Consulta = "select cod_seccion as codigo,descripcion from afSecciones"
        gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
                    & " and cod_departamento = '" & txtDeptCodigo.Text & "'"
    Else
        gBusquedas.Columna = "ut_codigo"
        gBusquedas.Orden = "ut_codigo"
        gBusquedas.Consulta = "select ut_codigo,ut_descripcion from UTRABAJO"
        gBusquedas.Filtro = ""
    End If
  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtSecDesc_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then optPropiedad.Item(0).SetFocus
If KeyCode = vbKeyF4 Then
    If mModoSif Then
        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Consulta = "select cod_seccion as codigo,descripcion from afSecciones"
        gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
                    & " and cod_departamento = '" & txtDeptCodigo.Text & "'"
    Else
        gBusquedas.Columna = "ut_descripcion"
        gBusquedas.Orden = "ut_descripcion"
        gBusquedas.Consulta = "select ut_codigo,ut_descripcion from UTRABAJO"
        gBusquedas.Filtro = ""
    End If
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub sbvGridColSize()
Dim i As Integer

For i = 1 To vGrid.MaxCols
  vGrid.ColWidth(i) = vGrid.MaxTextColWidth(i)
Next i

End Sub


Private Sub sbvGridCol(Index As Integer)

vGrid.DAutoSizeCols = DAutoSizeColsMax
vGrid.DAutoHeadings = False

Select Case Index
  Case 1 'Comprativo
      vGrid.MaxCols = 25
      vGrid.SetText 1, 0, "Año"
      vGrid.SetText 2, 0, "Mes"
      vGrid.SetText 3, 0, "Línea"
      vGrid.SetText 4, 0, "Línea Desc"
      vGrid.SetText 5, 0, "Retención"
      vGrid.SetText 6, 0, "Póliza"
      vGrid.SetText 7, 0, "Oficina"
      vGrid.SetText 8, 0, "Ultimo Mov."
      vGrid.SetText 9, 0, "Garantía"
      vGrid.SetText 10, 0, "Destino"
      vGrid.SetText 11, 0, "Recurso"
      vGrid.SetText 12, 0, "Plazo Restante"
      vGrid.SetText 13, 0, "Monto"
      vGrid.SetText 14, 0, "Saldo"
      vGrid.SetText 15, 0, "Cuota al Corte"
      vGrid.SetText 16, 0, "Tasa"
      vGrid.SetText 17, 0, "Plazo"
      vGrid.SetText 18, 0, "Pri.Deduc."
      vGrid.SetText 19, 0, "Mora Intereses"
      vGrid.SetText 20, 0, "Mora Cargos"
      vGrid.SetText 21, 0, "Mora Principal"
      vGrid.SetText 22, 0, "Mora Cuotas"
      vGrid.SetText 23, 0, "Mora Cta Antigua"
      vGrid.SetText 24, 0, "Antiguedad"
      vGrid.SetText 25, 0, "Antiguedad Anterior"
      
      vHeaders.Columnas = 25
      vHeaders.Headers(1) = "Año"
      vHeaders.Headers(2) = "Mes"
      vHeaders.Headers(3) = "Línea"
      vHeaders.Headers(4) = "Línea Desc"
      vHeaders.Headers(5) = "Retención"
      vHeaders.Headers(6) = "Póliza"
      vHeaders.Headers(7) = "Oficina"
      vHeaders.Headers(8) = "Ultimo Mov."
      vHeaders.Headers(9) = "Garantía"
      vHeaders.Headers(10) = "Destino"
      vHeaders.Headers(11) = "Recurso"
      vHeaders.Headers(12) = "Plazo Restante"
      vHeaders.Headers(13) = "Monto"
      vHeaders.Headers(14) = "Saldo"
      vHeaders.Headers(15) = "Cuota al Corte"""
      vHeaders.Headers(16) = "Tasa"
      vHeaders.Headers(17) = "Plazo"
      vHeaders.Headers(18) = "Pri.Deduc."
      vHeaders.Headers(19) = "Mora Intereses"
      vHeaders.Headers(20) = "Mora Cargos"
      vHeaders.Headers(21) = "Mora Principal"
      vHeaders.Headers(22) = "Mora Cuotas"
      vHeaders.Headers(23) = "Mora Cta Antigua"
      vHeaders.Headers(24) = "Antiguedad"
      vHeaders.Headers(25) = "Antiguedad Anterior"

  
  Case 2 'Morosidad
      vGrid.MaxCols = 36
      vGrid.SetText 1, 0, "Año"
      vGrid.SetText 2, 0, "Mes"
      vGrid.SetText 3, 0, "Línea"
      vGrid.SetText 4, 0, "Línea Desc"
      vGrid.SetText 5, 0, "Retención"
      vGrid.SetText 6, 0, "Póliza"
      vGrid.SetText 7, 0, "Oficina"
      vGrid.SetText 8, 0, "Ultimo Mov."
      vGrid.SetText 9, 0, "Garantía"
      vGrid.SetText 10, 0, "Destino"
      vGrid.SetText 11, 0, "Recurso"
      vGrid.SetText 12, 0, "Plazo Restante"
      vGrid.SetText 13, 0, "Monto"
      vGrid.SetText 14, 0, "Saldo"
      vGrid.SetText 15, 0, "Cuota al Corte"
      vGrid.SetText 16, 0, "Tasa"
      vGrid.SetText 17, 0, "Plazo"
      vGrid.SetText 18, 0, "Pri.Deduc."
      vGrid.SetText 19, 0, "Mora Intereses"
      vGrid.SetText 20, 0, "Mora Cargos"
      vGrid.SetText 21, 0, "Mora Principal"
      vGrid.SetText 22, 0, "Mora Cuotas"
      vGrid.SetText 23, 0, "Mora Cta Antigua"
      vGrid.SetText 24, 0, "Antiguedad"
      vGrid.SetText 25, 0, "Mora Financiera"
      vGrid.SetText 26, 0, "Mora Legal"
      vGrid.SetText 27, 0, "Cédula"
      vGrid.SetText 28, 0, "Nombre"
      vGrid.SetText 29, 0, "Empresa"
      vGrid.SetText 30, 0, "Provincia"
      vGrid.SetText 31, 0, "Comité Evaluador"
      vGrid.SetText 32, 0, "Dept.Id"
      vGrid.SetText 33, 0, "Deptartamento"
      vGrid.SetText 34, 0, "No.Operación"
      vGrid.SetText 35, 0, "Membresía"
      vGrid.SetText 36, 0, "Estado Persona"
      
      vHeaders.Columnas = 36
      vHeaders.Headers(1) = "Año"
      vHeaders.Headers(2) = "Mes"
      vHeaders.Headers(3) = "Línea"
      vHeaders.Headers(4) = "Línea Desc"
      vHeaders.Headers(5) = "Retención"
      vHeaders.Headers(6) = "Póliza"
      vHeaders.Headers(7) = "Oficina"
      vHeaders.Headers(8) = "Ultimo Mov."
      vHeaders.Headers(9) = "Garantía"
      vHeaders.Headers(10) = "Destino"
      vHeaders.Headers(11) = "Recurso"
      vHeaders.Headers(12) = "Plazo Restante"
      vHeaders.Headers(13) = "Monto"
      vHeaders.Headers(14) = "Saldo"
      vHeaders.Headers(15) = "Cuota al Corte"""
      vHeaders.Headers(16) = "Tasa"
      vHeaders.Headers(17) = "Plazo"
      vHeaders.Headers(18) = "Pri.Deduc."
      vHeaders.Headers(19) = "Mora Intereses"
      vHeaders.Headers(20) = "Mora Cargos"
      vHeaders.Headers(21) = "Mora Principal"
      vHeaders.Headers(22) = "Mora Cuotas"
      vHeaders.Headers(23) = "Mora Cta Antigua"
      vHeaders.Headers(24) = "Antiguedad"
      vHeaders.Headers(25) = "Mora Financiera"
      vHeaders.Headers(26) = "Mora Legal"
      vHeaders.Headers(27) = "Cédula"
      vHeaders.Headers(28) = "Nombre"
      vHeaders.Headers(29) = "Empresa"
      vHeaders.Headers(30) = "Provincia"
      vHeaders.Headers(31) = "Comité Evaluador"
      vHeaders.Headers(32) = "Dept.Id"
      vHeaders.Headers(33) = "Deptartamento"
      vHeaders.Headers(34) = "No.Operación"
      vHeaders.Headers(35) = "Membresía"
      vHeaders.Headers(36) = "Estado Persona"


   
  Case 3 'Antiguedad por Tipo
      vGrid.MaxCols = 37
      vGrid.SetText 1, 0, "Año"
      vGrid.SetText 2, 0, "Mes"
      vGrid.SetText 3, 0, "Línea"
      vGrid.SetText 4, 0, "Línea Desc"
      vGrid.SetText 5, 0, "Retención"
      vGrid.SetText 6, 0, "Póliza"
      vGrid.SetText 7, 0, "Oficina"
      vGrid.SetText 8, 0, "Ultimo Mov."
      vGrid.SetText 9, 0, "Garantía"
      vGrid.SetText 10, 0, "Destino"
      vGrid.SetText 11, 0, "Recurso"
      vGrid.SetText 12, 0, "Plazo Restante"
      vGrid.SetText 13, 0, "Monto"
      vGrid.SetText 14, 0, "Saldo"
      vGrid.SetText 15, 0, "Cuota al Corte"
      vGrid.SetText 16, 0, "Tasa"
      vGrid.SetText 17, 0, "Plazo"
      vGrid.SetText 18, 0, "Pri.Deduc."
      vGrid.SetText 19, 0, "Mora Intereses"
      vGrid.SetText 20, 0, "Mora Cargos"
      vGrid.SetText 21, 0, "Mora Principal"
      vGrid.SetText 22, 0, "Mora Cuotas"
      vGrid.SetText 23, 0, "Mora Cta Antigua"
      vGrid.SetText 24, 0, "Antiguedad"
      vGrid.SetText 25, 0, "Mora Financiera"
      vGrid.SetText 26, 0, "Mora Legal"
      vGrid.SetText 27, 0, "Cédula"
      vGrid.SetText 28, 0, "Nombre"
      vGrid.SetText 29, 0, "Empresa"
      vGrid.SetText 30, 0, "Provincia"
      vGrid.SetText 31, 0, "Comité Evaluador"
      vGrid.SetText 32, 0, "Dept.Id"
      vGrid.SetText 33, 0, "Deptartamento"
      vGrid.SetText 34, 0, "No.Operación"
      vGrid.SetText 35, 0, "Membresía"
      vGrid.SetText 36, 0, "Cartera Tipo"
      vGrid.SetText 37, 0, "Estado Persona"
      
      vHeaders.Columnas = 37
      vHeaders.Headers(1) = "Año"
      vHeaders.Headers(2) = "Mes"
      vHeaders.Headers(3) = "Línea"
      vHeaders.Headers(4) = "Línea Desc"
      vHeaders.Headers(5) = "Retención"
      vHeaders.Headers(6) = "Póliza"
      vHeaders.Headers(7) = "Oficina"
      vHeaders.Headers(8) = "Ultimo Mov."
      vHeaders.Headers(9) = "Garantía"
      vHeaders.Headers(10) = "Destino"
      vHeaders.Headers(11) = "Recurso"
      vHeaders.Headers(12) = "Plazo Restante"
      vHeaders.Headers(13) = "Monto"
      vHeaders.Headers(14) = "Saldo"
      vHeaders.Headers(15) = "Cuota al Corte"""
      vHeaders.Headers(16) = "Tasa"
      vHeaders.Headers(17) = "Plazo"
      vHeaders.Headers(18) = "Pri.Deduc."
      vHeaders.Headers(19) = "Mora Intereses"
      vHeaders.Headers(20) = "Mora Cargos"
      vHeaders.Headers(21) = "Mora Principal"
      vHeaders.Headers(22) = "Mora Cuotas"
      vHeaders.Headers(23) = "Mora Cta Antigua"
      vHeaders.Headers(24) = "Antiguedad"
      vHeaders.Headers(25) = "Mora Financiera"
      vHeaders.Headers(26) = "Mora Legal"
      vHeaders.Headers(27) = "Cédula"
      vHeaders.Headers(28) = "Nombre"
      vHeaders.Headers(29) = "Empresa"
      vHeaders.Headers(30) = "Provincia"
      vHeaders.Headers(31) = "Comité Evaluador"
      vHeaders.Headers(32) = "Dept.Id"
      vHeaders.Headers(33) = "Deptartamento"
      vHeaders.Headers(34) = "No.Operación"
      vHeaders.Headers(35) = "Membresía"
      vHeaders.Headers(36) = "Cartera Tipo"
      vHeaders.Headers(37) = "Estado Persona"
 



  Case 4 'Antiguedad por Base Días Real
      vGrid.MaxCols = 27
      vGrid.SetText 1, 0, "Año"
      vGrid.SetText 2, 0, "Mes"
      vGrid.SetText 3, 0, "Línea"
      vGrid.SetText 4, 0, "Línea Desc"
      vGrid.SetText 5, 0, "Oficina"
      vGrid.SetText 6, 0, "Garantía"
      vGrid.SetText 7, 0, "Destino"
      vGrid.SetText 8, 0, "Recurso"
      vGrid.SetText 9, 0, "Divisa"
      vGrid.SetText 10, 0, "Monto"
      vGrid.SetText 11, 0, "Saldo"
      vGrid.SetText 12, 0, "Mora Intereses"
      vGrid.SetText 13, 0, "Mora Cargos"
      vGrid.SetText 14, 0, "Mora Principal"
      vGrid.SetText 15, 0, "Mora Cuotas"
      vGrid.SetText 16, 0, "Mora Cta Antigua"
      vGrid.SetText 17, 0, "Mora Financiera"
      vGrid.SetText 18, 0, "Saldo + PA (Cbr)"
      vGrid.SetText 19, 0, "Mora Legal"
      vGrid.SetText 20, 0, "Empresa"
      vGrid.SetText 21, 0, "Provincia"
      vGrid.SetText 22, 0, "Comité Evaluador"
      vGrid.SetText 23, 0, "Cartera Tipo"
      vGrid.SetText 24, 0, "Estado Persona"
      vGrid.SetText 25, 0, "Antiguedad"
      vGrid.SetText 26, 0, "Estado Laboral"
      vGrid.SetText 27, 0, "Qty Operaciones"
      
      vHeaders.Columnas = 27
      vHeaders.Headers(1) = "Año"
      vHeaders.Headers(2) = "Mes"
      vHeaders.Headers(3) = "Línea"
      vHeaders.Headers(4) = "Línea Desc"
      vHeaders.Headers(5) = "Oficina"
      vHeaders.Headers(6) = "Garantía"
      vHeaders.Headers(7) = "Destino"
      vHeaders.Headers(8) = "Recurso"
      vHeaders.Headers(9) = "Divisa"
      vHeaders.Headers(10) = "Monto"
      vHeaders.Headers(11) = "Saldo"
      vHeaders.Headers(12) = "Mora Intereses"
      vHeaders.Headers(13) = "Mora Cargos"
      vHeaders.Headers(14) = "Mora Principal"
      vHeaders.Headers(15) = "Mora Cuotas"
      vHeaders.Headers(16) = "Mora Cta Antigua"
      vHeaders.Headers(17) = "Mora Financiera"
      vHeaders.Headers(18) = "Saldo + PA (Cbr)"
      vHeaders.Headers(19) = "Mora Legal"
      vHeaders.Headers(20) = "Empresa"
      vHeaders.Headers(21) = "Provincia"
      vHeaders.Headers(22) = "Comité Evaluador"
      vHeaders.Headers(23) = "Cartera Tipo"
      vHeaders.Headers(24) = "Estado Persona"
      vHeaders.Headers(25) = "Antiguedad"
      vHeaders.Headers(26) = "Estado Laboral"
      vHeaders.Headers(27) = "Qty Operaciones"



  Case 5 'Estimacion Resumen
      vGrid.MaxCols = 13
      vGrid.SetText 1, 0, "Año"
      vGrid.SetText 2, 0, "Mes"
      vGrid.SetText 3, 0, "Garantía"
      vGrid.SetText 4, 0, "Antiguedad Id"
      vGrid.SetText 5, 0, "Antiguedad"
      vGrid.SetText 6, 0, "% Mitigador"
      vGrid.SetText 7, 0, "Est. % Saldo Cubierto"
      vGrid.SetText 8, 0, "Est. % Saldo No Cubierto"
      vGrid.SetText 9, 0, "Total Saldo"
      vGrid.SetText 10, 0, "Qty Operaciones"
      vGrid.SetText 11, 0, "Monto Mitigado"
      vGrid.SetText 12, 0, "Saldo no Cubierto"
      vGrid.SetText 13, 0, "Monto Estimación"
      
      vHeaders.Columnas = 13
      vHeaders.Headers(1) = "Año"
      vHeaders.Headers(2) = "Mes"
      vHeaders.Headers(3) = "Garantía"
      vHeaders.Headers(4) = "Antiguedad Id"
      vHeaders.Headers(5) = "Antiguedad"
      vHeaders.Headers(6) = "% Mitigador"
      vHeaders.Headers(7) = "Est. % Saldo Cubierto"
      vHeaders.Headers(8) = "Est. % Saldo No Cubierto"
      vHeaders.Headers(9) = "Total Saldo"
      vHeaders.Headers(10) = "Qty Operaciones"
      vHeaders.Headers(11) = "Monto Mitigado"
      vHeaders.Headers(12) = "Saldo no Cubierto"
      vHeaders.Headers(13) = "Monto Estimación"


  Case 6 'Estimacion Detalle
      vGrid.MaxCols = 16
      vGrid.SetText 1, 0, "Año"
      vGrid.SetText 2, 0, "Mes"
      vGrid.SetText 3, 0, "Cédula"
      vGrid.SetText 4, 0, "Nombre"
      vGrid.SetText 5, 0, "No.Operación"
      vGrid.SetText 6, 0, "Línea Id"
      vGrid.SetText 7, 0, "Línea Desc"
      vGrid.SetText 8, 0, "Garantía"
      vGrid.SetText 9, 0, "Antiguedad Id"
      vGrid.SetText 10, 0, "Antiguedad"
      vGrid.SetText 11, 0, "% Mitigador"
      vGrid.SetText 12, 0, "Est. % Saldo No Cubierto"
      vGrid.SetText 13, 0, "Saldo"
      vGrid.SetText 14, 0, "Monto Mitigado"
      vGrid.SetText 15, 0, "Saldo no Cubierto"
      vGrid.SetText 16, 0, "Monto Estimación"
      
      vHeaders.Columnas = 16
      vHeaders.Headers(1) = "Año"
      vHeaders.Headers(2) = "Mes"
      vHeaders.Headers(3) = "Cédula"
      vHeaders.Headers(4) = "Nombre"
      vHeaders.Headers(5) = "No.Operación"
      vHeaders.Headers(6) = "Línea Id"
      vHeaders.Headers(7) = "Línea Desc"
      vHeaders.Headers(8) = "Garantía"
      vHeaders.Headers(9) = "Antiguedad Id"
      vHeaders.Headers(10) = "Antiguedad"
      vHeaders.Headers(11) = "% Mitigador"
      vHeaders.Headers(12) = "Est. % Saldo No Cubierto"
      vHeaders.Headers(13) = "Saldo"
      vHeaders.Headers(14) = "Monto Mitigado"
      vHeaders.Headers(15) = "Saldo no Cubierto"
      vHeaders.Headers(16) = "Monto Estimación"

End Select

End Sub

Private Sub sbProcesaCubos(pSQL As String, pGridCol As Integer)


On Error GoTo vError

Me.MousePointer = vbHourglass
    
tcMain.Item(1).Selected = True
    
lblStatus.Caption = "Procesando, Espere!"
DoEvents

Call sbvGridCol(pGridCol)

Call sbCargaGrid(vGrid, vGrid.MaxCols, pSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

Call sbvGridColSize

lblStatus.Caption = ""

Me.MousePointer = vbDefault


MsgBox "Informe Procesado!", vbInformation

'Exporta a Excel Automático
Call sbSIFGridExportar(vGrid, vHeaders, lblReporte.Caption)

Exit Sub

vError:
  Me.MousePointer = vbDefault


End Sub



Private Sub sbProcesa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, pGridCol As Integer

On Error GoTo vError

    If cboPeriodo.Text = "" Then
      MsgBox "No existen periodos registrados en el sistema.!", vbExclamation
      Exit Sub
    End If
    
    mAnio = 0
    mMes = 0
    pGridCol = 1
    
    strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
    Call OpenRecordSet(rs, strSQL)
        mAnio = rs!Anio
        mMes = rs!Mes
    rs.Close
    
    If mAnio = 0 Or mMes = 0 Then
        lblStatus.Caption = "Seleccione el periodo que desea cargar"
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass

    Select Case lblReporte.Tag
        Case "J04"
            strSQL = "exec spCbrAnalisisMorosidadCompartivo " & mAnio & "," & mMes
            pGridCol = 1
        
        Case "J05"
            strSQL = "exec spCbrAnalisisMorosidad " & mAnio & "," & mMes
            pGridCol = 2
        
        Case "J06"
            strSQL = "exec spCbrAnalisisMorosidadxTipo " & mAnio & "," & mMes
            pGridCol = 3
            
        Case "J07" 'Informe de Antiguedad Dias Real al Corte
            strSQL = "exec spCbrAnalisisCorteAntiguedadDiasReal " & mAnio & "," & mMes
            pGridCol = 4
        
        Case "J10" 'Estimacion Resumen
            strSQL = "exec spCbr_Estimacion_Resumen " & mAnio & "," & mMes
            pGridCol = 5
        
        Case "J10d" 'Estimacion Detalle
          
            strSQL = "exec spCbr_Estimacion_Detalle " & mAnio & "," & mMes
            pGridCol = 6
    End Select
    

Call sbProcesaCubos(strSQL, pGridCol)

lblStatus.Caption = "Proceso Concluido con éxito!"

Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


