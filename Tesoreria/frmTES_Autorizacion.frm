VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmTES_Autorizacion 
   Caption         =   "Autorización de Emisión de Documentos y Firmas Electrónicas"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14280
   Icon            =   "frmTES_Autorizacion.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   14280
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3495
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   5535
      _Version        =   1441793
      _ExtentX        =   9763
      _ExtentY        =   6165
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
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      FlatScrollBar   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.GroupBox fraRangos 
      Height          =   3735
      Left            =   7080
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   5055
      _Version        =   1441793
      _ExtentX        =   8911
      _ExtentY        =   6583
      _StockProps     =   79
      Caption         =   "Rangos de Autorización registrados al Usuario "
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
      Begin XtremeSuiteControls.PushButton btnCerrar_Rangos 
         Height          =   555
         Left            =   3480
         TabIndex        =   25
         Top             =   3120
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   979
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         Picture         =   "frmTES_Autorizacion.frx":6852
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
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
         Left            =   240
         TabIndex        =   41
         Top             =   1200
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "Corte"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1560
         Width           =   612
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Autorización General"
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
         Height          =   312
         Index           =   0
         Left            =   960
         TabIndex        =   39
         Top             =   480
         Width           =   1932
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Autorización Documento"
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
         Height          =   312
         Index           =   1
         Left            =   2880
         TabIndex        =   38
         Top             =   480
         Width           =   1932
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Emisión"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   312
         Index           =   2
         Left            =   960
         TabIndex        =   37
         Top             =   840
         Width           =   3852
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Firmas Electrónicas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   312
         Index           =   3
         Left            =   960
         TabIndex        =   36
         Top             =   1920
         Width           =   3852
      End
      Begin VB.Label lblEmAGInicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   960
         TabIndex        =   35
         Top             =   1200
         Width           =   1932
      End
      Begin VB.Label lblEmADInicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   2880
         TabIndex        =   34
         Top             =   1200
         Width           =   1932
      End
      Begin VB.Label lblEmAGCorte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   960
         TabIndex        =   33
         Top             =   1560
         Width           =   1932
      End
      Begin VB.Label lblEmADCorte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   2880
         TabIndex        =   32
         Top             =   1560
         Width           =   1932
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
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
         Left            =   240
         TabIndex        =   31
         Top             =   2280
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "Corte"
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
         Index           =   5
         Left            =   240
         TabIndex        =   30
         Top             =   2640
         Width           =   612
      End
      Begin VB.Label lblFiAGInicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   960
         TabIndex        =   29
         Top             =   2280
         Width           =   1932
      End
      Begin VB.Label lblFiADInicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   2880
         TabIndex        =   28
         Top             =   2280
         Width           =   1932
      End
      Begin VB.Label lblFiAGCorte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   960
         TabIndex        =   27
         Top             =   2640
         Width           =   1932
      End
      Begin VB.Label lblFiADCorte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   2880
         TabIndex        =   26
         Top             =   2640
         Width           =   1932
      End
   End
   Begin XtremeSuiteControls.GroupBox fraDuplicados 
      Height          =   4575
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   8655
      _Version        =   1441793
      _ExtentX        =   15261
      _ExtentY        =   8064
      _StockProps     =   79
      Caption         =   "Solicitudes que peresentan transacciones similares"
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
      Begin XtremeSuiteControls.ListView lswDuplicados 
         Height          =   3732
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   8412
         _Version        =   1441793
         _ExtentX        =   14838
         _ExtentY        =   6583
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
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         FlatScrollBar   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnCerrar_Duplicados 
         Height          =   432
         Left            =   7320
         TabIndex        =   23
         Top             =   120
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         Picture         =   "frmTES_Autorizacion.frx":720F
      End
   End
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   852
      Left            =   240
      TabIndex        =   44
      Top             =   7440
      Width           =   10452
      _Version        =   1441793
      _ExtentX        =   18436
      _ExtentY        =   1503
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   552
         Left            =   8880
         TabIndex        =   45
         Top             =   240
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   974
         _StockProps     =   79
         Caption         =   "&Autorizar"
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
         Picture         =   "frmTES_Autorizacion.frx":7BCC
      End
      Begin XtremeSuiteControls.FlatEdit txtCasos 
         Height          =   312
         Left            =   840
         TabIndex        =   46
         Top             =   360
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   2640
         TabIndex        =   47
         Top             =   360
         Width           =   2052
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtContraseña 
         Height          =   312
         Left            =   6480
         TabIndex        =   50
         Top             =   360
         Width           =   2052
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Clave Autorizador"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   4800
         TabIndex        =   51
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label lblEnd1 
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
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   732
      End
      Begin VB.Label lblEnd2 
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
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   1800
         TabIndex        =   48
         Top             =   360
         Width           =   852
      End
   End
   Begin XtremeSuiteControls.CheckBox chkMarcas 
      Height          =   210
      Left            =   360
      TabIndex        =   43
      Top             =   2870
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   12735
      _Version        =   1441793
      _ExtentX        =   22463
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Filtros:"
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
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   435
         Left            =   7800
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "Buscar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTES_Autorizacion.frx":82E5
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnRangos 
         Height          =   435
         Left            =   9120
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "Rangos"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTES_Autorizacion.frx":89E5
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.ComboBox cboTipoCuenta 
         Height          =   312
         Left            =   2760
         TabIndex        =   11
         ToolTipText     =   "Tipo de Cuenta Bancaria"
         Top             =   960
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2355
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtSolicitudInicial 
         Height          =   312
         Left            =   1440
         TabIndex        =   13
         Top             =   600
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.FlatEdit txtSolicitudCorte 
         Height          =   312
         Left            =   2760
         TabIndex        =   14
         Top             =   600
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.FlatEdit txtToken 
         Height          =   330
         Left            =   1440
         TabIndex        =   15
         Top             =   960
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   1440
         TabIndex        =   16
         Top             =   240
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
         Left            =   2760
         TabIndex        =   17
         Top             =   240
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
      Begin XtremeSuiteControls.CheckBox chkFechas 
         Height          =   252
         Left            =   4200
         TabIndex        =   18
         Top             =   240
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
      Begin XtremeSuiteControls.CheckBox chkSolicitudes 
         Height          =   252
         Left            =   4200
         TabIndex        =   19
         Top             =   600
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
      Begin XtremeSuiteControls.CheckBox chkCasosDuplicados 
         Height          =   255
         Left            =   10800
         TabIndex        =   20
         Top             =   240
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Verifica Duplicados ?"
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
      Begin XtremeSuiteControls.CheckBox chkBloqueos 
         Height          =   252
         Left            =   7920
         TabIndex        =   52
         Top             =   600
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "&Ver Casos Bloqueados ?"
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
      End
      Begin XtremeSuiteControls.CheckBox chkCuentaVerifica 
         Height          =   255
         Left            =   5520
         TabIndex        =   53
         Top             =   240
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Verifica Cuentas ?"
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
      Begin XtremeSuiteControls.CheckBox chkBancoCuentas 
         Height          =   255
         Left            =   7920
         TabIndex        =   54
         Top             =   240
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "&Ver Cuentas del mismo Banco?"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   1320
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10610
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
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   435
         Left            =   10320
         TabIndex        =   60
         ToolTipText     =   "Exportar Resultados"
         Top             =   1200
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   767
         _StockProps     =   79
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTES_Autorizacion.frx":90FE
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtAppCod 
         Height          =   315
         Left            =   5520
         TabIndex        =   62
         Top             =   960
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
      Begin XtremeSuiteControls.CheckBox chkSinpeCtas 
         Height          =   255
         Left            =   5520
         TabIndex        =   63
         Top             =   600
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Consulta Cuentas  Sinpe?"
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
         Appearance      =   16
         Value           =   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "App.Id:"
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
         Left            =   4200
         TabIndex        =   61
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Detalle:"
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
         Left            =   240
         TabIndex        =   59
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Fechas :"
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
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   972
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Solicitudes:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Token:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1092
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Autorizacion.frx":9268
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Autorizacion.frx":9582
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Autorizacion.frx":9E5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Autorizacion.frx":106BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Autorizacion.frx":16F20
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Autorizacion.frx":1D782
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Autorizacion.frx":23FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Autorizacion.frx":2A846
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   1
      Top             =   8370
      Width           =   14280
      _ExtentX        =   25188
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboDoc 
      Height          =   312
      Left            =   6240
      TabIndex        =   9
      Top             =   480
      Width           =   3012
      _Version        =   1441793
      _ExtentX        =   5318
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1560
      TabIndex        =   10
      Top             =   480
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
   Begin XtremeSuiteControls.PushButton btnAutorizacion 
      Height          =   312
      Index           =   0
      Left            =   1560
      TabIndex        =   55
      Top             =   120
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Emisión de Documentos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Checked         =   -1  'True
      Picture         =   "frmTES_Autorizacion.frx":310A8
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnAutorizacion 
      Height          =   312
      Index           =   1
      Left            =   4080
      TabIndex        =   56
      Top             =   120
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Firma Electrónica"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmTES_Autorizacion.frx":316DA
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.ComboBox cboTipoTS 
      Height          =   330
      Left            =   9480
      TabIndex        =   64
      Top             =   480
      Visible         =   0   'False
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
   Begin XtremeSuiteControls.Label lblTS 
      Height          =   255
      Left            =   9480
      TabIndex        =   65
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo de Transferencia"
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Autoriza:"
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
      TabIndex        =   58
      Top             =   120
      Width           =   1092
   End
   Begin XtremeShortcutBar.ShortcutCaption lblSolicitudes 
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   2760
      Width           =   10095
      _Version        =   1441793
      _ExtentX        =   17801
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Solicitudes Pendientes de Autorización"
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   972
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaptionTitle 
      Height          =   915
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   12735
      _Version        =   1441793
      _ExtentX        =   22463
      _ExtentY        =   1614
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
   End
End
Attribute VB_Name = "frmTES_Autorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mGrupoBancario As String

Private Sub btnAutorizacion_Click(Index As Integer)

btnAutorizacion.Item(0).Checked = False
btnAutorizacion.Item(1).Checked = False

btnAutorizacion.Item(Index).Checked = True

Call cbo_Click

End Sub

Private Sub btnBuscar_Click()
    Call sbBuscar
End Sub

Private Sub btnCerrar_Duplicados_Click()
If fraDuplicados.Visible Then
   fraDuplicados.Visible = False
End If
lsw.Visible = True
End Sub

Private Sub btnCerrar_Rangos_Click()
     If fraRangos.Visible Then
        fraRangos.Visible = False
        lsw.Visible = True
     Else
        fraRangos.Visible = True
        lsw.Visible = False
     End If
End Sub

Private Sub btnExportar_Click()
Call Excel_Exportar_Lsw(lsw)
End Sub

Private Sub btnRangos_Click()
     If fraRangos.Visible Then
        fraRangos.Visible = False
        lsw.Visible = True
     Else
        fraRangos.Visible = True
        lsw.Visible = False
     End If
End Sub

Private Sub cboDoc_Click()
txtCasos = 0
txtMonto = 0

lsw.ListItems.Clear

If cboDoc.ListCount > 0 Then
    If cboDoc.ItemData(cboDoc.ListIndex) = "TS" Then
        lblTS.Visible = True
        cboTipoTS.Visible = True
    Else
        lblTS.Visible = False
        cboTipoTS.Visible = False
    End If
End If

End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkMarcas_Click()
Dim i As Integer

vPaso = True

txtCasos = 0
txtMonto = 0

For i = 1 To lsw.ListItems.Count
 lsw.ListItems.Item(i).Checked = chkMarcas.Value
 
 If lsw.ListItems.Item(i).Checked Then
    txtMonto = CCur(txtMonto) + CCur(lsw.ListItems.Item(i).SubItems(3))
    txtCasos = CCur(txtCasos) + 1
 End If
 
Next i

txtCasos = Format(txtCasos, "###,###,###,##0")
txtMonto = Format(txtMonto, "Standard")

vPaso = False


End Sub

Private Sub chkSolicitudes_Click()

If chkSolicitudes.Value = vbChecked Then
   txtSolicitudInicial.Enabled = False
Else
   txtSolicitudInicial.Enabled = True
End If

txtSolicitudCorte.Enabled = txtSolicitudInicial.Enabled

End Sub

Private Sub sbGrupoBancario(pBancoId As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select dbo.fxTes_BancoSFN(" & pBancoId & ") as 'Codigo'"
Call OpenRecordSet(rs, strSQL)
  mGrupoBancario = rs!Codigo
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cbo_Click()

If vPaso Then Exit Sub

If cbo.ListCount = 0 Then Exit Sub

Call sbGrupoBancario(cbo.ItemData(cbo.ListIndex))

If btnAutorizacion.Item(0).Checked = True Then
    Call sbTesTiposDocsCargaCboAcceso(cboDoc, glogon.Usuario, cbo.ItemData(cbo.ListIndex), "A")
Else
    Call sbTesTiposDocsCargaCboAccesoFirmas(cboDoc, glogon.Usuario, cbo.ItemData(cbo.ListIndex), "A")
End If

Call sbAutorizacionInfo("B")

txtCasos = 0
txtMonto = 0

lsw.ListItems.Clear

End Sub


Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError


If Trim(txtContraseña) = "" Then
   MsgBox "No se puede Autorizar" & vbCrLf & "Suministre La Contraseña De Autorización", vbExclamation, "Faltan Datos"
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "Select * From Tes_Autorizaciones Where Clave='" _
       & fxTESCifrado(Trim(txtContraseña)) & "' and nombre = '" & glogon.Usuario _
       & "' and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
     MsgBox "No se puede Autorizar", vbExclamation, "Contraseña Incorrecta, o no Existe Nivel de Autorización"
     rs.Close
     Me.MousePointer = vbDefault
     Exit Sub
End If
rs.Close
   
   
PrgBar.Max = lsw.ListItems.Count + 1
PrgBar.Value = 1

PrgBar.Visible = True
strSQL = ""
For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked Then
  
     If btnAutorizacion.Item(0).Checked = True Then
       'Emision
            strSQL = strSQL & Space(10) & "Update Tes_Transacciones set Autoriza='S', Fecha_Autorizacion = dbo.MyGetdate()" _
                   & ", User_Autoriza = '" & glogon.Usuario _
                   & "' Where Nsolicitud = " & lsw.ListItems.Item(i).Text
            
            strSQL = strSQL & Space(10) & "exec spTesBitacora " & lsw.ListItems.Item(i).Text & ",'02','','" & glogon.Usuario & "'"
            
      Else
        'Firmas
            strSQL = strSQL & Space(10) & "Update Tes_Transacciones set FIRMAS_AUTORIZA_FECHA = dbo.MyGetdate()" _
                   & ", FIRMAS_AUTORIZA_USUARIO = '" & glogon.Usuario _
                   & "' Where Nsolicitud = " & lsw.ListItems.Item(i).Text
            
            strSQL = strSQL & Space(10) & "exec spTesBitacora " & lsw.ListItems.Item(i).Text & ",'04','','" & glogon.Usuario & "'"
      End If
      
      If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
      End If
                      
   End If
   
   PrgBar.Value = PrgBar.Value + 1
   
Next i

'Lote Final
If Len(strSQL) > 0 Then
      Call ConectionExecute(strSQL)
      strSQL = ""
End If


PrgBar.Visible = False
Call sbBuscar

Me.MousePointer = vbDefault


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSolicitudInicial.SetFocus
vError:
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
vError:
End Sub


Private Sub sbAutorizacionInfo(pTipo As String)
Dim strSQL As String, rs As New ADODB.Recordset


If pTipo = "G" Then
    lblEmAGInicio.Caption = "0.00"
    lblEmAGCorte.Caption = "0.00"
    lblEmADInicio.Caption = "0.00"
    lblEmADCorte.Caption = "0.00"
    
    lblFiAGInicio.Caption = "0.00"
    lblFiAGCorte.Caption = "0.00"
    lblFiADInicio.Caption = "0.00"
    lblFiADCorte.Caption = "0.00"
    
   strSQL = "select rango_gen_Inicio,rango_gen_corte,firmas_gen_inicio,firmas_gen_corte" _
          & " From TES_AUTORIZACIONES where NOMBRE = '" & glogon.Usuario & "'"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
        lblEmAGInicio.Caption = Format(rs!rango_gen_Inicio, "Standard")
        lblEmAGCorte.Caption = Format(rs!rango_gen_corte, "Standard")
        lblFiAGInicio.Caption = Format(rs!firmas_gen_inicio, "Standard")
        lblFiAGCorte.Caption = Format(rs!firmas_gen_corte, "Standard")
   End If
   rs.Close

Else
    lblEmADInicio.Caption = "0.00"
    lblEmADCorte.Caption = "0.00"
    
    lblFiADInicio.Caption = "0.00"
    lblFiADCorte.Caption = "0.00"
    
   strSQL = "select firmas_autoriza_inicio,firmas_autoriza_corte from TES_BANCO_FIRMASAUT" _
          & " where USUARIO = '" & glogon.Usuario & "' and ID_BANCO = " & cbo.ItemData(cbo.ListIndex) _
          & "  and aplica_rango_autorizacion = 1"

   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
        lblFiADInicio.Caption = Format(rs!firmas_autoriza_inicio, "Standard")
        lblFiADCorte.Caption = Format(rs!firmas_autoriza_corte, "Standard")
   End If
   rs.Close

End If



End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
vModulo = 9

mGrupoBancario = ""

With lsw.ColumnHeaders
  .Clear
  .Add , , "No.Solicitud", 1200
  .Add , , "Código", 1500, vbCenter
  .Add , , "Beneficiario", 3500
  .Add , , "Monto", 1900, vbRightJustify
  .Add , , "Fecha", 2500, vbCenter
  .Add , , "Revisión?", 1000, vbCenter
  .Add , , "Cuenta Bancaria", 3200, vbCenter
  .Add , , "Detalle", 4200
  .Add , , "App.Id", 4200
End With

With lswDuplicados.ColumnHeaders
  .Clear
  .Add , , "No.Solicitud", 1200
  .Add , , "Código", 1500, vbCenter
  .Add , , "Beneficiario", 3500
  .Add , , "Monto", 1900, vbRightJustify
  .Add , , "Tipo", 1500, vbCenter
  .Add , , "Cuenta Bancaria", 3200, vbCenter
End With

cboTipoTS.AddItem "Crédito Directos"
cboTipoTS.AddItem "Tiempo Real"
cboTipoTS.Text = "Tiempo Real"


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio

Call chkFechas_Click
Call chkSolicitudes_Click

'Carga Información de Rangos de Autorización a Nivel de Usuario-General
Call sbAutorizacionInfo("G")

cboTipoCuenta.AddItem "Todas"
cboTipoCuenta.AddItem "Locales"
cboTipoCuenta.AddItem "InterBanca"
cboTipoCuenta.Text = "Todas"


vPaso = True
    Call sbTesBancoCargaCboAccesoGestion(cbo, glogon.Usuario, "Autoriza")
vPaso = False
Call cbo_Click

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub Form_Resize()

On Error Resume Next

ShortcutCaptionTitle.Width = Me.Width

lblSolicitudes.Width = Me.Width - 480
lsw.Width = lblSolicitudes.Width

lsw.Height = Me.Height - (lblSolicitudes.top + lblSolicitudes.Height + gbResumen.Height + 600)

gbFiltros.Width = lsw.Width
gbResumen.Width = lsw.Width
gbResumen.top = lsw.top + lsw.Height + 80

End Sub



Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, curMonto As Currency, itmX As ListViewItem
Dim bDuplicado As Boolean, mLInterBanca As Integer
Dim iSupervisa As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass

fraRangos.Visible = False
fraDuplicados.Visible = False

lsw.Visible = True

lsw.ListItems.Clear
i = 0
curMonto = 0
bDuplicado = False

''Verifica si el banco requiere supervisión de movimientos
strSQL = "Select SUPERVISION from tes_bancos where id_banco = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  iSupervisa = rs!SUPERVISION
Else
  iSupervisa = 0
End If
rs.Close



strSQL = "select Bg.LCTA_INTERNA, Bg.LCTA_INTERBANCARIA " _
       & " from TES_BANCOS Tb inner join TES_BANCOS_GRUPOS Bg on Tb.COD_GRUPO = Bg.COD_GRUPO" _
       & " Where Tb.ID_BANCO = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
   mLInterBanca = 17
Else
   mLInterBanca = rs!LCTA_InterBancaria
End If
rs.Close


'Revision con Ajuste Automatico
If chkCuentaVerifica.Value = xtpChecked Then
    strSQL = "exec spTes_Cuentas_Revision_Automatica " & cbo.ItemData(cbo.ListIndex)
    Call ConectionExecute(strSQL)
End If

'Consulta
If chkCasosDuplicados.Value = vbChecked Then
    strSQL = "select T.nsolicitud,T.codigo,T.beneficiario,T.monto,T.fecha_solicitud,T.cta_Ahorros" _
           & ", dbo.fxTesSupervisa(CODIGO,BENEFICIARIO,monto,0,'T') as 'duplicado'" _
           & ", dbo.fxTes_Cuenta_Verifica(T.id_banco,T.codigo,T.cta_ahorros) as 'Cta_Verifica'" _
           & ", T.Detalle1 + T.detalle2 as 'Detalle', isnull(T.cod_App,'') as 'AppId'" _
           & " from Tes_Transacciones T inner join Tes_Bancos B on T.id_banco = B.id_banco" _
           & " where T.estado = 'P' and B.id_banco = " & cbo.ItemData(cbo.ListIndex) _
           & " and T.Tipo = '" & cboDoc.ItemData(cboDoc.ListIndex) & "'"
Else
    strSQL = "select T.nsolicitud,T.codigo,T.beneficiario,T.monto,T.fecha_solicitud,T.cta_Ahorros" _
           & ",0 as 'duplicado'" _
           & ", dbo.fxTes_Cuenta_Verifica(T.id_banco,T.codigo,T.cta_ahorros) as 'Cta_Verifica'" _
           & ", T.Detalle1 + T.detalle2 as 'Detalle', isnull(T.cod_App,'') as 'AppId'" _
           & " from Tes_Transacciones T inner join Tes_Bancos B on T.id_banco = B.id_banco" _
           & " where T.estado = 'P' and B.id_banco = " & cbo.ItemData(cbo.ListIndex) _
           & " and T.Tipo = '" & cboDoc.ItemData(cboDoc.ListIndex) & "'"
End If

If chkFechas.Value = vbUnchecked Then
   strSQL = strSQL & " and T.fecha_solicitud between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
          & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

End If

If chkSolicitudes.Value = vbUnchecked Then
   strSQL = strSQL & " and (T.nsolicitud >= " & CCur(txtSolicitudInicial) & " and nsolicitud <=" _
          & CCur(txtSolicitudCorte) & ")"
End If


If chkBloqueos.Value = vbUnchecked Then
  strSQL = strSQL & " and T.fecha_hold is null"
End If


If cboDoc.ItemData(cboDoc.ListIndex) = "TE" Then

    Select Case Mid(cboTipoCuenta.Text, 1, 1)
      Case "T" 'Todas
      Case "L" 'Cuentas Locales
          strSQL = strSQL & " and len(rtrim(T.cta_Ahorros)) <> " & mLInterBanca
      Case "I" 'Cuentas Interbancarias
          strSQL = strSQL & " and len(rtrim(T.cta_Ahorros)) = " & mLInterBanca
    End Select
    
    
    'Filtra Cuentas del mismo Banco
    If chkBancoCuentas.Value = xtpChecked Then
          strSQL = strSQL & " and (SUBSTRING( rtrim(T.cta_Ahorros) , 1,10) like '%" & mGrupoBancario & "%'" _
                 & " and len(rtrim(T.cta_Ahorros)) = " & mLInterBanca & ")"
    End If

End If 'Transferencias


If btnAutorizacion.Item(0).Checked = True Then
   strSQL = strSQL & " and T.fecha_autorizacion is null and T.monto between " _
          & CCur(lblEmAGInicio.Caption) & " and " & CCur(lblEmAGCorte.Caption)
   
   If Trim(txtToken.Text) <> "" Then
      strSQL = strSQL & " and T.id_token = '" & txtToken.Text & "'"
   End If
   
Else
    strSQL = strSQL & " and T.FIRMAS_AUTORIZA_FECHA is null and T.monto > B.firmas_hasta" _
                    & " and dbo.fxTesAutorizaFirmaAcceso('" & glogon.Usuario & "'," & cbo.ItemData(cbo.ListIndex) & ",T.monto) = 1"
   
End If

 

If txtDetalle.Text <> "" Then
    Call sbSIFCleanTxtInject(txtDetalle)
    strSQL = strSQL & " and (T.DETALLE1 + T.DETALLE2) like '%" & Trim(txtDetalle.Text) & "%'"
End If

If txtAppCod.Text <> "" Then
    Call sbSIFCleanTxtInject(txtDetalle)
    strSQL = strSQL & " and isnull(T.COD_APP,'') like '%" & Trim(txtAppCod.Text) & "%'"
End If



Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
     
     If rs!Duplicado = 1 And iSupervisa = 1 Then
       itmX.ForeColor = vbRed
       bDuplicado = True
     End If
     
     
     itmX.SubItems(1) = RTrim(rs!Codigo)
     itmX.SubItems(2) = rs!Beneficiario
     itmX.SubItems(3) = Format(rs!Monto, "Standard")
     itmX.SubItems(4) = Format(rs!fecha_solicitud, "yyyy-mm-dd")
     itmX.SubItems(5) = rs!Duplicado
     itmX.SubItems(6) = rs!Cta_Ahorros & ""
     itmX.SubItems(7) = rs!Detalle & ""
     itmX.SubItems(8) = rs!AppId & ""
     
     itmX.Checked = chkMarcas.Value
     
     If itmX.Checked Then
        curMonto = curMonto + rs!Monto
        i = i + 1
     End If
     
     If rs!Cta_Verifica = 0 And cboDoc.ItemData(cboDoc.ListIndex) = "TE" Then
              itmX.ForeColor = vbRed
              itmX.Bold = True
              itmX.TextBackColor = RGB(250, 219, 216) 'Rojo
              itmX.SubItems(6) = "No Existe!"
     End If
     
 rs.MoveNext
Loop
rs.Close

txtCasos.Text = Format(i, "###,###,###,##0")
txtMonto.Text = Format(curMonto, "Standard")

If bDuplicado = True Then
  MsgBox "Solicitudes duplicadas marcadas en rojo...", vbInformation
  lsw.Visible = False
  fraDuplicados.Visible = True
  Call sbCargaDuplicados
End If


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error GoTo vError
'Sumar y Restar Totales
If vPaso Then Exit Sub

If Item.Checked Then
   txtCasos = CInt(txtCasos) + 1
   txtMonto = CCur(txtMonto) + CCur(Item.SubItems(3))
Else
   txtCasos = CInt(txtCasos) - 1
   txtMonto = CCur(txtMonto) - CCur(Item.SubItems(3))
End If

txtCasos = Format(txtCasos, "###,###,###,##0")
txtMonto = Format(txtMonto, "Standard")

vError:
End Sub


Private Sub lswDuplicados_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswDuplicados.SortKey = ColumnHeader.Index - 1
  If lswDuplicados.SortOrder = 0 Then lswDuplicados.SortOrder = 1 Else lswDuplicados.SortOrder = 0
  lswDuplicados.Sorted = True
End Sub


Private Sub txtSolicitudInicial_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSolicitudCorte.SetFocus
vError:
End Sub

Private Sub sbCargaDuplicados()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iDias As Integer, itmX As ListViewItem
Dim i As Integer

lswDuplicados.ListItems.Clear

strSQL = "select SUPERVISION_DIAS from TES_BANCOS where ID_BANCO =" & cbo.ItemData(cbo.ListIndex) & " "
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    iDias = rs!SUPERVISION_DIAS
Else
  iDias = 5
End If
rs.Close

For i = 1 To lsw.ListItems.Count
    If lsw.ListItems.Item(i).SubItems(5) = 1 Then
        strSQL = "select T.NSOLICITUD,T.CODIGO,T.BENEFICIARIO,T.MONTO,T.TIPO,B.DESCRIPCION from TES_TRANSACCIONES T" _
                & " inner join TES_BANCOS B on T.ID_BANCO = B.ID_BANCO  where (FECHA_EMISION  is null or AUTORIZA  = 'N') " _
                & " and dbo.fxTesSupervisa(CODIGO,BENEFICIARIO,monto," & cbo.ItemData(cbo.ListIndex) & ",'T')= 1 " _
                & " and T.fecha_solicitud between '" & Format(DateAdd("D", -5, dtpInicio.Value), "yyyy/mm/dd") _
                & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' and T.codigo = '" & lsw.ListItems(i).SubItems(1) & "'"
        
        Call OpenRecordSet(rs, strSQL)
        
        Do While Not rs.EOF
         Set itmX = lswDuplicados.ListItems.Add(, , rs!NSolicitud)
             itmX.SubItems(1) = rs!Codigo
             itmX.SubItems(2) = rs!Beneficiario
             itmX.SubItems(3) = Format(rs!Monto, "Standard")
             itmX.SubItems(4) = rs!Tipo
              itmX.SubItems(5) = rs!DESCRIPCION
         rs.MoveNext
        Loop
        rs.Close
    End If
Next i
End Sub

Private Sub txtToken_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = ""
  gBusquedas.Columna = "id_token"
  gBusquedas.Orden = "id_token"
  gBusquedas.Consulta = "select id_token,registro_fecha,estado from Tes_tokens"
  gBusquedas.Filtro = ""
  gBusquedas.Orden = "registro_fecha desc"
  frmBusquedas.Show vbModal
  txtToken.Text = gBusquedas.Resultado
End If

End Sub
