VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmTES_DesAutorizaciones 
   Caption         =   "Des-Autorizaciones de Solicitudes"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   11565
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3495
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   6495
      _Version        =   1310723
      _ExtentX        =   11451
      _ExtentY        =   6159
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
      Left            =   3720
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   5055
      _Version        =   1310723
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
         Height          =   432
         Left            =   3480
         TabIndex        =   16
         Top             =   3120
         Width           =   1332
         _Version        =   1310723
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
         Picture         =   "frmTES_DesAutorizaciones.frx":0000
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
         TabIndex        =   32
         Top             =   2640
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
         TabIndex        =   31
         Top             =   2640
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
         TabIndex        =   30
         Top             =   2280
         Width           =   1932
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
         TabIndex        =   28
         Top             =   2640
         Width           =   612
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
         TabIndex        =   27
         Top             =   2280
         Width           =   612
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
         TabIndex        =   26
         Top             =   1560
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
         TabIndex        =   25
         Top             =   1560
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
         TabIndex        =   24
         Top             =   1200
         Width           =   1932
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
         TabIndex        =   23
         Top             =   1200
         Width           =   1932
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
         TabIndex        =   22
         Top             =   1920
         Width           =   3852
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
         TabIndex        =   21
         Top             =   840
         Width           =   3852
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
         TabIndex        =   20
         Top             =   480
         Width           =   1932
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
         TabIndex        =   19
         Top             =   480
         Width           =   1932
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
         TabIndex        =   18
         Top             =   1560
         Width           =   612
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
         TabIndex        =   17
         Top             =   1200
         Width           =   612
      End
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   0
      Top             =   9060
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.GroupBox gbFiltros 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   10095
      _Version        =   1310723
      _ExtentX        =   17806
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipoCuenta 
         Height          =   312
         Left            =   2760
         TabIndex        =   2
         ToolTipText     =   "Tipo de Cuenta Bancaria"
         Top             =   960
         Width           =   1332
         _Version        =   1310723
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
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtSolicitudInicial 
         Height          =   312
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   1332
         _Version        =   1310723
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSolicitudCorte 
         Height          =   312
         Left            =   2760
         TabIndex        =   4
         Top             =   600
         Width           =   1332
         _Version        =   1310723
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtToken 
         Height          =   312
         Left            =   1440
         TabIndex        =   5
         Top             =   960
         Width           =   1332
         _Version        =   1310723
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   1332
         _Version        =   1310723
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
         TabIndex        =   7
         Top             =   240
         Width           =   1332
         _Version        =   1310723
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
         TabIndex        =   8
         Top             =   240
         Width           =   1092
         _Version        =   1310723
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
         TabIndex        =   9
         Top             =   600
         Width           =   972
         _Version        =   1310723
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
         Left            =   5520
         TabIndex        =   10
         Top             =   600
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Duplicados?"
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
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   435
         Left            =   7560
         TabIndex        =   43
         Top             =   1200
         Width           =   1215
         _Version        =   1310723
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
         TextAlignment   =   1
         Appearance      =   6
         Picture         =   "frmTES_DesAutorizaciones.frx":09BD
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnRangos 
         Height          =   435
         Left            =   8760
         TabIndex        =   44
         Top             =   1200
         Width           =   1215
         _Version        =   1310723
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
         TextAlignment   =   1
         Appearance      =   6
         Picture         =   "frmTES_DesAutorizaciones.frx":10BD
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.CheckBox chkBloqueos 
         Height          =   255
         Left            =   5520
         TabIndex        =   45
         Top             =   240
         Width           =   2415
         _Version        =   1310723
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
      Begin XtremeSuiteControls.FlatEdit txtAppCod 
         Height          =   315
         Left            =   5520
         TabIndex        =   56
         Top             =   960
         Width           =   1935
         _Version        =   1310723
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Left            =   1440
         TabIndex        =   58
         Top             =   1320
         Width           =   6015
         _Version        =   1310723
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
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
         Left            =   120
         TabIndex        =   59
         Top             =   1320
         Width           =   1095
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
         TabIndex        =   57
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Token.:"
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
         TabIndex        =   13
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Solicitudes.:"
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
         TabIndex        =   12
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Fechas .:"
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
         TabIndex        =   11
         Top             =   240
         Width           =   972
      End
   End
   Begin XtremeSuiteControls.GroupBox fraDuplicados 
      Height          =   4575
      Left            =   1440
      TabIndex        =   33
      Top             =   3000
      Visible         =   0   'False
      Width           =   8655
      _Version        =   1310723
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
         TabIndex        =   34
         Top             =   720
         Width           =   8412
         _Version        =   1310723
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
         TabIndex        =   35
         Top             =   120
         Width           =   1332
         _Version        =   1310723
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
         Picture         =   "frmTES_DesAutorizaciones.frx":17D6
      End
   End
   Begin XtremeSuiteControls.ComboBox cboDoc 
      Height          =   312
      Left            =   6120
      TabIndex        =   36
      Top             =   480
      Width           =   3012
      _Version        =   1310723
      _ExtentX        =   5318
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1440
      TabIndex        =   37
      Top             =   480
      Width           =   4692
      _Version        =   1310723
      _ExtentX        =   8281
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnAutorizacion 
      Height          =   312
      Index           =   0
      Left            =   1440
      TabIndex        =   38
      Top             =   120
      Width           =   2532
      _Version        =   1310723
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
      Picture         =   "frmTES_DesAutorizaciones.frx":2193
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnAutorizacion 
      Height          =   312
      Index           =   1
      Left            =   3960
      TabIndex        =   39
      Top             =   120
      Width           =   2532
      _Version        =   1310723
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
      Picture         =   "frmTES_DesAutorizaciones.frx":27C5
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   855
      Left            =   120
      TabIndex        =   46
      Top             =   7320
      Width           =   10455
      _Version        =   1310723
      _ExtentX        =   18436
      _ExtentY        =   1503
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   552
         Left            =   8880
         TabIndex        =   47
         Top             =   240
         Width           =   1812
         _Version        =   1310723
         _ExtentX        =   3196
         _ExtentY        =   974
         _StockProps     =   79
         Caption         =   "&Desautorizar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmTES_DesAutorizaciones.frx":2ECC
      End
      Begin XtremeSuiteControls.FlatEdit txtCasos 
         Height          =   312
         Left            =   840
         TabIndex        =   48
         Top             =   360
         Width           =   972
         _Version        =   1310723
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
         TabIndex        =   49
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
         PasswordChar    =   "*"
         Appearance      =   2
         UseVisualStyle  =   0   'False
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
         TabIndex        =   53
         Top             =   360
         Width           =   852
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
         TabIndex        =   52
         Top             =   360
         Width           =   732
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
   End
   Begin XtremeSuiteControls.CheckBox chkMarcas 
      Height          =   210
      Left            =   360
      TabIndex        =   54
      Top             =   2760
      Width           =   210
      _Version        =   1310723
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Caption         =   "CheckBox1"
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   2
   End
   Begin XtremeShortcutBar.ShortcutCaption lblSolicitudes 
      Height          =   375
      Left            =   120
      TabIndex        =   55
      Top             =   2640
      Width           =   10095
      _Version        =   1310723
      _ExtentX        =   17801
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Solicitudes Autorizadas"
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
      Caption         =   "Desautoriza:"
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
      Left            =   0
      TabIndex        =   42
      Top             =   120
      Width           =   1092
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
      Left            =   120
      TabIndex        =   40
      Top             =   480
      Width           =   972
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaptionTitle 
      Height          =   800
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   12732
      _Version        =   1310723
      _ExtentX        =   22458
      _ExtentY        =   1411
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
Attribute VB_Name = "frmTES_DesAutorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


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


For i = 1 To lsw.ListItems.Count
 lsw.ListItems.Item(i).Checked = chkMarcas.Value
Next i


End Sub

Private Sub chkSolicitudes_Click()

If chkSolicitudes.Value = vbChecked Then
   txtSolicitudInicial.Enabled = False
Else
   txtSolicitudInicial.Enabled = True
End If

txtSolicitudCorte.Enabled = txtSolicitudInicial.Enabled

End Sub

Private Sub cbo_Click()

If vPaso Then Exit Sub

If cbo.ListCount = 0 Then
   cbo.AddItem " "
   cbo.ItemData(cbo.NewIndex) = 0
   cbo.Text = " "
End If

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
   
   
Prgbar.Max = lsw.ListItems.Count + 1
Prgbar.Value = 1

Prgbar.Visible = True

strSQL = ""
For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked Then
  
  
     If btnAutorizacion.Item(0).Checked = True Then
       'Emision
      
            strSQL = strSQL & Space(10) & "Update Tes_Transacciones set Autoriza='N', Fecha_Autorizacion = Null" _
                   & ", User_Autoriza = Null Where Nsolicitud = " & lsw.ListItems.Item(i).Text
            
            strSQL = strSQL & Space(10) & "exec spTesBitacora " & lsw.ListItems.Item(i).Text & ",'03','Emisión de Documento','" & glogon.Usuario & "'"
    
     
      
      Else
        'Firmas
            strSQL = strSQL & Space(10) & "Update Tes_Transacciones set FIRMAS_AUTORIZA_FECHA = Null" _
                   & ", FIRMAS_AUTORIZA_USUARIO = Null Where Nsolicitud = " & lsw.ListItems.Item(i).Text
            
            strSQL = strSQL & Space(10) & "exec spTesBitacora " & lsw.ListItems.Item(i).Text & ",'03','Firmas Electrónicas','" & glogon.Usuario & "'"
          
      End If
      
      If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
      End If
                      
   End If
   
   Prgbar.Value = Prgbar.Value + 1
   
Next i
   
'Lote Final
If Len(strSQL) > 0 Then
      Call ConectionExecute(strSQL)
      strSQL = ""
End If

Prgbar.Visible = False
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

Me.Height = 9096
Me.Width = 12024

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio

With lsw.ColumnHeaders
  .Clear
  .Add , , "No.Solicitud", 1200
  .Add , , "Código", 1500, vbCenter
  .Add , , "Beneficiario", 3500
  .Add , , "Monto", 1900, vbRightJustify
  .Add , , "Fecha", 2500, vbCenter
  .Add , , "Duplicado?", 1000, vbCenter
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

cboTipoCuenta.AddItem "Todas"
cboTipoCuenta.AddItem "Locales"
cboTipoCuenta.AddItem "InterBanca"
cboTipoCuenta.Text = "Todas"


Call chkFechas_Click
Call chkSolicitudes_Click

'Carga Información de Rangos de Autorización a Nivel de Usuario-General
Call sbAutorizacionInfo("G")

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

lsw.Height = Me.Height - 4500

gbFiltros.Width = lsw.Width
gbResumen.Width = lsw.Width
gbResumen.top = lsw.top + lsw.Height + 80


End Sub




Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError
'Sumar y Restar Totales

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


Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, curMonto As Currency, itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear
i = 0
curMonto = 0


strSQL = "select T.nsolicitud,T.codigo,T.beneficiario,T.monto,T.fecha_solicitud,T.cta_Ahorros" _
    & ",0 as 'duplicado'" _
    & ", dbo.fxTes_Cuenta_Verifica(T.id_banco,T.codigo,T.cta_ahorros) as 'Cta_Verifica'" _
    & ", T.Detalle1 + T.detalle2 as 'Detalle', isnull(T.cod_App,'') as 'AppId'" _
    & " from Tes_Transacciones T inner join Tes_Bancos B on T.id_banco = B.id_banco" _
    & " where T.estado = 'P' and B.id_banco = " & cbo.ItemData(cbo.ListIndex) _
    & " and T.Tipo = '" & cboDoc.ItemData(cboDoc.ListIndex) & "'"

       
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

If btnAutorizacion.Item(0).Checked = True Then
   strSQL = strSQL & " and T.fecha_autorizacion is not null and T.monto between " _
          & CCur(lblEmAGInicio.Caption) & " and " & CCur(lblEmAGCorte.Caption)
Else
   strSQL = strSQL & " and T.FIRMAS_AUTORIZA_FECHA is not null and T.monto > B.firmas_hasta"
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
     
     itmX.Checked = chkMarcas.Value
     
     itmX.SubItems(1) = RTrim(rs!Codigo)
     itmX.SubItems(2) = rs!Beneficiario
     itmX.SubItems(3) = Format(rs!Monto, "Standard")
     itmX.SubItems(4) = Format(rs!fecha_solicitud, "yyyy-mm-dd")
     itmX.SubItems(5) = rs!Duplicado
     itmX.SubItems(6) = rs!Cta_Ahorros & ""
     itmX.SubItems(7) = rs!Detalle & ""
     itmX.SubItems(8) = rs!AppId & ""
     
     
     If itmX.Checked Then
        curMonto = curMonto + rs!Monto
        i = i + 1
     End If
     
 rs.MoveNext
Loop
rs.Close

txtCasos = Format(i, "###,###,###,##0")
txtMonto = Format(curMonto, "Standard")

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
    Call sbBuscar
  Case "Rangos"
     If fraRangos.Visible Then
        fraRangos.Visible = False
     Else
        fraRangos.Visible = True
     End If
End Select

End Sub

Private Sub tlbRango_ButtonClick(ByVal Button As MSComctlLib.Button)
     If fraRangos.Visible Then
        fraRangos.Visible = False
     Else
        fraRangos.Visible = True
     End If
End Sub

Private Sub txtSolicitudInicial_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSolicitudCorte.SetFocus
vError:
End Sub


