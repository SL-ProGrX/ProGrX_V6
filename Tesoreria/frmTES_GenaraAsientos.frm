VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmTES_GeneraAsientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traslado de Asientos a contabilidad"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1013
   Icon            =   "frmTES_GenaraAsientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   11175
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4332
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   10932
      _Version        =   1441793
      _ExtentX        =   19283
      _ExtentY        =   7641
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Frame fraReportes 
      Caption         =   "Reportes de Traslados"
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
      Height          =   2412
      Left            =   6120
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   4935
      Begin XtremeSuiteControls.CheckBox chkRepTodas 
         Height          =   252
         Left            =   2520
         TabIndex        =   27
         Top             =   960
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas las Fechas"
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
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpRepInicio 
         Height          =   312
         Left            =   2040
         TabIndex        =   23
         Top             =   480
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
      Begin XtremeSuiteControls.DateTimePicker dtpRepCorte 
         Height          =   312
         Left            =   3360
         TabIndex        =   24
         Top             =   480
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
      Begin XtremeSuiteControls.PushButton cmdRepAplica 
         Height          =   312
         Left            =   1320
         TabIndex        =   25
         Top             =   1920
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   7
         Picture         =   "frmTES_GenaraAsientos.frx":6852
      End
      Begin XtremeSuiteControls.PushButton cmdRepCerrar 
         Height          =   312
         Left            =   3120
         TabIndex        =   26
         Top             =   1920
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Cerrar"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   7
         Picture         =   "frmTES_GenaraAsientos.frx":6F59
      End
      Begin XtremeSuiteControls.CheckBox chkRepBancoActual 
         Height          =   252
         Left            =   2520
         TabIndex        =   28
         Top             =   1320
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Solo el la Cuenta actual "
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
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Rango de Fechas:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1812
      End
   End
   Begin VB.Timer Timer1 
      Left            =   10440
      Top             =   240
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   8055
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prgGenera 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   0
      Top             =   7635
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbOperacion 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   7800
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9878
            MinWidth        =   9878
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   672
      Left            =   8400
      TabIndex        =   17
      Top             =   960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   1182
      _StockProps     =   79
      Caption         =   "&Buscar"
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
      Picture         =   "frmTES_GenaraAsientos.frx":767A
   End
   Begin XtremeSuiteControls.PushButton cmdReportes 
      Height          =   672
      Left            =   9720
      TabIndex        =   18
      Top             =   960
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   1182
      _StockProps     =   79
      Caption         =   "&Reportes"
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
      Picture         =   "frmTES_GenaraAsientos.frx":8098
   End
   Begin XtremeSuiteControls.PushButton cmdGenerar 
      Height          =   672
      Left            =   9720
      TabIndex        =   19
      Top             =   6720
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   1182
      _StockProps     =   79
      Caption         =   "&Generar"
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
      Picture         =   "frmTES_GenaraAsientos.frx":8854
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1200
      TabIndex        =   21
      Top             =   960
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
      Left            =   1200
      TabIndex        =   22
      Top             =   1320
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2640
      TabIndex        =   29
      Top             =   120
      Width           =   6492
      _Version        =   1441793
      _ExtentX        =   11456
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   2640
      TabIndex        =   30
      Top             =   480
      Width           =   6492
      _Version        =   1441793
      _ExtentX        =   11456
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkBancosTodos 
      Height          =   204
      Left            =   9240
      TabIndex        =   31
      Top             =   120
      Width           =   204
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   16
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkDocumentosTodos 
      Height          =   204
      Left            =   9240
      TabIndex        =   32
      Top             =   480
      Width           =   204
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   16
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton btnEstado 
      Height          =   672
      Index           =   0
      Left            =   2640
      TabIndex        =   33
      Top             =   960
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   1185
      _StockProps     =   79
      Caption         =   "Emitidos"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   7
      Picture         =   "frmTES_GenaraAsientos.frx":9059
   End
   Begin XtremeSuiteControls.PushButton btnEstado 
      Height          =   672
      Index           =   1
      Left            =   4080
      TabIndex        =   34
      Top             =   960
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   1185
      _StockProps     =   79
      Caption         =   "Anulados"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   7
      Picture         =   "frmTES_GenaraAsientos.frx":9837
   End
   Begin XtremeSuiteControls.PushButton btnEstado 
      Height          =   672
      Index           =   2
      Left            =   5520
      TabIndex        =   35
      Top             =   960
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   1185
      _StockProps     =   79
      Caption         =   "Anulados [Emitidos en meses anteriores]"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   7
      Picture         =   "frmTES_GenaraAsientos.frx":A1CC
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   204
      Left            =   240
      TabIndex        =   36
      Top             =   1860
      Width           =   204
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   16
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption lblBanco 
      Height          =   372
      Left            =   120
      TabIndex        =   37
      Top             =   1800
      Width           =   10932
      _Version        =   1441793
      _ExtentX        =   19283
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asientos con Error"
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
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   16
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asientos a Trasladar"
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
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   15
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label lblAsientosError 
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label lblMontoError 
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
      ForeColor       =   &H000000C0&
      Height          =   252
      Left            =   4800
      TabIndex        =   13
      Top             =   8040
      Width           =   2292
   End
   Begin VB.Label lblAsientos 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Número de Asientos"
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
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label lblMonto 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total (Monto)"
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
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Label lblMontoGenerar 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Documento...:"
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
      Height          =   312
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   2052
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Bancaria..:"
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
      Height          =   312
      Index           =   5
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   1812
   End
   Begin VB.Label lblAsientosGenerar 
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
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11892
   End
End
Attribute VB_Name = "frmTES_GeneraAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Function fxAsientoNumeroValido(pTipo As String, pNumero As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select count(*) as Existe from CntX_Asientos where cod_contabilidad = " & GLOBALES.gEnlace _
     & " and Tipo_Asiento = '" & pTipo & "' and Num_asiento = '" & pNumero & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe > 0 Then
  pNumero = pNumero & "v" & (rs!Existe + 1)
End If

fxAsientoNumeroValido = pNumero
End Function

Private Sub sbAsientoEmisionTransac()

Dim strSQL As String, i As Long
Dim pNTesoreria As Long, pDocumento As String

On Error GoTo vError

 
'Eliminando Cuentas sin mapeo en Contabilidad

strSQL = "delete Asi " _
       & " from TES_TRANSACCIONES Tra inner join TES_TRANS_ASIENTO Asi on Tra.Nsolicitud = Asi.NSOLICITUD" _
       & " where Asi.CUENTA_CONTABLE not in(select cod_cuenta from CNTX_CUENTAS) " _
       & " and Tra.Estado_Asiento = 'P' And Tra.Estado in('T','I') and Tra.fecha_emision between '" _
       & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Call ConectionExecute(strSQL)
 
 
With lsw.ListItems
 
    For i = 1 To .Count
     If .Item(i).Checked Then
       prgGenera.Max = prgGenera.Max + 1
     End If
    Next i


    strSQL = ""
    For i = 1 To .Count
    
     pNTesoreria = .Item(i).Text
     pDocumento = .Item(i).SubItems(1)
        
     If .Item(i).Checked And Len(Trim(pDocumento)) > 0 Then
        strSQL = strSQL & Space(10) & "exec spTES_Asientos_Traslado_Individual " & pNTesoreria & ",'" & glogon.Usuario & "'"
        prgGenera.Value = prgGenera.Value + 1
        
        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
     End If 'Item CHECK
         
    Next i
    
    'Ultimo Lote
    If Len(strSQL) > 0 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If
    
    prgGenera.Value = 0

End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub btnEstado_Click(Index As Integer)
Dim i As Integer

For i = 0 To btnEstado.Count - 1
   btnEstado.Item(i).Checked = False
Next i
btnEstado.Item(Index).Checked = True

lblBanco.Visible = True
chkTodos.Visible = True
lsw.Visible = True

fraReportes.Visible = False

Select Case Index
    Case 0 'Emitidos
        Call sbConsulta_Emitidos
    Case 1 'Anulados
        Call sbConsulta_Anulados
    Case 2 'Anulados en Periodos Anteriores
        Call sbConsulta_Anulados_Anteriores
End Select

End Sub

Private Sub cbo_Click()

If vPaso Then Exit Sub

If cbo.ListCount = 0 Then
   cbo.AddItem " "
   cbo.ItemData(cbo.NewIndex) = 0
   cbo.Text = " "
End If

Call sbTesTiposDocsCargaCboAcceso(cboTipo, glogon.Usuario, cbo.ItemData(cbo.ListIndex), "X")

lblBanco.Caption = cbo.Text

Call btnEstado_Click(0)

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select id_Banco,descripcion from Tes_Bancos"
  gBusquedas.Filtro = " and Aplica_Cheques = 1"
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado2 <> "" Then cbo.Text = Trim(gBusquedas.Resultado2)
  Call cbo_Click
End If
End Sub

Private Sub cboTipo_Click()

If vPaso Then Exit Sub

Call btnEstado_Click(0)


End Sub

Private Sub chkBancosTodos_Click()
If chkBancosTodos.Value = vbChecked Then
   cbo.Enabled = False
Else
   cbo.Enabled = True
End If
End Sub

Private Sub chkDocumentosTodos_Click()
If chkDocumentosTodos.Value = vbChecked Then
   cboTipo.Enabled = False
Else
   cboTipo.Enabled = True
End If
End Sub

Private Sub chkRepTodas_Click()
If chkRepTodas.Value = vbChecked Then
  dtpRepCorte.Enabled = False
  dtpRepInicio.Enabled = False
Else
  dtpRepCorte.Enabled = True
  dtpRepInicio.Enabled = True
End If

End Sub

Private Sub chkTodos_Click()
Dim i As Integer

lblAsientosGenerar.Caption = 0
lblMontoGenerar.Caption = 0

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
  lblAsientosGenerar.Caption = CInt(lblAsientosGenerar.Caption) + 1
  lblMontoGenerar.Caption = Format(CCur(lblMontoGenerar.Caption) + CCur(lsw.ListItems.Item(i).SubItems(2)), "Standard")
Next i

If chkTodos.Value = vbUnchecked Then
    lblAsientosGenerar.Caption = 0
    lblMontoGenerar.Caption = 0
End If

End Sub

Private Sub cmdBuscar_Click()
Select Case True
  Case btnEstado.Item(0).Checked
     Call sbConsulta_Emitidos
  
  Case btnEstado.Item(1).Checked
     Call sbConsulta_Anulados
     
  Case btnEstado.Item(2).Checked
     Call sbConsulta_Anulados_Anteriores

  Case Else
     lsw.ListItems.Clear
End Select

End Sub

Private Sub cmdGenerar_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Genera a Contabilidad las solicitudes emitidas y las anuladas durante el mes.
'               Las solicitudes anuladas en meses anteriores no las emite a Contabilidad,
'               sino que las despliega en un reporte.
'REFERENCIAS:   sbAsientoEmisionTransac - (Genera a Contabilidad las solicitudes emitidas)
'               GeneraAnulados - (Genera a Contabilidad las solicitudes anuladas durante el
'               mes)
'               fxFechaServidor - (Devuelve la fecha del servidor)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
    Case btnEstado.Item(0).Checked 'Emitidos
   
            stbOperacion.Panels(1).Text = "Generando Emitidos..........................."
            
            Call sbAsientoEmisionTransac
            
            stbOperacion.Panels(1).Text = ""
             
            Call sbConsulta_Emitidos

    Case btnEstado.Item(1).Checked 'Anulados
              stbOperacion.Panels(1).Text = "Generando Anulados..........................."
            '  Call GeneraAnulados
              stbOperacion.Panels(1).Text = ""
              
              Call sbConsulta_Anulados

    Case btnEstado.Item(2).Checked 'Anulados anteriormente
           With frmContenedor.Crt
            .Reset
            .WindowShowGroupTree = True
            .WindowShowRefreshBtn = True
            .WindowShowPrintSetupBtn = True
            .WindowState = crptMaximized
            .WindowShowSearchBtn = True
            .WindowTitle = "Reportes Módulo de Banking"
            
            .Connect = glogon.ConectRPT
            
            .Formulas(0) = "Fecha='" & Format(fxFechaServidor, ("dd/mm/yyyy")) & "'"
            .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
            
            .ReportFileName = SIFGlobal.fxPathReportes("Banking_AsientosAnulados.rpt")
            strSQL = "{CHEQUES.ID_BANCO} =" & cbo.ItemData(cbo.ListIndex) & " And "
            strSQL = strSQL & "{CHEQUES.TIPO}='" & fxCodigoCbo(cboTipo) & "' And "
            strSQL = strSQL & "{CHEQUES.ESTADO_ASIENTO}='P' And {CHEQUES.ESTADO}='A'"
            strSQL = strSQL & " And {CHEQUES.FECHA_ANULA} < Date("
            strSQL = strSQL & Year(fxFechaServidor) & ","
            strSQL = strSQL & Format(Month(fxFechaServidor), "00") & "," & "01)"
        
            .SelectionFormula = strSQL
        
            .PrintReport
           End With
End Select

Me.MousePointer = vbDefault

Exit Sub
vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdRepAplica_Click()
Dim strSQL As String

Me.MousePointer = vbHourglass


With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Banking"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "Fecha='" & Format(fxFechaServidor, ("dd/mm/yyyy")) & "'"
    .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    
    strSQL = ""
    
    If chkRepTodas.Value = vbChecked Then
        .Formulas(1) = "Rango='Todas las Fechas'"
    Else
        strSQL = "CDATE({CHEQUES.FECHA_ASIENTO}) in Date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd") _
               & ") to Date(" & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ")"
        .Formulas(1) = "Rango='Del  " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & "  Al  " & Format(dtpRepCorte.Value, "dd/mm/yyyy") & "'"
    End If
    
    .ReportFileName = SIFGlobal.fxPathReportes("Banking_AsientosGenerados.rpt")

 

    If chkRepBancoActual.Value = vbChecked Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{BANCOS.ID_BANCO} = " & cbo.ItemData(cbo.ListIndex)
    End If
    .SelectionFormula = strSQL
    .PrintReport
End With


Me.MousePointer = vbDefault

End Sub

Private Sub cmdRepCerrar_Click()
lblBanco.Visible = True
chkTodos.Visible = True
lsw.Visible = True

fraReportes.Visible = False

End Sub

Private Sub cmdReportes_Click()

If fraReportes.Visible Then

    lblBanco.Visible = True
    chkTodos.Visible = True
    lsw.Visible = True
    
    fraReportes.Visible = False
Else
    lblBanco.Visible = False
    chkTodos.Visible = False
    lsw.Visible = False
    
    fraReportes.Visible = True
    
    dtpRepInicio.Value = fxFechaServidor
    dtpRepCorte.Value = dtpRepInicio.Value

End If

End Sub

Private Sub Form_Activate()
 vModulo = 9

End Sub


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 9

On Error GoTo vError

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "No.Solicitud", 1500
    .Add , , "No.Documento", 2100
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Beneficiario", 4500
    .Add , , "Tipo", 1100, vbCenter
    .Add , , "Cuenta", 4000
End With


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

Me.Picture = Me.Icon

vPaso = True
Call sbTesBancoCargaCboAccesoGestion(cbo, glogon.Usuario, "Asientos")
vPaso = False
Call cbo_Click

lblBanco.Caption = cbo.Text

Call Formularios(Me)
Call RefrescaTags(Me)

Timer1.Interval = 1

Call chkBancosTodos_Click
Call chkDocumentosTodos_Click

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If Item.Checked Then
  lblAsientosGenerar.Caption = CInt(lblAsientosGenerar.Caption) + 1
  lblMontoGenerar.Caption = Format(CCur(lblMontoGenerar.Caption) + CCur(Item.SubItems(2)), "Standard")
Else
  lblAsientosGenerar.Caption = CInt(lblAsientosGenerar.Caption) - 1
  lblMontoGenerar.Caption = Format(CCur(lblMontoGenerar.Caption) - CCur(Item.SubItems(2)), "Standard")
End If

vError:
End Sub

Private Sub sbConsulta_Anulados()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Despliega en pantalla las solicitudes cuyo estado de asiento esta pendiente
'               y cuya fecha de asiento es nula, y a la vez tienen el estado de la solicitud
'               Anulado y la fecha de anulacion esta dentro del mes Actual.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, lng As Long, curMonto As Currency
Dim lngErr As Long, curMontoErr As Currency



'    strSQL = "Select * FROM Tes_Transacciones where Tipo='" & fxCodigoCbo(cboTipo) & "'"
'    strSQL = strSQL & " And Id_Banco=" & cbo.ItemData(cbo.ListIndex) & " And Estado_Asiento='P'"
'    strSQL = strSQL & " And Estado='A' And DatePart(month,Fecha_Anula)=" & Month(fxFechaServidor)
'    strSQL = strSQL & " And DatePart(year,Fecha_Anula)=" & Year(fxFechaServidor)
'    strSQL = strSQL & " And DatePart(month,Fecha_Anula)=DatePart(month,Fecha_emision)"
'    strSQL = strSQL & " And DatePart(year,Fecha_Anula)=DatePart(year,Fecha_emision)"
'    strSQL = strSQL & " And Fecha_Asiento is not Null"
On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "Select T.Nsolicitud,T.Ndocumento,T.Monto,T.Fecha_Emision,T.Beneficiario,T.Tipo,B.descripcion as 'BancoDesc'" _
       & " FROM Tes_Transacciones T inner join Tes_Bancos B on T.id_Banco = B.id_Banco" _
       & " Where T.Estado_Asiento = 'P' And T.Fecha_Asiento is not Null And T.Estado = 'A'" _
       & " and T.Fecha_Anula between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"

If chkBancosTodos.Value = vbUnchecked Then
   strSQL = strSQL & " and T.id_Banco = " & cbo.ItemData(cbo.ListIndex)
End If

If chkDocumentosTodos.Value = vbUnchecked Then
   strSQL = strSQL & " and T.Tipo = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
End If

strSQL = strSQL & " Order by T.Id_Banco,T.Tipo,T.Fecha_Anula,T.NDocumento"

Call OpenRecordSet(rs, strSQL)

curMonto = 0
lng = 0

lngErr = 0
curMontoErr = 0
  
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
      itmX.SubItems(1) = rs!nDocumento & ""
      itmX.SubItems(2) = Format(rs!Monto, "Standard")
      itmX.SubItems(3) = rs!Fecha_Anula
      itmX.SubItems(4) = rs!Beneficiario
      itmX.SubItems(5) = rs!Tipo
      itmX.SubItems(6) = rs!BancoDesc
      
      itmX.Checked = chkTodos.Value
      
      If Trim(rs!nDocumento & "") = "" Then
         itmX.Bold = True
         itmX.ForeColor = vbRed
      
         curMontoErr = curMontoErr + rs!Monto
         lngErr = lngErr + 1
      Else
         curMonto = curMonto + rs!Monto
         lng = lng + 1
      End If
      
     
  rs.MoveNext
Loop
rs.Close
  
lblAsientosGenerar.Caption = lng
lblMontoGenerar.Caption = Format(curMonto, "Standard")

lblAsientosError.Caption = lngErr
lblMontoError.Caption = Format(curMontoErr, "Standard")


Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbConsulta_Anulados_Anteriores()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Despliega en pantalla las solicitudes cuyo estado de asiento esta pendiente
'               y a la vez tienen el estado de la solicitud Anulado y la fecha de anulacion
'               es menor al mes actual.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, lng As Long, curMonto As Currency
Dim lngErr As Long, curMontoErr As Currency
Dim vFecha As Date

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear
vFecha = fxFechaServidor

strSQL = "Select T.nsolicitud,T.ndocumento,T.monto,T.fecha_emision,T.beneficiario,T.Tipo,B.descripcion as 'BancoDesc'" _
       & " from Tes_Transacciones T inner join Tes_Bancos B on T.id_Banco = B.id_Banco" _
       & " where T.Estado_Asiento = 'P' And T.Estado = 'A' And T.Fecha_Anula < '" _
       & Year(vFecha) & "/" & Format(Month(vFecha), "00") & "/01'"

If chkBancosTodos.Value = vbUnchecked Then
   strSQL = strSQL & " and T.id_Banco = " & cbo.ItemData(cbo.ListIndex)
End If

If chkDocumentosTodos.Value = vbUnchecked Then
   strSQL = strSQL & " and T.Tipo = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
End If

strSQL = strSQL & " Order by T.Id_Banco,T.Tipo,T.Fecha_Emision,T.NDocumento"


curMonto = 0
lng = 0

lngErr = 0
curMontoErr = 0

Call OpenRecordSet(rs, strSQL)
 
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
      itmX.SubItems(1) = rs!nDocumento
      itmX.SubItems(2) = Format(rs!Monto, "Standard")
      itmX.SubItems(3) = rs!Fecha_Emision
      itmX.SubItems(4) = rs!Beneficiario
      itmX.SubItems(5) = rs!Tipo
      itmX.SubItems(6) = rs!BancoDesc
      
      itmX.Checked = chkTodos.Value
      
      If Trim(rs!nDocumento & "") = "" Then
         itmX.Bold = True
         itmX.ForeColor = vbRed
      
         curMontoErr = curMontoErr + rs!Monto
         lngErr = lngErr + 1
      Else
         curMonto = curMonto + rs!Monto
         lng = lng + 1
      End If
      
     
  rs.MoveNext
Loop
rs.Close
  
lblAsientosGenerar.Caption = lng
lblMontoGenerar.Caption = Format(curMonto, "Standard")

lblAsientosError.Caption = lngErr
lblMontoError.Caption = Format(curMontoErr, "Standard")


Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbConsulta_Emitidos()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Despliega en pantalla las solicitudes cuyo estado de asiento esta pendiente
'               y a la vez tienen el estado de la solicitud impreso o transferido.
'REFERENCIAS:   ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, lng As Long, curMonto As Currency
Dim lngErr As Long, curMontoErr As Currency


On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "Select T.Nsolicitud,T.Ndocumento,T.Monto,T.Fecha_Emision,T.Beneficiario,T.Tipo,B.descripcion as 'BancoDesc'" _
       & " FROM Tes_Transacciones T inner join Tes_Bancos B on T.id_Banco = B.id_Banco" _
       & " Where T.Estado_Asiento = 'P' And T.Estado in('T','I') and T.fecha_emision between '" _
       & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

If chkBancosTodos.Value = vbUnchecked Then
   strSQL = strSQL & " and T.id_Banco = " & cbo.ItemData(cbo.ListIndex)
End If

If chkDocumentosTodos.Value = vbUnchecked Then
   strSQL = strSQL & " and T.Tipo = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
End If




strSQL = strSQL & " Order by T.Id_Banco,T.Tipo,T.Fecha_Emision,T.NDocumento"

Call OpenRecordSet(rs, strSQL)

curMonto = 0
lng = 0

lngErr = 0
curMontoErr = 0
  
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!NSolicitud)
      itmX.SubItems(1) = rs!nDocumento & ""
      itmX.SubItems(2) = Format(rs!Monto, "Standard")
      itmX.SubItems(3) = rs!Fecha_Emision
      itmX.SubItems(4) = rs!Beneficiario
      itmX.SubItems(5) = rs!Tipo
      itmX.SubItems(6) = rs!BancoDesc
      
      itmX.Checked = chkTodos.Value
      
      If Trim(rs!nDocumento & "") = "" Then
         itmX.Bold = True
         itmX.ForeColor = vbRed
      
         curMontoErr = curMontoErr + rs!Monto
         lngErr = lngErr + 1
      Else
         curMonto = curMonto + rs!Monto
         lng = lng + 1
      End If
      
     
  rs.MoveNext
Loop
rs.Close
  
lblAsientosGenerar.Caption = lng
lblMontoGenerar.Caption = Format(curMonto, "Standard")

lblAsientosError.Caption = lngErr
lblMontoError.Caption = Format(curMontoErr, "Standard")


Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub Timer1_Timer()

Timer1.Interval = 0

Me.MousePointer = vbHourglass
    
Call btnEstado_Click(0)

Me.MousePointer = vbDefault

End Sub


