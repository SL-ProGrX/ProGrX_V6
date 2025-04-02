VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmActivos_Reportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   Icon            =   "frmActivos_Reportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.CheckBox chkTipoActivo 
      Height          =   312
      Left            =   8520
      TabIndex        =   23
      Top             =   1440
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   312
      Left            =   -2760
      TabIndex        =   22
      Top             =   -3000
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   556
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Left            =   0
      TabIndex        =   20
      Top             =   6480
      Width           =   10575
      _Version        =   1441793
      _ExtentX        =   18648
      _ExtentY        =   2561
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   732
         Left            =   7680
         TabIndex        =   21
         Top             =   240
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Appearance      =   14
         Picture         =   "frmActivos_Reportes.frx":030A
      End
      Begin XtremeSuiteControls.CheckBox chkInformeResumen 
         Height          =   312
         Left            =   4680
         TabIndex        =   31
         Top             =   240
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Informe Resumen    "
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
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   2775
      Left            =   0
      TabIndex        =   9
      Top             =   3600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Reporte"
         Object.Width           =   4834
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpHistorico 
      Height          =   315
      Left            =   5880
      TabIndex        =   10
      Top             =   4800
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   5760
      TabIndex        =   11
      Top             =   3960
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
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   315
      Left            =   7080
      TabIndex        =   12
      Top             =   3960
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
      Enabled         =   0   'False
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   315
      Left            =   5760
      TabIndex        =   13
      Top             =   3600
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
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
   Begin XtremeSuiteControls.ComboBox cboMod 
      Height          =   315
      Left            =   5880
      TabIndex        =   14
      Top             =   5160
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   2520
      TabIndex        =   15
      Top             =   1440
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
   Begin XtremeSuiteControls.ComboBox cboDep 
      Height          =   312
      Left            =   2520
      TabIndex        =   16
      Top             =   1800
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
   Begin XtremeSuiteControls.ComboBox cboSec 
      Height          =   312
      Left            =   2520
      TabIndex        =   17
      Top             =   2160
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
   Begin XtremeSuiteControls.FlatEdit txtResponsable 
      Height          =   315
      Left            =   5880
      TabIndex        =   19
      Top             =   5520
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4678
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkDepartamentos 
      Height          =   312
      Left            =   8520
      TabIndex        =   24
      Top             =   1800
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkSeccion 
      Height          =   312
      Left            =   8520
      TabIndex        =   25
      Top             =   2160
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkEstados 
      Height          =   315
      Left            =   8520
      TabIndex        =   26
      Top             =   3600
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   315
      Left            =   8520
      TabIndex        =   27
      Top             =   3960
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRes 
      Height          =   315
      Left            =   8640
      TabIndex        =   28
      Top             =   5520
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkMod 
      Height          =   315
      Left            =   8640
      TabIndex        =   29
      Top             =   5160
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
   End
   Begin XtremeSuiteControls.CheckBox chkActivoDetalle 
      Height          =   315
      Left            =   5880
      TabIndex        =   30
      Top             =   5880
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Incluir detalle de los activos"
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
      Appearance      =   16
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboLocaliza 
      Height          =   330
      Left            =   2520
      TabIndex        =   34
      Top             =   2520
      Width           =   5895
      _Version        =   1441793
      _ExtentX        =   10398
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.CheckBox chkUbicacion 
      Height          =   315
      Left            =   8520
      TabIndex        =   36
      Top             =   2520
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación"
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
      Index           =   4
      Left            =   960
      TabIndex        =   35
      Top             =   2520
      Width           =   1575
   End
   Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   33
      Top             =   3120
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11451
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Filtros:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   32
      Top             =   3120
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7006
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Informes:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Activos Fijos"
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
      Height          =   492
      Index           =   3
      Left            =   1800
      TabIndex        =   18
      Top             =   360
      Width           =   7212
   End
   Begin VB.Label lblx06 
      BackStyle       =   0  'Transparent
      Caption         =   "Responsable"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblx05 
      BackStyle       =   0  'Transparent
      Caption         =   "Modificaciones"
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
      Left            =   4200
      TabIndex        =   7
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label lblHistorico 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label lblx04 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label lblx01 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Activo"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblx02 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Activo"
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
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   1572
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmActivos_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mReportKey As String
Dim vTitulo As String, vSubTitulo As String

Private Sub cboDep_Click()
If Not vPaso Then Exit Sub


Dim strSQL As String

strSQL = "select rtrim(cod_seccion) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from Activos_secciones where cod_departamento = '" & cboDep.ItemData(cboDep.ListIndex) _
       & "' order by cod_seccion"
Call sbCbo_Llena_New(cboSec, strSQL, False, True)

End Sub

Private Sub chkDepartamentos_Click()
If chkDepartamentos.Value = vbChecked Then
 cboDep.Enabled = False
Else
 cboDep.Enabled = True
End If
End Sub

Private Sub chkEstados_Click()
If chkEstados.Value = vbChecked Then
 cboEstado.Enabled = False
Else
 cboEstado.Enabled = True
End If
End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
 dtpInicio.Enabled = False
 dtpCorte.Enabled = False
Else
 dtpInicio.Enabled = True
 dtpCorte.Enabled = True
End If

End Sub

Private Sub chkMod_Click()
If chkMod.Value = vbChecked Then
 cboMod.Enabled = False
Else
 cboMod.Enabled = True
End If

End Sub

Private Sub chkRes_Click()
If chkRes.Value = vbChecked Then
 txtResponsable.Enabled = False
Else
 txtResponsable.Enabled = True
End If
End Sub

Private Sub chkSeccion_Click()
If chkSeccion.Value = vbChecked Then
 cboSec.Enabled = False
Else
 cboSec.Enabled = True
End If
End Sub

Private Sub chkTipoActivo_Click()
If chkTipoActivo.Value = vbChecked Then
 cboTipo.Enabled = False
Else
 cboTipo.Enabled = True
End If
End Sub

Private Function fxSQL(pTipo As String) As String
Dim vCadena As String

vCadena = ""
vSubTitulo = ""
vTitulo = ""




Select Case pTipo
  Case "LA" 'Lista de Activos (Información Actual)
    If chkTipoActivo.Value = vbUnchecked Then
       vCadena = "{Activos_Principal.TIPO_ACTIVO} = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
       vSubTitulo = vSubTitulo & " ¦ TIPO: " & cboTipo.ItemData(cboTipo.ListIndex)
    End If
    
    If chkDepartamentos.Value = vbUnchecked Then
       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
       vCadena = vCadena & "{Activos_Principal.COD_DEPARTAMENTO} = '" & cboDep.ItemData(cboDep.ListIndex) & "'"
       vSubTitulo = vSubTitulo & " ¦ DEPT: " & cboDep.ItemData(cboDep.ListIndex)
    End If
    
    If chkSeccion.Value = vbUnchecked Then
       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
       vCadena = vCadena & "{Activos_Principal.COD_SECCION} = '" & cboSec.ItemData(cboSec.ListIndex) & "'"
       vSubTitulo = vSubTitulo & " ¦ SEC: " & cboSec.ItemData(cboSec.ListIndex)
    End If
    
    If chkUbicacion.Value = vbUnchecked Then
       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
       vCadena = vCadena & "{Activos_Principal.COD_LOCALIZA} = '" & cboLocaliza.ItemData(cboLocaliza.ListIndex) & "'"
       vSubTitulo = vSubTitulo & " ¦ UBICA: " & cboLocaliza.ItemData(cboLocaliza.ListIndex)
    End If
    
    
    If chkEstados.Value = vbUnchecked Then
      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
      Select Case Mid(cboEstado.Text, 1, 2)
        Case "01" 'Vigentes
           vCadena = vCadena & "{Activos_Principal.ESTADO} = 'A' AND ({Activos_Principal.VALOR_LIBROS_PERIODO} > {Activos_Principal.VALOR_DESECHO} OR {Activos_Principal.depreciacion_ac} > 0 "
        Case "02" 'Depreciados
           vCadena = vCadena & "{Activos_Principal.ESTADO} = 'A' AND {Activos_Principal.VALOR_LIBROS_PERIODO} <= {Activos_Principal.VALOR_DESECHO}"
        Case "03" 'Retirados
           vCadena = vCadena & "{Activos_Principal.ESTADO} = 'R'"
      End Select
      vSubTitulo = vSubTitulo & " ¦ ESTADO: " & cboEstado.Text
    End If
     
    If chkRes.Value = vbUnchecked Then
      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
      
      vCadena = vCadena & "{Activos_Principal.IDENTIFICACION} = '" & txtResponsable.Tag & "'"
      vSubTitulo = vSubTitulo & " ¦ RESPONSABLE: " & txtResponsable.Tag
    End If
     
    If chkFechas.Value = vbUnchecked Then
      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
        vSubTitulo = vSubTitulo & " ¦ ADQUISICION: " & Format(dtpInicio.Value, "dd/mm/yyyy") _
                   & " - " & Format(dtpCorte.Value, "dd/mm/yyyy")
       
        vCadena = vCadena & "{Activos_Principal.FECHA_ADQUISICION} in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    End If
    

  '-------------------------------------------------------------------------------------------------------------------
  
  Case "LH" 'Auxiliar: Lista de Activos
       Call dtpHistorico_Change
       
       vCadena = "{vActivos_AuxiliarConsolidado.ANIO} = " & Year(dtpHistorico.Value) _
                & " AND {vActivos_AuxiliarConsolidado.MES} = " & Month(dtpHistorico.Value)
       vSubTitulo = " PERIODO: " & lblHistorico.Caption & " ¦ " & fxActivos_PeriodoEstado(dtpHistorico.Value)
    
    If chkTipoActivo.Value = vbUnchecked Then
       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
       vCadena = vCadena & "{vActivos_AuxiliarConsolidado.TIPO_ACTIVO} = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
       vSubTitulo = vSubTitulo & " ¦ TIPO: " & cboTipo.ItemData(cboTipo.ListIndex)
    End If
    
    If chkDepartamentos.Value = vbUnchecked Then
       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
       vCadena = vCadena & "{vActivos_AuxiliarConsolidado.COD_DEPARTAMENTO} = '" & cboDep.ItemData(cboDep.ListIndex) & "'"
       vSubTitulo = vSubTitulo & " ¦ DEPT: " & cboDep.ItemData(cboDep.ListIndex)
    End If
    
    If chkSeccion.Value = vbUnchecked Then
       If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
       vCadena = vCadena & "{vActivos_AuxiliarConsolidado.COD_SECCION} = '" & cboSec.ItemData(cboSec.ListIndex) & "'"
       vSubTitulo = vSubTitulo & " ¦ SEC: " & cboSec.ItemData(cboSec.ListIndex)
    End If
    
    If chkEstados.Value = vbUnchecked Then
              
      Select Case Mid(cboEstado.Text, 1, 2)
        Case "01" 'Vigentes
  '         vCadena = vCadena & "{vActivos_AuxiliarConsolidado.VALOR_LIBROS_CONSOLIDADO} > {vActivos_AuxiliarConsolidado.VALOR_DESECHO}"
        Case "02" 'Depreciados
           If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
           vCadena = vCadena & "{vActivos_AuxiliarConsolidado.VALOR_LIBROS_CONSOLIDADO} <= {vActivos_AuxiliarConsolidado.VALOR_DESECHO}"
      
'        Case "01" 'Vigentes
'           vCadena = vCadena & "{Activos_Principal.VALOR_LIBROS_PERIODO} > {Activos_Principal.VALOR_DESECHO} OR {Activos_Principal.depreciacion_ac} > 0 "
'        Case "02" 'Depreciados
'           vCadena = vCadena & "{Activos_Principal.ESTADO} = 'A' AND {Activos_Principal.VALOR_LIBROS_PERIODO} <= {Activos_Principal.VALOR_DESECHO}"
'        Case "03" 'Retirados
'           vCadena = vCadena & "{Activos_Principal.ESTADO} = 'R'"
      
      End Select
      vSubTitulo = vSubTitulo & " ¦ ESTADO: " & cboEstado.Text
    End If
     
    If chkRes.Value = vbUnchecked Then
      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
      
      vCadena = vCadena & "{vActivos_AuxiliarConsolidado.IDENTIFICACION} = '" & txtResponsable.Tag & "'"
      vSubTitulo = vSubTitulo & " ¦ RESPONSABLE: " & txtResponsable.Tag
    End If
     
    If chkFechas.Value = vbUnchecked Then
      If Len(vCadena) > 0 Then vCadena = vCadena & " AND "
        vSubTitulo = vSubTitulo & " ¦ ADQUISICION: " & Format(dtpInicio.Value, "dd/mm/yyyy") _
                   & " - " & Format(dtpCorte.Value, "dd/mm/yyyy")
       
        vCadena = vCadena & "{vActivos_AuxiliarConsolidado.FECHA_ADQUISICION} in date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                & ") to date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    End If


End Select

fxSQL = vCadena

End Function

Private Sub chkUbicacion_Click()
If chkUbicacion.Value = vbChecked Then
 cboLocaliza.Enabled = False
Else
 cboLocaliza.Enabled = True
End If
End Sub

Private Sub cmdReporte_Click()
Dim vSQL As String


With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Activos Fijos"
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"

  vSQL = fxSQL(Mid(mReportKey, 1, 2))

Select Case mReportKey
  Case "LA001"   'Lista General
    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
        
    If chkInformeResumen.Value = vbUnchecked Then
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneral.rpt")
    Else
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralRsm.rpt")
    End If
   
  Case "LA002"   'Lista General x Departamento
    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
        
    If chkInformeResumen.Value = vbUnchecked Then
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralDeptTipo.rpt")
    Else
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralDeptTipoRsm.rpt")
    End If
    
  Case "LA003"   'Lista General x Persona
    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
        
    If chkInformeResumen.Value = vbUnchecked Then
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralPersona.rpt")
    Else
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralPersonaRsm.rpt")
    End If
    
  Case "LA005"   'Lista General x Ubicacion
    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
        
    If chkInformeResumen.Value = vbUnchecked Then
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralUbicacion.rpt")
    Else
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_ListadoGeneralUbicacionRsm.rpt")
    End If
    
  Case "LA004"   'Informe Contable
    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
        
    .ReportFileName = SIFGlobal.fxPathReportes("Activos_InformeContable.rpt")
    
  Case "LH001"   'Auxiliar: Lista General
    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
        
    If chkInformeResumen.Value = vbUnchecked Then
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneral.rpt")
    Else
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneralRsm.rpt")
    End If
   
  Case "LH002"   'Auxiliar: Lista x Departamento
    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
        
    If chkInformeResumen.Value = vbUnchecked Then
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneralDept.rpt")
    Else
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneralDeptRsm.rpt")
    End If
    
  Case "LH003"   'Auxiliar: Lista x Persona
    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
        
    If chkInformeResumen.Value = vbUnchecked Then
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneralPersona.rpt")
    Else
       .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxListadoGeneralPersonaRsm.rpt")
    End If
    
  Case "LH004"   'Auxiliar: Informe Contable
    .Formulas(3) = "fxSubTitulo = '" & UCase(vSubTitulo) & "'"
        
    .ReportFileName = SIFGlobal.fxPathReportes("Activos_AuxInformeContable.rpt")
  
  
  Case 4 'Depreciacion Historica
  Case 5 'Boletas
  Case 6 'Lista de Modificaciones
  Case 7 'Asignacion de Polizas
  Case 8 'Lista de Responsables
  Case 9 'Lista de Departamentos y Secciones
  Case 10 'Lista de Proveedores
  Case 11 'Lista de Tipos Activos
End Select
    
 .SelectionFormula = vSQL
 .PrintReport
End With

End Sub

Private Sub dtpHistorico_Change()

Select Case Month(dtpHistorico.Value)
 Case 1
    lblHistorico.Caption = "ENERO DE " & Year(dtpHistorico.Value)
 Case 2
    lblHistorico.Caption = "FEBRERO DE " & Year(dtpHistorico.Value)
 Case 3
    lblHistorico.Caption = "MARZO DE " & Year(dtpHistorico.Value)
 Case 4
    lblHistorico.Caption = "ABRIL DE " & Year(dtpHistorico.Value)
 Case 5
    lblHistorico.Caption = "MAYO DE " & Year(dtpHistorico.Value)
 Case 6
    lblHistorico.Caption = "JUNIO DE " & Year(dtpHistorico.Value)
 Case 7
    lblHistorico.Caption = "JULIO DE " & Year(dtpHistorico.Value)
 Case 8
    lblHistorico.Caption = "AGOSTO DE " & Year(dtpHistorico.Value)
 Case 9
    lblHistorico.Caption = "SETIEMBRE DE " & Year(dtpHistorico.Value)
 Case 10
    lblHistorico.Caption = "OCTUBRE DE " & Year(dtpHistorico.Value)
 Case 11
    lblHistorico.Caption = "NOVIEMBRE DE " & Year(dtpHistorico.Value)
 Case 12
    lblHistorico.Caption = "DICIEMBRE DE " & Year(dtpHistorico.Value)
End Select

End Sub

Private Sub sbListaReportes()

With lsw.ListItems
   .Clear
   .Add , "LA001", "Lista General"
   .Add , "LA002", "Lista General x Departamento"
   .Add , "LA003", "Lista General x Persona"
   .Add , "LA005", "Lista General x Ubicación"
   .Add , "LA004", "Informe Contable"
   
   
   .Add , "LH001", "Auxiliar: Lista General"
   .Add , "LH002", "Auxiliar: Lista x Departamento"
   .Add , "LH003", "Auxiliar: Lista x Persona"
   .Add , "LH004", "Auxiliar: Informe Contable"
   
   
'lsw.ListItems.Clear
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Lista de Activos"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Justificaciones"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Depreciación Actual"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Depreciación Historica"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Boletas de Asignación"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Modificaciones"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Pólizas"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Listas de Responsables"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Departamentos / Secciones"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Lista de Proveedores"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Tipos de Activos"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Conciliación"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Asientos"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Traslados y Cambios Resp."
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Traslado Dep/Sec (Ubicacion)"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Traslado Responsabilidad"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Tipos de Obras"
'lsw.ListItems.Add lsw.ListItems.Count + 1, , "Obras en Proceso"

End With
lsw.ListItems.Item(1).ForeColor = vbBlue
lsw.ListItems.Item(1).Bold = vbBlue

lsw.ListItems.Item(1).Selected = True

Call sbOpciones(1)



End Sub


Private Sub dtpInicio_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()

vModulo = 36

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub lsw_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).ForeColor = vbBlack
  lsw.ListItems.Item(i).Bold = False
Next

lsw.SelectedItem.Bold = True
lsw.SelectedItem.ForeColor = vbBlue

mReportKey = lsw.SelectedItem.Key

Call sbOpciones(Mid(lsw.SelectedItem.Key, 1, 2))

End Sub

Private Sub sbOpciones(pTipo As String)

lblx01.ForeColor = vbBlack
lblx02.ForeColor = vbBlack
'lblx03.ForeColor = vbBlack
lblx04.ForeColor = vbBlack
lblx05.ForeColor = vbBlack
lblx06.ForeColor = vbBlack
chkActivoDetalle.ForeColor = vbBlack

Select Case pTipo
  Case "LA"   'Lista de Activos
    lblx01.ForeColor = vbBlue
    lblx02.ForeColor = vbBlue
'    lblx03.ForeColor = vbBlue
    chkActivoDetalle.ForeColor = vbBlue
  Case "JU" 'Justificaciones
  Case "LH" 'Lista Historica
    lblx04.ForeColor = vbBlue
  Case "BT" 'Boletas de Traslado
    lblx06.ForeColor = vbBlue
  Case "LM" 'Lista de Modificaciones
    lblx01.ForeColor = vbBlue
    lblx05.ForeColor = vbBlue
  Case "AP" 'Asignacion de Polizas
  Case "LR" 'Lista de Responsables
  Case "DS" 'Lista de Departamentos y Secciones
  Case "LP" 'Lista de Proveedores
  Case "TA" 'Lista de Tipos Activos
End Select

End Sub


Private Sub sbInicializa()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


vPaso = False
    

cboMod.Clear
cboMod.AddItem "Adiciones y Mejoras"
cboMod.AddItem "Retiros (Salidas)"
cboMod.AddItem "Revaluaciones"
cboMod.AddItem "Deterioros y Devalorizaciones"
cboMod.Text = "Adiciones y Mejoras"


cboEstado.Clear
cboEstado.AddItem "01 - Vigentes"
cboEstado.AddItem "02 - Depreciados"
cboEstado.AddItem "03 - Retirados"
cboEstado.Text = "01 - Vigentes"

strSQL = "select rtrim(tipo_activo) as 'IdX' , rtrim(descripcion) as 'ItmX'" _
       & " from Activos_tipo_activo order by tipo_activo"
Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

strSQL = "select rtrim(cod_departamento) as 'IdX',  rtrim(descripcion) as 'ItmX'" _
       & " from Activos_departamentos order by cod_departamento"
Call sbCbo_Llena_New(cboDep, strSQL, False, True)

strSQL = "select rtrim(COD_LOCALIZA) as 'Idx', rtrim(descripcion) as 'ItmX'" _
     & " from ACTIVOS_LOCALIZACIONES Where Activa = 1 order by descripcion"
Call sbCbo_Llena_New(cboLocaliza, strSQL, False, True)


vPaso = True

Call sbListaReportes

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
dtpHistorico.Value = gActivos.Periodo

Call dtpHistorico_Change


Call cboDep_Click


Call chkTipoActivo_Click
Call chkDepartamentos_Click
Call chkSeccion_Click
Call chkUbicacion_Click

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub

Private Sub txtResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select IDENTIFICACION, NOMBRE , Departamento, Seccion " _
                    & " From vActivos_Personas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
     txtResponsable.Text = gBusquedas.Resultado2
     txtResponsable.Tag = gBusquedas.Resultado
  End If

End If
End Sub
