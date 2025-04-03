VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmActivos_ObrasProceso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Obras en Proceso"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   11055
      _Version        =   1572864
      _ExtentX        =   19500
      _ExtentY        =   9763
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
      ItemCount       =   5
      SelectedItem    =   3
      Item(0).Caption =   "Datos"
      Item(0).ControlCount=   29
      Item(0).Control(0)=   "txtRegistro"
      Item(0).Control(1)=   "txtDisponible"
      Item(0).Control(2)=   "txtDesembolso"
      Item(0).Control(3)=   "txtFiniquito"
      Item(0).Control(4)=   "txtNotas"
      Item(0).Control(5)=   "txtPresuActual"
      Item(0).Control(6)=   "txtAdendums"
      Item(0).Control(7)=   "txtPresupuesto"
      Item(0).Control(8)=   "txtUbicacion"
      Item(0).Control(9)=   "txtEncargado"
      Item(0).Control(10)=   "cbo"
      Item(0).Control(11)=   "dtpInicio"
      Item(0).Control(12)=   "dtpFinaliza"
      Item(0).Control(13)=   "Label2(13)"
      Item(0).Control(14)=   "Label2(12)"
      Item(0).Control(15)=   "Label2(11)"
      Item(0).Control(16)=   "Label2(10)"
      Item(0).Control(17)=   "Label2(9)"
      Item(0).Control(18)=   "Label2(7)"
      Item(0).Control(19)=   "Label2(6)"
      Item(0).Control(20)=   "Label2(5)"
      Item(0).Control(21)=   "Label2(4)"
      Item(0).Control(22)=   "Label2(3)"
      Item(0).Control(23)=   "Label2(2)"
      Item(0).Control(24)=   "lblEstado"
      Item(0).Control(25)=   "Label2(1)"
      Item(0).Control(26)=   "Label2(0)"
      Item(0).Control(27)=   "txtProveedor"
      Item(0).Control(28)=   "Label4"
      Item(1).Caption =   "Adendums"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGridA"
      Item(2).Caption =   "Desembolsos"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGridD"
      Item(3).Caption =   "Finiquito"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "cmdAplicar"
      Item(3).Control(1)=   "cboFiniquito"
      Item(3).Control(2)=   "dtpFiniquito"
      Item(3).Control(3)=   "Label3(0)"
      Item(4).Caption =   "Resultados"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "lsw"
      Item(4).Control(1)=   "cmdDetallar"
      Item(4).Control(2)=   "scTitulo"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3975
         Left            =   -70000
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1572864
         _ExtentX        =   19288
         _ExtentY        =   7011
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
         Appearance      =   21
      End
      Begin VB.ComboBox cboFiniquito 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmActivos_ObrasProceso.frx":0000
         Left            =   360
         List            =   "frmActivos_ObrasProceso.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   3972
      End
      Begin XtremeSuiteControls.PushButton cmdDetallar 
         Height          =   375
         Left            =   -67240
         TabIndex        =   5
         Top             =   4920
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1572864
         _ExtentX        =   9546
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   ">>>  Detalla los activos o mejoras a realizar con esta obra  <<<"
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
         Appearance      =   21
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   1920
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4254
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aplicar Finiquito"
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
         Appearance      =   21
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFiniquito 
         Height          =   315
         Left            =   4320
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
         _Version        =   1572864
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
      Begin FPSpreadADO.fpSpread vGridD 
         Height          =   4935
         Left            =   -69880
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   8705
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
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frmActivos_ObrasProceso.frx":0067
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridA 
         Height          =   4935
         Left            =   -69880
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
         _ExtentY        =   8705
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmActivos_ObrasProceso.frx":0778
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   -67840
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
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
      Begin XtremeSuiteControls.DateTimePicker dtpFinaliza 
         Height          =   315
         Left            =   -64960
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1995
         Left            =   -64960
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   4695
         _Version        =   1572864
         _ExtentX        =   8281
         _ExtentY        =   3519
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
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   330
         Left            =   -67840
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   5055
         _Version        =   1572864
         _ExtentX        =   8916
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
      Begin XtremeSuiteControls.FlatEdit txtProveedor 
         Height          =   330
         Left            =   -67840
         TabIndex        =   34
         Top             =   960
         Visible         =   0   'False
         Width           =   7575
         _Version        =   1572864
         _ExtentX        =   13361
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtEncargado 
         Height          =   330
         Left            =   -67840
         TabIndex        =   35
         Top             =   1320
         Visible         =   0   'False
         Width           =   7575
         _Version        =   1572864
         _ExtentX        =   13361
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtUbicacion 
         Height          =   330
         Left            =   -67840
         TabIndex        =   36
         Top             =   2040
         Visible         =   0   'False
         Width           =   7575
         _Version        =   1572864
         _ExtentX        =   13361
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtFiniquito 
         Height          =   330
         Left            =   -62680
         TabIndex        =   38
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPresupuesto 
         Height          =   330
         Left            =   -67840
         TabIndex        =   39
         Top             =   2760
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAdendums 
         Height          =   330
         Left            =   -67840
         TabIndex        =   40
         Top             =   3120
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPresuActual 
         Height          =   330
         Left            =   -67840
         TabIndex        =   41
         Top             =   3480
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDesembolso 
         Height          =   330
         Left            =   -67840
         TabIndex        =   42
         Top             =   3840
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDisponible 
         Height          =   330
         Left            =   -67840
         TabIndex        =   43
         Top             =   4200
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRegistro 
         Height          =   330
         Left            =   -67840
         TabIndex        =   44
         Top             =   4560
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Left            =   -64960
         TabIndex        =   32
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Notas:"
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
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   -69760
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Encargado"
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
         Left            =   -69760
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Estado"
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
         Height          =   315
         Left            =   -62680
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Index           =   2
         Left            =   -69760
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Finalización"
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
         Left            =   -66280
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   4
         Left            =   -69760
         TabIndex        =   23
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Presupuesto"
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
         Left            =   -69760
         TabIndex        =   22
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "(+) Adendums"
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
         Left            =   -69760
         TabIndex        =   21
         Top             =   3120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Presu. Actual"
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
         Index           =   7
         Left            =   -69760
         TabIndex        =   20
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Finiquito"
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
         Index           =   9
         Left            =   -63520
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Desembolsos"
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
         Index           =   10
         Left            =   -69760
         TabIndex        =   18
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible"
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
         Index           =   11
         Left            =   -69760
         TabIndex        =   17
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente de Registro"
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
         Index           =   12
         Left            =   -69760
         TabIndex        =   16
         Top             =   4560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
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
         Index           =   13
         Left            =   -69760
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Datos del Finiquito:"
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
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Left            =   -70000
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1572864
         _ExtentX        =   19288
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Resultados y Desgloce de la Obra en Nuevos Activos o Mejoras y Adiciones a existentes"
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
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6825
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Usuario que Registra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   3775
            MinWidth        =   3775
            Object.ToolTipText     =   "Distribuido"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   330
      Left            =   3120
      TabIndex        =   33
      Top             =   720
      Width           =   6855
      _Version        =   1572864
      _ExtentX        =   12091
      _ExtentY        =   582
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
      Height          =   330
      Left            =   1440
      TabIndex        =   37
      Top             =   720
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2990
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Descripción"
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
      Height          =   252
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   480
      Width           =   6852
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Contrato"
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
      Height          =   252
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1692
   End
End
Attribute VB_Name = "frmActivos_ObrasProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vEdita As Boolean, vCodigo As String
Dim vPaso As Boolean

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedor.SetFocus
End Sub

Private Sub cmdAplicar_Click()
Dim strSQL As String

On Error GoTo vError

If lblEstado.Tag <> "P" Then
  MsgBox "Este Contrato ya se finiquito anteriormente...", vbExclamation
  Exit Sub
End If

strSQL = "update Activos_obras set estado = '" & Mid(cboFiniquito.Text, Len(cboFiniquito.Text) - 1, 1) _
       & "',fecha_finiquito = '" & Format(dtpFiniquito.Value, "yyyy/mm/dd") _
       & "' where contrato = '" & vCodigo & "'"
Call ConectionExecute(strSQL)

Call sbConsulta(vCodigo)

MsgBox "Finiquito Registrado Satisfactoriamente...", vbInformation

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdDetallar_Click()

If lblEstado.Tag = "I" Or lblEstado.Tag = "S" Then
    gAsistente.Documento = txtCodigo
    gAsistente.Proveedor = txtProveedor.Tag
    gAsistente.VU = CCur(txtDisponible)
    gAsistente.Tipo = "O"
    Call sbFormsCall("frmActivos_Main", , , , , Me, True)
Else
 MsgBox "Se debe de registrar el finiquito satisfactorio ...", vbInformation
End If

End Sub

Private Sub dtpFinaliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUbicacion.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFinaliza.SetFocus
End Sub

Public Sub gWizardX()
Call sbConsulta(txtCodigo)
End Sub

Private Sub Form_Activate()
vModulo = 36
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
vModulo = 36

With lsw.ColumnHeaders
        .Clear
        .Add , , "Id", 1800
        .Add , , "Tipo", 2100, vbCenter
        .Add , , "Placa Id", 2100, vbCenter
        .Add , , "Monto", 2500, vbRightJustify
        .Add , , "Id Add", 1500, vbCenter
        .Add , , "Nombre", 3500
        .Add , , "Tipo Activo", 3500
End With

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

  
End Sub

Private Sub sbLimpiaPantalla()


tcMain.Item(0).Selected = True
txtDescripcion.Text = ""

vCodigo = ""
txtCodigo = ""

lblEstado.Caption = ""


strSQL = "select rtrim(cod_tipo) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from Activos_obras_tipos"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

dtpInicio.Value = fxFechaServidor
dtpFinaliza.Value = dtpInicio.Value

txtNotas = ""
txtProveedor = ""
txtProveedor.Tag = ""

txtEncargado = ""
txtFiniquito = ""
txtUbicacion = ""

txtPresuActual = 0
txtPresupuesto = 0
txtDesembolso = 0
txtAdendums = 0
txtDisponible = 0
txtRegistro = 0
StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""
StatusBarX.Panels(3).Text = 0


tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False
tcMain.Item(3).Enabled = False
tcMain.Item(4).Enabled = False


End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim rs As New ADODB.Recordset, i As Integer

If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.Col = i
     vGrid.Text = ""
    Next i
End If

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
      Case 1 'Secuencia
        vGrid.Text = CStr(rs!secuencia)
      Case 2 'Tipo de Desembolso
        vGrid.Text = CStr(rs!cod_desembolso)
        vGrid.TextTip = TextTipFloating
        vGrid.CellNote = rs!desembolso
      Case 3 'Proveedor
        vGrid.Text = CStr(rs!COD_PROVEEDOR)
        vGrid.TextTip = TextTipFloating
        vGrid.CellNote = rs!Proveedor
      Case 4 'Documento
        vGrid.Text = rs!Documento
      Case 5 'Fecha
        vGrid.Text = rs!fecha
      Case 6 'Monto
        vGrid.Text = CStr(rs!monto)
    End Select
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End Sub




Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Item.Index

  Case 1 'Adendums
     strSQL = "select cod_Adendum,descripcion,fecha,monto from Activos_obras_ade" _
           & " where contrato = '" & vCodigo & "'"
     Call sbCargaGrid(vGridA, 4, strSQL)
  Case 2 'Desembolsos
     strSQL = "select D.*,T.descripcion as Desembolso,P.descripcion as Proveedor" _
            & " from Activos_Obras_Desem D inner join Activos_obras_tDesem T on D.cod_desembolso = T.cod_desembolso" _
            & " inner join Activos_Proveedores P on D.cod_proveedor = P.cod_Proveedor" _
            & " where D.contrato = '" & vCodigo & "'"
     Call sbCargaGridLocal(vGridD, 6, strSQL)
            
  Case 3 'Finiquito
    cboFiniquito.Text = "Finalizacion Satisfactoria (S)"
    dtpFiniquito.Value = fxFechaServidor
    
  Case 4 'Resultados
    strSQL = "select O.ID_RESULTADOS,'ACTIVO' as Tipo,O.num_placa,A.valor_historico as Monto" _
           & ",O.id_adicion,A.nombre,T.descripcion as TA" _
           & " from Activos_obras_resultados O inner join Activos_Principal A on O.num_placa = A.num_placa" _
           & " inner join Activos_tipo_activo T on A.tipo_activo = T.tipo_activo" _
           & " where O.tipo = 'A' and O.contrato = '" & vCodigo & "'"
    strSQL = strSQL & " UNION "
    strSQL = strSQL & "select O.ID_RESULTADOS,'MEJORAS' as Tipo,O.num_placa,A.Monto" _
           & ",O.id_adicion,A.descripcion as nombre,T.descripcion as TA" _
           & " from Activos_obras_resultados O inner join Activos_retiro_adicion A on O.num_placa = A.num_placa" _
           & " and O.id_adicion = A.ID_ADDRET" _
           & " inner join Activos_justificaciones T on A.cod_justificacion = T.cod_justificacion" _
           & " where O.tipo = 'M' and O.contrato = '" & vCodigo & "'"
    Call OpenRecordSet(rs, strSQL, 0)
    
    lsw.ListItems.Clear
    
    Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!ID_RESULTADOS)
         itmX.SubItems(1) = rs!Tipo
         itmX.SubItems(2) = rs!num_placa
         itmX.SubItems(3) = Format(rs!monto, "Standard")
         itmX.SubItems(4) = rs!id_adicion & ""
         itmX.SubItems(5) = rs!Nombre
         itmX.SubItems(6) = rs!ta
     rs.MoveNext
    Loop
    rs.Close
  
End Select

vError:
Me.MousePointer = vbDefault

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
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
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select contrato,descripcion from Activos_obras"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select O.*, T.descripcion as 'TipoObra'" _
       & ",P.descripcion as Proveedor" _
       & " from Activos_obras O inner join Activos_obras_Tipos T on O.cod_tipo = T.cod_tipo" _
       & " inner join cxp_proveedores P on O.cod_proveedor = P.cod_proveedor" _
       & " where O.contrato = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vPaso = False
    
  vCodigo = rs!contrato
  txtCodigo = rs!contrato
 
  txtDescripcion = rs!Descripcion
  Call sbCboAsignaDato(cbo, rs!TipoObra, True, rs!cod_tipo)
  
  lblEstado.Tag = rs!Estado
  
  Select Case rs!Estado
   Case "P"
     lblEstado.Caption = "En Proceso"
   Case "I"
     lblEstado.Caption = "Finiquito Incompleto"
   Case "S"
     lblEstado.Caption = "Finiquito Satisfactorio"
   Case "D"
     lblEstado.Caption = "Proyecto Descartado"
  End Select
  
  txtNotas = rs!Notas
  
  txtProveedor = rs!Proveedor
  txtProveedor.Tag = rs!COD_PROVEEDOR
    
  txtFiniquito = rs!fecha_finiquito & ""
  txtEncargado = rs!encargado
  dtpInicio.Value = rs!fecha_Inicio
  dtpFinaliza.Value = rs!fecha_estimada
  txtUbicacion = rs!ubicacion
  
  txtPresupuesto.Text = Format(rs!presu_original, "Standard")
  txtAdendums.Text = Format(rs!addendums, "Standard")
  txtPresuActual.Text = Format(rs!presu_actual, "Standard")
  txtDesembolso.Text = Format(rs!desembolsado, "Standard")
  txtDisponible.Text = Format(rs!presu_actual - rs!desembolsado, "Standard")
  txtRegistro.Text = Format(rs!presu_actual - rs!distribuido, "Standard")
  
  tcMain.Item(0).Selected = True
  
  StatusBarX.Panels(1).Text = rs!Registro_Usuario & ""
  StatusBarX.Panels(2).Text = rs!Registro_fecha & ""
  StatusBarX.Panels(3).Text = Format(rs!distribuido, "Standard")
  
    tcMain.Item(1).Enabled = True
    tcMain.Item(2).Enabled = True
    tcMain.Item(3).Enabled = True
    tcMain.Item(4).Enabled = True
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String, i As Integer, x As Boolean

x = False
vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion de la Obra no es válido ..."
If dtpInicio.Value > dtpFinaliza.Value Then vMensaje = vMensaje & vbCrLf & " - La fecha de Adquisición no puede ser menor a la de instalacion ..."

If Not IsNumeric(txtPresupuesto) Then vMensaje = vMensaje & vbCrLf & " - El Presupuesto no es válido ..."
If txtProveedor.Tag = "" Then vMensaje = vMensaje & vbCrLf & " - Proveedor no es válido ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim i As Integer

On Error GoTo vError

If vEdita Then

 If CLng(txtDesembolso) = 0 Then
    strSQL = "update Activos_obras set descripcion = '" & UCase(txtDescripcion) _
           & "',encargado ='" & txtEncargado & "', notas = '" & txtNotas _
           & "',cod_proveedor = '" & txtProveedor.Tag & "',presu_original = " _
           & CCur(txtPresupuesto) & ",presu_actual = " & CCur(txtPresupuesto) _
           & ",ubicacion = '" & txtUbicacion _
           & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & "',fecha_estimada = '" & Format(dtpFinaliza.Value, "yyyy/mm/dd") _
           & "',cod_tipo = '" & cbo.ItemData(cbo.ListIndex) & "' where contrato = '" & vCodigo & "'"
    Call ConectionExecute(strSQL)
 End If
  
'  Call sbBitacora("Modifica", "Tipo Activo : " & vCodigo)

Else
    
   vCodigo = txtCodigo
   strSQL = "insert into Activos_obras(contrato,cod_tipo,descripcion,estado,encargado,cod_proveedor" _
          & ",fecha_inicio,fecha_estimada,notas,ubicacion,presu_original,addendums,presu_actual" _
          & ",desembolsado,distribuido, registro_usuario, registro_fecha) values('" & vCodigo & "','" & cbo.ItemData(cbo.ListIndex) _
          & "','" & UCase(txtDescripcion) & "','P','" & txtEncargado & "','" & txtProveedor.Tag _
          & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") & "','" & Format(dtpFinaliza.Value, "yyyy/mm/dd") _
          & "','" & txtNotas & "','" & txtUbicacion & "'," & CCur(txtPresupuesto) & ",0," & CCur(txtPresupuesto) & ",0,0,'" _
          & glogon.Usuario & "',getdate())"
   Call ConectionExecute(strSQL)
   ' Call sbBitacora("Registra", "Bodega: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(vCodigo)

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete Activos_Obras where contrato = '" & vCodigo & "'"
  If CCur(txtDesembolso) = 0 Then Call ConectionExecute(strSQL)
  
'  Call sbBitacora("Elimina", "Tipo Activo : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "contrato"
  gBusquedas.Orden = "contrato"
  gBusquedas.Consulta = "select contrato,descripcion from Activos_obras"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select contrato,descripcion from Activos_obras"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub


Private Sub txtEncargado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicio.SetFocus
End Sub

Private Sub txtPresupuesto_GotFocus()
On Error GoTo vError
txtPresupuesto = CCur(txtPresupuesto)
vError:
End Sub

Private Sub txtPresupuesto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtPresupuesto_LostFocus()
On Error GoTo vError
txtPresupuesto = Format(CCur(txtPresupuesto), "Standard")
vError:
End Sub

Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEncargado.SetFocus
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_proveedor,descripcion from Activos_proveedores"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtProveedor.Tag) Then
       txtProveedor.Tag = gBusquedas.Resultado
       txtProveedor = gBusquedas.Resultado2
       txtEncargado.SetFocus
    End If
End If

End Sub

Private Sub txtUbicacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPresupuesto.SetFocus
End Sub


Private Function fxGuardarA() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarA = 0
vGridA.Row = vGridA.ActiveRow
vGridA.Col = 1


If lblEstado.Tag <> "P" Then
   MsgBox "Este Contrato ya se finiquito, no se puede modificar...", vbExclamation
   Exit Function
End If


strSQL = "select coalesce(count(*),0) as Existe from Activos_obras_ade" _
       & " where cod_adendum = '" & vGridA.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGridA.Text) = "" Then Exit Function
  
  strSQL = "insert into Activos_obras_ade(cod_adendum,contrato,descripcion,fecha,monto) values('" _
         & UCase(vGridA.Text) & "','" & vCodigo & "','"
  vGridA.Col = 2
  strSQL = strSQL & UCase(vGridA.Text) & "','"
  vGridA.Col = 3
  strSQL = strSQL & Format(vGridA.Text, "yyyy/mm/dd") & "',"
  vGridA.Col = 4
  strSQL = strSQL & CCur(vGridA.Text) & ")"
  
  Call ConectionExecute(strSQL)

  strSQL = "update Activos_obras set addendums = addendums + " & CCur(vGridA.Text) _
         & ",presu_actual = presu_actual + " & CCur(vGridA.Text) _
         & " where contrato = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  vGridA.Col = 1
'  Call sbBitacora("Registra", "Departamento : " & vGridA.Text)

Else 'Actualizar
  MsgBox "No se puede modificar la informacion procesada...", vbExclamation
End If
rs.Close

fxGuardarA = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Function


Private Sub vGridA_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
'MsgBox "Columna : " & vGridA.Col
'MsgBox "Columna Activa: " & vGridA.ActiveCol
'MsgBox "Fila : " & vGridA.Row
'MsgBox "Fila Activa: " & vGridA.ActiveRow

If vGridA.ActiveCol = vGridA.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardarA
  If i = 0 Then Exit Sub
  vGridA.Row = vGridA.ActiveRow
  If vGridA.MaxRows <= vGridA.ActiveRow Then
    vGridA.MaxRows = vGridA.MaxRows + 1
    vGridA.Row = vGridA.MaxRows
  End If
End If

End Sub


Private Function fxGuardarD() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarD = 0
vGridD.Row = vGridD.ActiveRow
vGridD.Col = 1


If lblEstado.Tag <> "P" Then
   MsgBox "Este Contrato ya se finiquito, no se puede modificar...", vbExclamation
   Exit Function
End If


If Trim(vGridD.Text) = "" Then 'Insertar
  
    strSQL = "select coalesce(count(*),0) + 1 as Secuencia from Activos_obras_desem" _
           & " where contrato = '" & vCodigo & "'"
    Call OpenRecordSet(rs, strSQL, 0)
  
  strSQL = "insert into Activos_obras_desem(secuencia,contrato,cod_desembolso,cod_proveedor,documento,fecha,monto) values(" _
         & rs!secuencia & ",'" & vCodigo & "','"
  vGridD.Col = 2
  strSQL = strSQL & Trim(vGridD.Text) & "','"
  vGridD.Col = 3
  strSQL = strSQL & Trim(vGridD.Text) & "','"
  vGridD.Col = 4
  strSQL = strSQL & Trim(vGridD.Text) & "','"
  vGridD.Col = 5
  strSQL = strSQL & Format(vGridD.Text, "yyyy/mm/dd") & "',"
  vGridD.Col = 6
  strSQL = strSQL & CCur(vGridD.Text) & ")"
    
  Call ConectionExecute(strSQL)

  strSQL = "update Activos_obras set desembolsado = desembolsado + " & CCur(vGridD.Text) _
         & ",presu_actual = presu_actual - " & CCur(vGridD.Text) _
         & " where contrato = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  vGridD.Col = 1
'  Call sbBitacora("Registra", "Departamento : " & vGridD.Text)

Else 'Actualizar
  MsgBox "No se puede modificar la informacion procesada...", vbExclamation
End If
rs.Close

fxGuardarD = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Function


Private Sub vGridD_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGridD.ActiveCol = vGridD.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardarD
  If i = 0 Then Exit Sub
  vGridD.Row = vGridD.ActiveRow
  If vGridD.MaxRows <= vGridD.ActiveRow Then
    vGridD.MaxRows = vGridD.MaxRows + 1
    vGridD.Row = vGridD.MaxRows
  End If
End If

If vGridD.ActiveCol = 2 And KeyCode = vbKeyF4 Then
 'Busca Tipos de Desembolsos
 gBusquedas.Resultado = ""
 gBusquedas.Resultado2 = ""
 gBusquedas.Columna = "cod_desembolso"
 gBusquedas.Orden = "cod_desembolso"
 gBusquedas.Consulta = "select cod_desembolso,descripcion from Activos_obras_tDesem"
 gBusquedas.Filtro = ""
 frmBusquedas.Show vbModal
 vGridD.Col = 2
 vGridD.Row = vGridD.ActiveRow
 vGridD.Text = gBusquedas.Resultado
 vGridD.TextTip = TextTipFloating
 vGridD.CellNote = gBusquedas.Resultado2
End If

If vGridD.ActiveCol = 3 And KeyCode = vbKeyF4 Then
 'Busca Proveedores
 gBusquedas.Resultado = ""
 gBusquedas.Resultado2 = ""
 gBusquedas.Columna = "descripcion"
 gBusquedas.Orden = "descripcion"
 gBusquedas.Consulta = "select cod_proveedor,descripcion from Activos_proveedores"
 gBusquedas.Filtro = ""
 frmBusquedas.Show vbModal
 vGridD.Col = 3
 vGridD.Row = vGridD.ActiveRow
 vGridD.Text = gBusquedas.Resultado
 vGridD.TextTip = TextTipFloating
 vGridD.CellNote = gBusquedas.Resultado2
End If


End Sub

