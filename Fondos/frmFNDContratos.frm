VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.TaskPanel.v24.0.0.ocx"
Begin VB.Form frmFNDContratos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contratos"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   Icon            =   "frmFNDContratos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   11970
   Begin XtremeTaskPanel.TaskPanel tpMain 
      Height          =   6888
      Left            =   0
      TabIndex        =   105
      Top             =   840
      Width           =   2760
      _Version        =   1572864
      _ExtentX        =   4868
      _ExtentY        =   12150
      _StockProps     =   64
      VisualTheme     =   17
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeSuiteControls.PushButton btnCalculadora 
      Height          =   375
      Left            =   9600
      TabIndex        =   128
      Top             =   0
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Calculadora de Inversiones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Picture         =   "frmFNDContratos.frx":08CA
   End
   Begin XtremeSuiteControls.GroupBox gb_Main 
      Height          =   2292
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   9132
      _Version        =   1572864
      _ExtentX        =   16108
      _ExtentY        =   4043
      _StockProps     =   79
      Caption         =   "Contrato.:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboOperadora 
         Height          =   330
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   7695
         _Version        =   1572864
         _ExtentX        =   13573
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
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   1440
         TabIndex        =   60
         Top             =   840
         Width           =   1332
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtContrato 
         Height          =   312
         Left            =   1440
         TabIndex        =   61
         Top             =   1200
         Width           =   1332
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   315
         Left            =   7680
         TabIndex        =   62
         Top             =   840
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   2760
         TabIndex        =   63
         Top             =   840
         Width           =   4932
         _Version        =   1572864
         _ExtentX        =   8700
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   312
         Left            =   1440
         TabIndex        =   71
         Top             =   1560
         Width           =   1332
         _Version        =   1572864
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFecha 
         Height          =   312
         Left            =   1440
         TabIndex        =   72
         Top             =   1920
         Width           =   1332
         _Version        =   1572864
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAcumAportes 
         Height          =   312
         Left            =   4080
         TabIndex        =   73
         Top             =   1200
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAcumRend 
         Height          =   312
         Left            =   4080
         TabIndex        =   77
         Top             =   1560
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAcumTotal 
         Height          =   312
         Left            =   4080
         TabIndex        =   78
         Top             =   1920
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtMontoTransito 
         Height          =   315
         Left            =   6960
         TabIndex        =   80
         Top             =   1560
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDisponible 
         Height          =   315
         Left            =   6960
         TabIndex        =   81
         Top             =   1920
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtCuentaCliente 
         Height          =   315
         Left            =   6960
         TabIndex        =   79
         Top             =   1200
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
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
         Index           =   24
         Left            =   5880
         TabIndex        =   84
         Top             =   1920
         Width           =   1452
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "En Tránsito:"
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
         Index           =   23
         Left            =   5880
         TabIndex        =   83
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label Label7 
         Caption         =   "IBAN:"
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
         Index           =   25
         Left            =   5880
         TabIndex        =   82
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Aportes"
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
         Index           =   21
         Left            =   3000
         TabIndex        =   76
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Rendimiento"
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
         Index           =   20
         Left            =   3000
         TabIndex        =   75
         Top             =   1560
         Width           =   1452
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   19
         Left            =   3000
         TabIndex        =   74
         Top             =   1920
         Width           =   1332
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Index           =   0
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   9252
         _Version        =   1572864
         _ExtentX        =   16319
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Contrato"
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
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   1212
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Operadora"
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
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   1212
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   1212
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7725
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Usuario de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3246
            MinWidth        =   3246
            Object.ToolTipText     =   "Fecha de Modificacion"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Usuario de Modificacion"
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
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5172
      Left            =   2760
      TabIndex        =   9
      Top             =   2880
      Width           =   9132
      _Version        =   1572864
      _ExtentX        =   16108
      _ExtentY        =   9123
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
      PaintManager.Position=   2
      ItemCount       =   9
      Item(0).Caption =   "General"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "gb_General"
      Item(1).Caption =   "Adicional"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gb_Complementarios"
      Item(2).Caption =   "Destinos"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "gb_Destinos"
      Item(3).Caption =   "Beneficiarios"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "gb_Beneficiarios"
      Item(4).Caption =   "Sub Cuentas"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "gb_SubCuentas"
      Item(5).Caption =   "Retiros"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "gb_Retiros"
      Item(6).Caption =   "Cupones"
      Item(6).ControlCount=   1
      Item(6).Control(0)=   "gb_Cupones"
      Item(7).Caption =   "Bitácora"
      Item(7).ControlCount=   2
      Item(7).Control(0)=   "ShortcutCaption1(9)"
      Item(7).Control(1)=   "lswBitacora"
      Item(8).Caption =   "TP"
      Item(8).ControlCount=   1
      Item(8).Control(0)=   "gbTasaPreferencial"
      Begin XtremeSuiteControls.ListView lswBitacora 
         Height          =   4215
         Left            =   -70000
         TabIndex        =   107
         Top             =   600
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   7435
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
      Begin XtremeSuiteControls.GroupBox gb_General 
         Height          =   4650
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   9135
         _Version        =   1572864
         _ExtentX        =   16113
         _ExtentY        =   8202
         _StockProps     =   79
         Caption         =   "Datos Generales .:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboPlazoInversion 
            Height          =   330
            Left            =   1680
            TabIndex        =   110
            Top             =   2160
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.CheckBox chkDeducePlanilla 
            Height          =   255
            Left            =   6720
            TabIndex        =   90
            Top             =   1800
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Deducir por Planilla?"
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
            Appearance      =   17
         End
         Begin XtremeSuiteControls.GroupBox gbCupones 
            Height          =   1095
            Left            =   240
            TabIndex        =   11
            Top             =   3120
            Width           =   8295
            _Version        =   1572864
            _ExtentX        =   14631
            _ExtentY        =   1931
            _StockProps     =   79
            Caption         =   "Programación de Cupones "
            ForeColor       =   8388608
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
            BorderStyle     =   1
            Begin XtremeSuiteControls.CheckBox chkCuponPaga 
               Height          =   495
               Left            =   480
               TabIndex        =   111
               Top             =   600
               Width           =   1215
               _Version        =   1572864
               _ExtentX        =   2143
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Paga Cupón?"
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
               Alignment       =   1
            End
            Begin XtremeSuiteControls.ComboBox cboCuponFrecuencia 
               Height          =   330
               Left            =   1800
               TabIndex        =   12
               Top             =   720
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3625
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
            Begin XtremeSuiteControls.FlatEdit txtCuponProximo 
               Height          =   330
               Left            =   3840
               TabIndex        =   91
               Top             =   720
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
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
               BackColor       =   16777215
               Alignment       =   2
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCuponUltimo 
               Height          =   330
               Left            =   6000
               TabIndex        =   92
               Top             =   720
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
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
               BackColor       =   16777215
               Alignment       =   2
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label7 
               Caption         =   "Cupón:"
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
               Index           =   11
               Left            =   1800
               TabIndex        =   15
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Próximo Pago.:"
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
               Index           =   13
               Left            =   3960
               TabIndex        =   14
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label7 
               Caption         =   "Ultimo Pago.:"
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
               Index           =   14
               Left            =   6000
               TabIndex        =   13
               Top             =   480
               Width           =   1215
            End
         End
         Begin XtremeSuiteControls.ComboBox cboPlazo 
            Height          =   312
            Left            =   2400
            TabIndex        =   16
            Top             =   2160
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1720
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
         Begin XtremeSuiteControls.ComboBox cboVendedor 
            Height          =   330
            Left            =   1680
            TabIndex        =   17
            Top             =   840
            Width           =   7215
            _Version        =   1572864
            _ExtentX        =   12726
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
         Begin XtremeSuiteControls.FlatEdit txtCedula 
            Height          =   330
            Left            =   1680
            TabIndex        =   64
            Top             =   480
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
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
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   330
            Left            =   3360
            TabIndex        =   65
            Top             =   480
            Width           =   5535
            _Version        =   1572864
            _ExtentX        =   9763
            _ExtentY        =   582
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   312
            Left            =   1680
            TabIndex        =   66
            Top             =   1440
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInversion 
            Height          =   312
            Left            =   1680
            TabIndex        =   67
            Top             =   1800
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
            Height          =   315
            Left            =   4680
            TabIndex        =   68
            Top             =   1200
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2984
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpCorte 
            Height          =   315
            Left            =   1680
            TabIndex        =   69
            Top             =   2520
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.FlatEdit txtPlazo 
            Height          =   330
            Left            =   1680
            TabIndex        =   70
            Top             =   2160
            Width           =   732
            _Version        =   1572864
            _ExtentX        =   1291
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTasa 
            Height          =   312
            Left            =   4680
            TabIndex        =   85
            Top             =   1440
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtPtsAdd 
            Height          =   312
            Left            =   4680
            TabIndex        =   86
            Top             =   1800
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtIntereses 
            Height          =   312
            Left            =   4680
            TabIndex        =   87
            Top             =   2160
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtTasaTipo 
            Height          =   312
            Left            =   4680
            TabIndex        =   88
            Top             =   2520
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtOperacion 
            Height          =   315
            Left            =   6720
            TabIndex        =   89
            Top             =   2520
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkTasaPreferencial 
            Height          =   255
            Left            =   6720
            TabIndex        =   112
            ToolTipText     =   "Este Contrato Tiene Tasa Preferencial"
            Top             =   1440
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tasa Preferencial?"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   372
            Index           =   1
            Left            =   -120
            TabIndex        =   53
            Top             =   0
            Width           =   9252
            _Version        =   1572864
            _ExtentX        =   16319
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Datos Generales .:"
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
         Begin VB.Label Label6 
            Caption         =   "Identificación"
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
            Left            =   360
            TabIndex        =   29
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Ejecutivo"
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
            Left            =   360
            TabIndex        =   28
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "No.Operación:"
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
            Index           =   17
            Left            =   6720
            TabIndex        =   27
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Tasa Tipo"
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
            Left            =   3600
            TabIndex        =   26
            Top             =   2520
            Width           =   852
         End
         Begin VB.Label Label7 
            Caption         =   "Pts.Add."
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
            Index           =   12
            Left            =   3600
            TabIndex        =   25
            Top             =   1800
            Width           =   852
         End
         Begin VB.Label Label7 
            Caption         =   "Vencimiento"
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
            Left            =   360
            TabIndex        =   24
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label lblMensualidad 
            Caption         =   "Mensualidad"
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
            Left            =   360
            TabIndex        =   23
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Plazo"
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
            Left            =   360
            TabIndex        =   22
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label lblInversion 
            Caption         =   "Inversión"
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
            Left            =   360
            TabIndex        =   21
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Tasa Ref."
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
            Index           =   9
            Left            =   3600
            TabIndex        =   20
            Top             =   1440
            Width           =   852
         End
         Begin VB.Label Label7 
            Caption         =   "Intereses"
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
            Index           =   10
            Left            =   3600
            TabIndex        =   19
            Top             =   2160
            Width           =   852
         End
         Begin VB.Label lblPorcentajeDesc 
            Caption         =   "Porcentaje"
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
            Left            =   3600
            TabIndex        =   18
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.GroupBox gb_SubCuentas 
         Height          =   4416
         Left            =   -70000
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   9372
         _Version        =   1572864
         _ExtentX        =   16531
         _ExtentY        =   7789
         _StockProps     =   79
         Caption         =   "Sub Cuentas .:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin FPSpreadADO.fpSpread vGrid 
            Height          =   3492
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   8892
            _Version        =   524288
            _ExtentX        =   15684
            _ExtentY        =   6159
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
            SpreadDesigner  =   "frmFNDContratos.frx":0D64
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   372
            Index           =   5
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   9132
            _Version        =   1572864
            _ExtentX        =   16108
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Sub Cuentas.:"
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
      Begin XtremeSuiteControls.GroupBox gb_Retiros 
         Height          =   4296
         Left            =   -70000
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1572864
         _ExtentX        =   16108
         _ExtentY        =   7578
         _StockProps     =   79
         Caption         =   "Retiros y Liquidaciones .:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   3012
            Left            =   0
            TabIndex        =   103
            Top             =   480
            Width           =   9012
            _Version        =   1572864
            _ExtentX        =   15896
            _ExtentY        =   5313
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
         Begin XtremeSuiteControls.PushButton cmdReversion 
            Height          =   615
            Left            =   5760
            TabIndex        =   33
            Top             =   3600
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Reversión"
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
            Picture         =   "frmFNDContratos.frx":13BF
         End
         Begin XtremeSuiteControls.PushButton cmdBoleta 
            Height          =   615
            Left            =   7440
            TabIndex        =   34
            Top             =   3600
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Boleta"
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
            Picture         =   "frmFNDContratos.frx":1D4C
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   372
            Index           =   6
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   9132
            _Version        =   1572864
            _ExtentX        =   16108
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Retiros y Liquidaciones.:"
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
         Begin VB.Label lblRetLiq 
            Caption         =   "Seleccione un retiro o liquidación ?"
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
            Left            =   360
            TabIndex        =   35
            Top             =   3720
            Width           =   4212
         End
      End
      Begin XtremeSuiteControls.GroupBox gb_Beneficiarios 
         Height          =   4296
         Left            =   -70000
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1572864
         _ExtentX        =   16108
         _ExtentY        =   7578
         _StockProps     =   79
         Caption         =   "Beneficiarios .:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswBeneficiarios 
            Height          =   3012
            Left            =   0
            TabIndex        =   102
            Top             =   480
            Width           =   9012
            _Version        =   1572864
            _ExtentX        =   15896
            _ExtentY        =   5313
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
         Begin XtremeSuiteControls.PushButton cmdEditar 
            Height          =   615
            Left            =   7440
            TabIndex        =   37
            Top             =   3600
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Editar"
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
            Picture         =   "frmFNDContratos.frx":2508
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   372
            Index           =   4
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   9132
            _Version        =   1572864
            _ExtentX        =   16108
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Beneficiarios.:"
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
      Begin XtremeSuiteControls.GroupBox gb_Destinos 
         Height          =   4650
         Left            =   -70000
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1572864
         _ExtentX        =   16113
         _ExtentY        =   8202
         _StockProps     =   79
         Caption         =   "Destinos o Plan de Inversión .:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin FPSpreadADO.fpSpread gDestinos 
            Height          =   4095
            Left            =   0
            TabIndex        =   130
            Top             =   480
            Width           =   9135
            _Version        =   524288
            _ExtentX        =   16113
            _ExtentY        =   7223
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
            SpreadDesigner  =   "frmFNDContratos.frx":2EDB
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   372
            Index           =   3
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   9132
            _Version        =   1572864
            _ExtentX        =   16108
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Destino o Plan de Inversión.:"
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
      Begin XtremeSuiteControls.GroupBox gb_Complementarios 
         Height          =   4296
         Left            =   -70000
         TabIndex        =   39
         Top             =   120
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1572864
         _ExtentX        =   16108
         _ExtentY        =   7578
         _StockProps     =   79
         Caption         =   "Datos Complementarios.:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1212
            Left            =   360
            TabIndex        =   40
            Top             =   600
            Width           =   8532
            _Version        =   1572864
            _ExtentX        =   15049
            _ExtentY        =   2138
            _StockProps     =   79
            Caption         =   "Incrementos y Renovaciones"
            ForeColor       =   8388608
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
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox cboIncTipo 
               Height          =   312
               Left            =   3000
               TabIndex        =   93
               Top             =   360
               Width           =   1332
               _Version        =   1572864
               _ExtentX        =   2355
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
            Begin XtremeSuiteControls.ComboBox cboRenueva 
               Height          =   312
               Left            =   6480
               TabIndex        =   94
               Top             =   360
               Width           =   1332
               _Version        =   1572864
               _ExtentX        =   2355
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
            Begin XtremeSuiteControls.FlatEdit txtIncAnual 
               Height          =   312
               Left            =   3000
               TabIndex        =   97
               Top             =   720
               Width           =   1332
               _Version        =   1572864
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtExc 
               Height          =   312
               Left            =   6480
               TabIndex        =   98
               Top             =   720
               Width           =   1332
               _Version        =   1572864
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Cap. Exc. (%)"
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
               Left            =   5160
               TabIndex        =   44
               Top             =   720
               Width           =   1212
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor Incremento"
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
               Left            =   1320
               TabIndex        =   43
               Top             =   720
               Width           =   1572
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Renovación"
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
               Index           =   6
               Left            =   5040
               TabIndex        =   42
               Top             =   360
               Width           =   1332
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Tipo de Incremento"
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
               Left            =   1320
               TabIndex        =   41
               Top             =   360
               Width           =   1572
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   1092
            Left            =   360
            TabIndex        =   45
            Top             =   1920
            Width           =   8532
            _Version        =   1572864
            _ExtentX        =   15049
            _ExtentY        =   1926
            _StockProps     =   79
            Caption         =   "Información para Desembolsos "
            ForeColor       =   8388608
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
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox cboBanco 
               Height          =   330
               Left            =   1800
               TabIndex        =   95
               Top             =   360
               Width           =   6615
               _Version        =   1572864
               _ExtentX        =   11668
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
            Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
               Height          =   312
               Left            =   1800
               TabIndex        =   96
               Top             =   720
               Width           =   1932
               _Version        =   1572864
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
            Begin XtremeSuiteControls.ComboBox cboCuenta 
               Height          =   330
               Left            =   4920
               TabIndex        =   109
               Top             =   720
               Width           =   3495
               _Version        =   1572864
               _ExtentX        =   6165
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
            Begin VB.Label Label8 
               Caption         =   "Documento"
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
               Left            =   480
               TabIndex        =   48
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label9 
               Caption         =   "Banco"
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
               Left            =   480
               TabIndex        =   47
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Cuenta "
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
               Left            =   4080
               TabIndex        =   46
               Top             =   720
               Width           =   852
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   1455
            Left            =   360
            TabIndex        =   49
            Top             =   3240
            Width           =   8535
            _Version        =   1572864
            _ExtentX        =   15055
            _ExtentY        =   2566
            _StockProps     =   79
            Caption         =   "Información para Pago a Terceros (Albacea)"
            ForeColor       =   8388608
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
            BorderStyle     =   1
            Begin XtremeSuiteControls.FlatEdit txtAlbaceaCed 
               Height          =   312
               Left            =   1680
               TabIndex        =   99
               Top             =   480
               Width           =   1692
               _Version        =   1572864
               _ExtentX        =   2984
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
            Begin XtremeSuiteControls.FlatEdit txtAlbaceaNom 
               Height          =   312
               Left            =   3360
               TabIndex        =   100
               Top             =   480
               Width           =   5052
               _Version        =   1572864
               _ExtentX        =   8911
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.CheckBox chkPagoTercero 
               Height          =   252
               Left            =   1680
               TabIndex        =   101
               Top             =   840
               Width           =   5652
               _Version        =   1572864
               _ExtentX        =   9970
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Activar el Pago a Tercero en Retiros/Liquidaciones"
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
               Appearance      =   16
            End
            Begin VB.Label Label6 
               Caption         =   "Cédula"
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
               Left            =   840
               TabIndex        =   50
               Top             =   480
               Width           =   852
            End
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   372
            Index           =   2
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   9132
            _Version        =   1572864
            _ExtentX        =   16108
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Datos Complementarios.:"
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
      Begin XtremeSuiteControls.GroupBox gb_Cupones 
         Height          =   4296
         Left            =   -70000
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1572864
         _ExtentX        =   16108
         _ExtentY        =   7578
         _StockProps     =   79
         Caption         =   "Cupones del Plan.:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswCupones 
            Height          =   3735
            Left            =   0
            TabIndex        =   104
            Top             =   480
            Width           =   9015
            _Version        =   1572864
            _ExtentX        =   15901
            _ExtentY        =   6588
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   372
            Index           =   7
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   9132
            _Version        =   1572864
            _ExtentX        =   16108
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Cupones.:"
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
      Begin XtremeSuiteControls.GroupBox gbTasaPreferencial 
         Height          =   4695
         Left            =   -70000
         TabIndex        =   113
         Top             =   120
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1572864
         _ExtentX        =   16113
         _ExtentY        =   8281
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnTP_Refresh 
            Height          =   375
            Left            =   3840
            TabIndex        =   114
            ToolTipText     =   "Revisa si fue autorizada!"
            Top             =   2400
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Estado Gestión"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmFNDContratos.frx":34AD
         End
         Begin XtremeSuiteControls.FlatEdit txtGestionId 
            Height          =   315
            Left            =   6600
            TabIndex        =   115
            Top             =   615
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtGestionEstado 
            Height          =   315
            Left            =   6600
            TabIndex        =   116
            Top             =   975
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTP_Calculada 
            Height          =   330
            Left            =   1920
            TabIndex        =   117
            Top             =   600
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
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
            Text            =   "0"
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTP_Margen 
            Height          =   330
            Left            =   1920
            TabIndex        =   118
            Top             =   960
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
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
            Text            =   "0"
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTP_Solicitada 
            Height          =   330
            Left            =   1920
            TabIndex        =   119
            Top             =   1320
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
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
            Text            =   "0"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnTP_Solicitar 
            Height          =   375
            Left            =   2160
            TabIndex        =   120
            ToolTipText     =   "Revisa si fue autorizada!"
            Top             =   2400
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Solicitar"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmFNDContratos.frx":3BAD
         End
         Begin XtremeSuiteControls.PushButton btnTP_Cerrar 
            Height          =   375
            Left            =   5640
            TabIndex        =   127
            ToolTipText     =   "Revisa si fue autorizada!"
            Top             =   2400
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Cerrar"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmFNDContratos.frx":42C6
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   375
            Left            =   0
            TabIndex        =   126
            Top             =   0
            Width           =   9135
            _Version        =   1572864
            _ExtentX        =   16113
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Solicitud de Tasa Preferencial"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   125
            Top             =   600
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tasa Calculada:"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   124
            Top             =   960
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Margen Permitido:"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   123
            Top             =   1320
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tasa Solicitada:"
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
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Estado de la Gestión ..:"
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
            Left            =   3960
            TabIndex        =   122
            Top             =   975
            Width           =   2415
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión Id..:"
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
            Left            =   3960
            TabIndex        =   121
            Top             =   600
            Width           =   2415
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   9
         Left            =   -70000
         TabIndex        =   108
         Top             =   120
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1572864
         _ExtentX        =   16108
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Bitácora Especial de Cambios del Contrato:"
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
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   3240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":4904
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":4B79
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":4E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":4F95
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":512D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":52D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":5477
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":5602
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":578D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":5891
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":5B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":5C2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":5EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":5F6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":610E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDContratos.frx":62AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnTasaPreferencial 
      Height          =   375
      Left            =   6960
      TabIndex        =   129
      ToolTipText     =   "Revisa si fue autorizada!"
      Top             =   0
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Solicita Tasa Preferencial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Picture         =   "frmFNDContratos.frx":6355
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Index           =   8
      Left            =   0
      TabIndex        =   106
      Top             =   480
      Width           =   2772
      _Version        =   1572864
      _ExtentX        =   4890
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Detalle:"
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
   End
End
Attribute VB_Name = "frmFNDContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vGuardar As Boolean
Dim vMontoMin As Currency, vPlazoMin As Long, vInversionMin As Currency, vSubCuentas As Boolean
Dim vBusqueda As String, vPaso As Boolean, vTipoCDP As Boolean, vCDPCuponesAplica As Boolean
Dim vTasaPreferencial As Boolean, vTasaMargenNegociacion As Currency

Private Type FndCambios
  vPlazo        As Integer
  vCuota        As Currency
  vInversion    As Currency
  vDescPlazo    As String
  vDedPlanilla  As Integer
End Type

Dim vCambios As FndCambios
Dim vCarga As Boolean

Dim mTipoDeduc As String, mPorcRef As Currency

Dim strSQL As String, rs As New ADODB.Recordset

Const Id_TaskItem_General = 0
Const Id_TaskItem_Complementarios = 1
Const Id_TaskItem_Destinos = 2
Const Id_TaskItem_Beneficiarios = 3
Const Id_TaskItem_SubCuentas = 4
Const Id_TaskItem_Retiros = 5
Const Id_TaskItem_Cupones = 6
Const Id_TaskItem_Bitacora = 7


Private Sub sbTaskPanel_Load()


    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    tpMain.VisualTheme = xtpTaskPanelThemeOffice2016
  
    Set Group = tpMain.Groups.Add(0, "Registro")
    Group.ToolTip = "Información Principal para el Registro del Contrato"
    Group.Special = True

    
    Group.Items.Add Id_TaskItem_General, "General", xtpTaskItemTypeLink, 4
    Group.Items.Add Id_TaskItem_Complementarios, "Complementarios", xtpTaskItemTypeLink, 8
    
    Set Group = tpMain.Groups.Add(0, "Detalles")
    Group.ToolTip = "Datos adicionales del Plan"
    
    Group.Items.Add Id_TaskItem_Destinos, "Destinos", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Beneficiarios, "Beneficiarios", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_SubCuentas, "Sub Cuentas", xtpTaskItemTypeLink, 9
    
    Set Group = tpMain.Groups.Add(0, "Info adicional")
    Group.Items.Add Id_TaskItem_Retiros, "Retiros", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Cupones, "Cupones", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Bitacora, "Bitácora", xtpTaskItemTypeLink, 3
    
    tpMain.SetImageList imlTaskPanelIcons
    
'    tpMain.SetMargins 5, 5, 5, 5, 5
   

End Sub


Private Sub sbTaskPanel_Accion(ItemId As Integer)

Select Case ItemId
  Case Id_TaskItem_General  'General
    tcMain.Item(0).Selected = True
  
  Case Id_TaskItem_Complementarios 'Complementarios
    tcMain.Item(1).Selected = True
  
  Case Id_TaskItem_Destinos 'Destinos
      
      Call sbConsulta_Detalle("Destinos")
  
  Case Id_TaskItem_Beneficiarios 'Beneficiaros

      Call sbConsulta_Detalle("Beneficiarios")
  
  Case Id_TaskItem_SubCuentas 'Sub Cuentas

      Call sbConsulta_Detalle("Sub Cuentas")
  
  Case Id_TaskItem_Retiros 'Retiros

      Call sbConsulta_Detalle("Retiros")
  
  Case Id_TaskItem_Cupones 'Cupones

      Call sbConsulta_Detalle("Cupones")
      
      
  Case Id_TaskItem_Bitacora 'Bitacora

      Call sbConsulta_Detalle("Bitacora")
     
  
End Select



End Sub


Function fxConsecutivoContrato() As Long
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select isnull(Consecutivo,0) + 1 as Seq From fnd_planes where cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " and cod_plan = '" & Trim(txtCodigo) & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   fxConsecutivoContrato = rs!Seq
   strSQL = "Update fnd_planes set Consecutivo = " & rs!Seq _
          & " Where cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
          & " and cod_plan='" & Trim(txtCodigo) & "'"
   Call ConectionExecute(strSQL)
End If
rs.Close

End Function

Private Function fxVerificaDatos() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, i As Integer

On Error GoTo vError

vMensaje = ""

'Verifica si no hay otro contrato activo
If vEdita = False Then
    'Selecciona el numero de contraos activos por persona por plan
    strSQL = "select num_contratos_activos from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " and cod_plan = '" & txtCodigo & "'"
    Call OpenRecordSet(rs, strSQL)
        i = rs!NUM_CONTRATOS_ACTIVOS
    rs.Close
    

    strSQL = "select isnull(count(*),0) as existe from fnd_contratos" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " and estado = 'A' and cedula = '" & txtCedula.Text & "' and cod_plan = '" & txtCodigo.Text & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe >= i Then
      vMensaje = " - Esta persona ha superado el número máximo de contratos activos en este plan..." & vbCrLf
    End If
    rs.Close
    
End If


strSQL = "exec dbo.spFND_ValidaEstados '" & txtCodigo.Text & "'," & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!encontrado = 0 Then
    vMensaje = " - El estado de esta persona no aplica para en este plan o el Plan se encuentra inactivo..." & vbCrLf
End If


strSQL = "select dbo.fxFnd_Contrato_Valida_Plazo(" & cboOperadora.ItemData(cboOperadora.ListIndex) & " ,'" & txtCodigo.Text _
       & "', " & txtPlazo.Text & ") as 'Plazo_Valida'" _
       & ", dbo.fxFnd_Seguridad_Acceso_Planes('" & glogon.Usuario & "', " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & ", '" & txtCodigo.Text & "') as 'Acceso_Valida'"


Call OpenRecordSet(rs, strSQL)

If rs!Plazo_Valida = 0 Then
    vMensaje = " - El Plazo se encuentra fuera del Rango Permitido por el Plan..." & vbCrLf
End If

If rs!Acceso_Valida = 0 Then
    vMensaje = " - El usuario no tiene Autorización para gestionar este Plan..." & vbCrLf
End If


If txtCodigo.Text = "" Then vMensaje = vMensaje & " - Indique el Plan..." & vbCrLf
If txtCedula.Text = "" Then vMensaje = vMensaje & " - Especifique la persona ..." & vbCrLf
If Not IsNumeric(txtPorcentaje.Text) Then vMensaje = vMensaje & " - El Porcentaje de deducción no es válido..." & vbCrLf
If Not IsNumeric(txtMonto.Text) Then vMensaje = vMensaje & " - La cuota especificada no es válida..." & vbCrLf
If Not IsNumeric(txtInversion.Text) Then vMensaje = vMensaje & " - La inversión especificada no es válida..." & vbCrLf
If Not IsNumeric(txtPlazo.Text) Then vMensaje = vMensaje & " - El plazo especificado no es válido..." & vbCrLf
If Not IsNumeric(txtIncAnual.Text) Then vMensaje = vMensaje & " - El % de Incremento anual no es válido..." & vbCrLf
If Not IsNumeric(txtExc.Text) Then vMensaje = vMensaje & " - El % de Capitalización no es válido..." & vbCrLf



If txtTasa = "" Then txtTasa = 0

If mTipoDeduc = "P" And Len(vMensaje) = 0 Then
   If CCur(txtPorcentaje.Text) > 100 Or CCur(txtPorcentaje.Text) < 0 Then vMensaje = vMensaje & "El Porcentaje de Deducción no es válido!"
End If

'Verifica Montos Minimos
If Len(vMensaje) = 0 Then
   strSQL = "select PLAZO_MINIMO * case when PLAZO_TIPO = 'M' then 30 else 1 end as 'Plazo_Minimo', MONTO_MINIMO, INVERSION_MINIMO" _
          & " from fnd_Planes where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
          & " and cod_plan = '" & txtCodigo.Text & "'"
   Call OpenRecordSet(rs, strSQL)
   If Mid(cboPlazo.Text, 1, 1) = "D" Then
        If CLng(txtPlazo.Text) < rs!Plazo_Minimo Then vMensaje = vMensaje & " - El Plazo no cumple con el plazo mínimo permitido..." & vbCrLf
   End If
   If Mid(cboPlazo.Text, 1, 1) = "M" Then
        If CLng(txtPlazo.Text) * 30 < rs!Plazo_Minimo Then vMensaje = vMensaje & " - El Plazo no cumple con el plazo mínimo permitido.." & vbCrLf
   End If
  If mTipoDeduc = "M" Then
    If txtMonto.Visible And CCur(txtMonto.Text) < rs!MONTO_MINIMO Then vMensaje = vMensaje & " - El monto es menor al mínimo permitido..." & vbCrLf
    If cboPlazoInversion.Visible And CCur(txtInversion.Text) < rs!INVERSION_MINIMO Then vMensaje = vMensaje & " - El monto de la INVERSIÓN es menor al mínimo permitido..." & vbCrLf
  End If
End If



If Len(vMensaje) = 0 Then
  fxVerificaDatos = True
Else
  fxVerificaDatos = False
  MsgBox vMensaje, vbExclamation
End If

Exit Function

vError:
  vMensaje = vMensaje & " - Error de Procesamiento!"
  fxVerificaDatos = False
  MsgBox vMensaje, vbExclamation


End Function


Private Function fxCuponFrecuencia(pDato As String) As String
Dim vResultado As String

  Select Case Trim(pDato)
    Case "N"
      vResultado = "No Aplica"
    Case "M"
      vResultado = "Mensuales"
    Case "T"
      vResultado = "Trimestrales"
    Case "S"
      vResultado = "Semestrales"
    Case "Q"
      vResultado = "Cuatrimestral"
    Case "A"
      vResultado = "Anual"
    Case "V"
      vResultado = "Al Vencimiento"
    
    
    Case "No Aplica"
      vResultado = "N"
    Case "Mensuales"
      vResultado = "M"
    Case "Trimestrales"
      vResultado = "T"
    Case "Semestrales"
      vResultado = "S"
    Case "Cuatrimestral"
      vResultado = "Q"
    Case "Anual"
      vResultado = "A"
    Case "Al Vencimiento"
      vResultado = "V"
    
    Case Else
      If Len(pDato) > 1 Then
          vResultado = "N"
      Else
          vResultado = "No Aplica"
      End If
  End Select

fxCuponFrecuencia = vResultado

End Function



Private Sub sbLimpiaPantalla()


 mTipoDeduc = "M"
 mPorcRef = 0
 vTasaPreferencial = False
 
gb_General.Visible = True
gb_Complementarios.Visible = False


gb_Destinos.Visible = False
gb_Beneficiarios.Visible = False
gb_SubCuentas.Visible = False
gb_Cupones.Visible = False
gb_Retiros.Visible = False

vCodigo = 0
txtContrato.Text = ""
txtEstado.Text = "ACTIVO"
txtCedula.Text = ""
txtCedula.Locked = False

txtNombre = ""
txtNombre.Locked = True

txtMonto = 0
txtInversion = 0
txtPorcentaje = 0


txtPlazo.Text = "30"
cboPlazo.Text = "Días"
txtTasa.Text = "0"

txtDivisa.Text = ""

'cboCuponFrecuencia.Text = "No Aplica"
cboCuponFrecuencia.Locked = True

txtAcumAportes.Text = "0"
txtAcumRend.Text = "0"
txtAcumTotal.Text = "0"
txtMontoTransito.Text = "0"
txtDisponible.Text = "0"

txtCuentaCliente.Text = ""

txtCuponProximo.Text = ""
txtCuponUltimo.Text = ""
txtTasaTipo.Text = "Variable"
txtPtsAdd.Text = "0"

chkDeducePlanilla.Value = vbChecked
chkPagoTercero.Value = vbUnchecked

cboCuenta.Clear
txtExc.Text = "0"
dtpCorte.Value = fxFechaServidor
txtFecha.Text = Format(dtpCorte.Value, "yyyy/mm/dd")

cboRenueva.Text = "SI"
cboIncTipo.Text = "Porcentaje"
txtIncAnual.Text = "0"

cboTipoDocumento.Text = "Transferencia"

txtAlbaceaCed.Text = ""
txtAlbaceaNom.Text = ""

txtInversion.Locked = False
cboPlazo.Locked = False
txtPlazo.Locked = False
dtpCorte.Enabled = True

txtIntereses.Text = "0"

End Sub


Private Sub btnCalculadora_Click()
 
If txtCedula.Text = "" Or txtNombre.Text = "" Then
 MsgBox "Consulte una Persona Primero!", vbExclamation
 Exit Sub
End If

 GLOBALES.gCedulaActual = txtCedula.Text
 Call sbFormsCall("frmFnd_Calculadora_Inversiones")
End Sub

Private Sub btnTasaPreferencial_Click()

On Error GoTo vError

If txtContrato.Text = "0" Or txtContrato.Text = "" Then
  MsgBox "Registre el Contrato Primero, luego indique la solicitud de Tasa Preferencial!", vbInformation
Else
  tcMain.Item(8).Selected = True
  
  txtTP_Calculada.Text = txtTasa.Text
  txtTP_Margen.Text = Format(vTasaMargenNegociacion, "Standard")
  txtTP_Solicitada.Text = txtTasa.Text
      
  txtGestionId.Text = ""
  txtGestionEstado.Text = ""
  
  strSQL = "exec spFnd_TP_Solicitud_Ultima " & cboOperadora.ItemData(cboOperadora.ListIndex) _
        & ", '" & txtCodigo.Text & "', " & txtContrato.Text & ", '" & txtCedula.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then
    txtTP_Calculada.Text = Format(rs!TASA_CALCULADA, "Standard")
    txtTP_Margen.Text = Format(rs!MARGEN_MAXIMO, "Standard")
    txtTP_Solicitada.Text = Format(rs!TASA_SOLICITADA, "Standard")
        
    txtGestionId.Text = CStr(rs!ID_TP)
    txtGestionEstado.Text = rs!ESTADO_DESC
    
  End If
End If

Exit Sub

vError:

End Sub

Private Sub btnTP_Cerrar_Click()
 tcMain.Item(0).Selected = True
End Sub

Private Sub btnTP_Refresh_Click()

On Error GoTo vError

If Not IsNumeric(txtGestionId.Text) Then Exit Sub

strSQL = "exec spFnd_TP_Estado " & txtGestionId.Text
Call OpenRecordSet(rs, strSQL)


txtGestionId.Text = rs!Gestion_Id
txtGestionEstado.Text = rs!Gestion_Estado

If Mid(rs!Gestion_Estado, 1, 1) <> "P" Then
   MsgBox "La Gestión ya fue resuelta!", vbInformation
   Call sbConsultaContrato(txtContrato.Text)
End If

Exit Sub

vError:

End Sub

Private Sub btnTP_Solicitar_Click()


On Error GoTo vError

Dim vMargen As Currency, vMensaje As String
Dim vCuponFrecuencia As Long

If Mid(txtGestionEstado.Text, 1, 1) = "A" Then Exit Sub
If chkTasaPreferencial.Value = xtpChecked Then Exit Sub

vMensaje = ""
vMargen = CCur(txtTP_Solicitada.Text) - CCur(txtTP_Calculada.Text)

If vMargen <= 0 Then
   vMensaje = " - La Tasa Solicitada no puede ser menor o igual a la Tasa Oficial!" & vbCrLf
End If

If vMargen > CCur(txtTP_Margen.Text) Then
   vMensaje = " - La Tasa Solicitada Excede el Margen de Negociación Permitido para este Producto!" & vbCrLf
End If

If Len(vMensaje) > 0 Then
 MsgBox vMensaje, vbExclamation
 Exit Sub
End If


'spFnd_TP_Solicitud (@Operadora smallint, @Plan varchar(10), @Contrato int, @Cedula varchar(20), @TasaCalulada dec(7,2), @MargenMax dec(7,2), @TasaSolicitada dec(7,2)
'                        , @Plazo int, @fCuponId int, @Inversion dec(16,2),  @Usuario varchar(30))
                       
If chkCuponPaga.Value = xtpChecked Then
    vCuponFrecuencia = cboCuponFrecuencia.ItemData(cboCuponFrecuencia.ListIndex)
Else
    vCuponFrecuencia = 0
End If
                       
strSQL = "exec spFnd_TP_Solicitud " & cboOperadora.ItemData(cboOperadora.ListIndex) & ", '" & txtCodigo.Text & "', " & txtContrato.Text _
       & ", '" & txtCedula.Text & "', " & CCur(txtTP_Calculada.Text) & ", " & CCur(txtTP_Margen.Text) & ", " & CCur(txtTP_Solicitada.Text) _
       & ", " & txtPlazo.Text & ", " & cboCuponFrecuencia.ItemData(cboCuponFrecuencia.ListIndex) & ", " & CCur(txtInversion.Text) _
       & ", '" & glogon.Usuario & "'"
       
Call OpenRecordSet(rs, strSQL)


txtGestionId.Text = rs!Gestion_Id
txtGestionEstado.Text = rs!Gestion_Estado

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboBanco_Click()

If vPaso Or cboBanco.ListCount = 0 Or cboBanco.Text = "" Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & txtCedula.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub


Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoDocumento.SetFocus
End Sub


Private Sub cboCuponFrecuencia_Click()
Dim vAddMonths As Integer

If vCarga Or vPaso Then Exit Sub
If cboCuponFrecuencia.ListCount = 0 Then Exit Sub
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


strSQL = "exec spFnd_Cupon_Frecuencia_Meses " & cboCuponFrecuencia.ItemData(cboCuponFrecuencia.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.BOF Then
    vAddMonths = rs!Frecuencia_Meses
Else
    vAddMonths = 0
End If
rs.Close

If chkCuponPaga.Value = xtpUnchecked Then
 'Al Vencimiento
 vAddMonths = 1000
End If

Select Case vAddMonths
  Case 0 'No Definido
      txtCuponProximo.Text = ""
  Case 1000 'Al Vencimiento
      txtCuponProximo.Text = Format(dtpCorte.Value, "yyyy/mm/dd")
  Case Else
    If Trim(txtCuponUltimo.Text) = "" Then
        txtCuponProximo.Text = Format(DateAdd("m", vAddMonths, CDate(txtFecha.Text)), "yyyy/mm/dd")
    Else
        txtCuponProximo.Text = Format(DateAdd("m", vAddMonths, CDate(txtCuponUltimo.Text)), "yyyy/mm/dd")
    End If
End Select

Call txtPlazo_KeyUp(vbKeyF4, 0)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboIncTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtIncAnual.SetFocus
End Sub

Private Sub cboOperadora_Click()

If vPaso Then Exit Sub

txtCodigo_LostFocus
If Trim(txtContrato) <> "" Then Call sbConsultaContrato(Trim(txtContrato))

End Sub


Private Sub cboPlazo_Click()
On Error GoTo vError

If Mid(cboPlazo.Text, 1, 1) = "D" Then
    dtpCorte.Value = DateAdd("d", CDbl(txtPlazo), CDate(txtFecha))
Else
    dtpCorte.Value = DateAdd("m", CDbl(txtPlazo), CDate(txtFecha))
End If

txtTasa = fxTasaRef(txtPlazo, Mid(cboPlazo.Text, 1, 1), txtCodigo, cboOperadora.ItemData(cboOperadora.ListIndex))


If CCur(txtTasa) > 0 Then
   txtIntereses = CCur(txtInversion) * IIf((Mid(cboPlazo.Text, 1, 1) = "D"), CLng(txtPlazo), CLng(txtPlazo) * 30) * CCur(txtTasa) / 36500
   txtIntereses = Format(txtIntereses, "Standard")
End If


vError:
End Sub

Private Sub cboPlazoInversion_Click()
If vPaso Then Exit Sub


If cboPlazoInversion.Visible And cboPlazoInversion.ListIndex > 0 Then
 
 vPaso = True
    strSQL = "exec spFnd_Cupon_Frecuencia " & cboPlazoInversion.ItemData(cboPlazoInversion.ListIndex) & ",  '" & txtCodigo.Text & "'"
    Call sbCbo_Llena_New(cboCuponFrecuencia, strSQL, False, True)
 vPaso = False
 
 
  strSQL = "exec spFnd_Inversion_Plazos_Dias " & cboPlazoInversion.ItemData(cboPlazoInversion.ListIndex)
  Call OpenRecordSet(rs, strSQL)
    If Mid(cboPlazo.Text, 1, 1) = "D" Then
      txtPlazo.Text = rs!PLAZO_DIAS
    Else
      txtPlazo.Text = rs!Plazo_Meses
    End If
  rs.Close
      
 Call txtPlazo_KeyUp(vbKeyF4, 0)
End If



End Sub

Private Sub cboRenueva_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboIncTipo.SetFocus
End Sub


Private Sub cboTipoDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCuenta.SetFocus
End Sub

Private Sub cboVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub chkCuponPaga_Click()
If vPaso Then Exit Sub
Call cboCuponFrecuencia_Click
End Sub

Private Sub cmdBoleta_Click()

Me.MousePointer = vbHourglass

On Error GoTo vError

With frmContenedor.Crt
  .Reset
  .WindowShowGroupTree = True
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "Reportes del Módulo de Fondos"
  
  .Connect = glogon.ConectRPT
    
    .ReportFileName = SIFGlobal.fxPathReportes("Fondos_LiquidacionBoleta.rpt")
    .SelectionFormula = "{FND_LIQUIDACION.CONSEC} =" & lblRetLiq.Tag
    .Formulas(0) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
    
    .SubreportToChange = "sbAsiento"
    If GLOBALES.SysDocVersion = 1 Then
      .StoredProcParam(0) = "LI"
    Else
      .StoredProcParam(0) = "FLIQ"
    End If
  
    .StoredProcParam(1) = lblRetLiq.Tag
    .StoredProcParam(2) = 1

    .PrintReport

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdEditar_Click()

If Trim(txtCedula) = "" Then Exit Sub

If Not fxAplicaBeneficiarios Then
    MsgBox "Para este plan no aplica beneficiarios", vbInformation
    Exit Sub
End If

GLOBALES.gCedulaActual = Trim(txtCedula)

GLOBALES.gTag = ""
GLOBALES.gTag2 = ""
GLOBALES.gTag3 = ""

GLOBALES.gTag = cboOperadora.ItemData(cboOperadora.ListIndex) 'operadora
GLOBALES.gTag2 = txtCodigo ' plan
GLOBALES.gTag3 = txtContrato ' contrato
'frmFNDBeneficiarios_Contratos.lblOperadora.Caption = "Operadora:  " & Trim(cboOperadora.Text)
'frmFNDBeneficiarios_Contratos.lblPlan.Caption = "Plan:  " & Trim(txtCodigo.Text)
'frmFNDBeneficiarios_Contratos.lblContrato.Caption = "Contrato:  " & Trim(txtContrato.Text)
'frmFNDBeneficiarios_Contratos.lblOperadora.Tag = cboOperadora.ItemData(cboOperadora.ListIndex)
'frmFNDBeneficiarios_Contratos.lblPlan.Tag = txtCodigo
'frmFNDBeneficiarios_Contratos.lblContrato.Tag = txtContrato


  Call sbFormsCall("frmFNDBeneficiarios_Contratos", 1, 522, 0, False)
  
End Sub

Private Sub dtpCorte_Change()



On Error GoTo vError

If Mid(cboPlazo.Text, 1, 1) = "D" Then
    txtPlazo = CInt(DateDiff("d", CDate(txtFecha), dtpCorte.Value))
Else
    txtPlazo = CInt(DateDiff("m", CDate(txtFecha), dtpCorte.Value))
End If

txtTasa = fxTasaRef(txtPlazo, Mid(cboPlazo.Text, 1, 1), txtCodigo, cboOperadora.ItemData(cboOperadora.ListIndex))

If CCur(txtTasa) > 0 Then
   txtIntereses = CCur(txtInversion) * IIf((Mid(cboPlazo.Text, 1, 1) = "D"), CLng(txtPlazo), CLng(txtPlazo) * 30) * CCur(txtTasa) / 36500
   txtIntereses = Format(txtIntereses, "Standard")
End If

vError:

End Sub

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboRenueva.SetFocus
End Sub


Public Sub sbConsultaExterna(pOperadora As Long, pFondos As String, pContrato As Long)

txtCodigo.Text = CStr(pFondos)
txtContrato.Text = CStr(pContrato)
  
Call sbConsultaContrato(txtContrato.Text)

End Sub


Public Sub sbConsultaExternaNuevoCnt(pCedula As String, pNombre As String)

'Realiza la Opcion de Nuevo con Pre-Carga de la Persona
vEdita = False
Call sbLimpiaPantalla
Call sbToolBar(Me.tlb, "edicion")
txtCodigo.SetFocus
      
txtCedula.Text = pCedula
txtNombre.Text = pNombre

End Sub


Private Sub Form_Activate()
 vModulo = 18 'Fondo de Inversion
End Sub

Private Sub Form_Load()
Dim strSQL As String, i As Integer


vModulo = 18 'Fondo de Inversion
vGrid.AppearanceStyle = fxGridStyle

Call sbTaskPanel_Load

For i = 0 To tcMain.ItemCount - 1
    tcMain.Item(i).Enabled = False
Next i


vEdita = True

vPaso = True


'Retiros
With lsw.ColumnHeaders
    .Clear
    .Add , , "No.Ret/Liq", 1500
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2100, vbCenter
    .Add , , "Aportes", 2100, vbRightJustify
    .Add , , "Rendimiento", 2100, vbRightJustify
    .Add , , "Estado", 2100, vbCenter
End With


'With lswDestinos.ColumnHeaders
'    .Clear
'    .Add , , "Destino", 6500
'End With


cboPlazo.Clear
cboPlazo.AddItem "Días"
cboPlazo.AddItem "Meses"
cboPlazo.Text = "Días"

cboIncTipo.Clear
cboIncTipo.AddItem "Porcentaje"
cboIncTipo.AddItem "Monto"
cboIncTipo.Text = "Monto"

cboRenueva.Clear
cboRenueva.AddItem "SI"
cboRenueva.AddItem "NO"
cboRenueva.Text = "SI"

cboTipoDocumento.Clear
cboTipoDocumento.AddItem "Cheque"
cboTipoDocumento.AddItem "Transferencia"
cboTipoDocumento.Text = "Transferencia"

Call sbToolBarIconos(Me.tlb)
Call sbToolBar(Me.tlb, "nuevo")

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'Idx' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

strSQL = "select rtrim(nombre) as 'ItmX',cod_vendedor as 'IdX' from FND_vendedores"
Call sbCbo_Llena_New(cboVendedor, strSQL, False, True)


strSQL = "select id_Banco as 'IdX', rtrim(descripcion) as 'ItmX' from Tes_Bancos where Estado = 'A'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)


vPaso = True
    strSQL = "SELECT ID_FRECUENCIACUPON as 'Idx' , dbo.fxSys_Cadena_Capitaliza ( CUPON ) as 'ItmX'" _
           & " FROM FND_CDP_FRECUENCIACUPONES Where Estado = 1 Order by FRECUENCIA_DIAS asc"
    Call sbCbo_Llena_New(cboCuponFrecuencia, strSQL, False, True)
    
    strSQL = "select ID_PLAZO as 'IdX', dbo.fxSys_Cadena_Capitaliza ( PLAZO ) as 'ItmX'" _
           & " From FND_CDP_PLAZOS  Where Estado = 1  Order by PLAZO_DIAS  asc"
    Call sbCbo_Llena_New(cboPlazoInversion, strSQL, False, True)
vPaso = False
'
'With cboCuponFrecuencia
'  .Clear
'  .AddItem "No Aplica"
'  .AddItem "Mensuales"
'  .AddItem "Trimestrales"
'  .AddItem "Semestrales"
'  .AddItem "Anual"
'  .AddItem "Cuatrimestral"
'  .AddItem "Al Vencimiento"
'End With
'

Call sbLimpiaPantalla


vPaso = False
vGuardar = True

 
'Call Formularios(Me)
'Call RefrescaTags(Me)
 
End Sub


Private Sub sbConsultaContrato(xCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spFnd_Contrato_Consulta " & cboOperadora.ItemData(cboOperadora.ListIndex) & ", '" _
       & Trim(txtCodigo) & "', " & xCodigo & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(Me.tlb, "activo")
  
  tcMain.Item(0).Selected = True
  vEdita = True
  vCarga = True
  
  vCodigo = rs!COD_CONTRATO
  txtContrato.Text = vCodigo
  txtDescripcion.Text = rs!PlanDesc & ""
  
  txtDivisa.Text = Trim(rs!Cod_Moneda & "")
  
  
  mTipoDeduc = rs!Tipo_Deduc
  mPorcRef = rs!Porc_Deduc
  
  txtMonto.Text = Format(rs!Monto, "Standard")
  txtPorcentaje.Text = Format(rs!Porc_Deduc, "Standard")
  
  txtInversion.Text = Format(IIf(IsNull(rs!Inversion), 0, rs!Inversion), "Standard")
  vCambios.vInversion = txtInversion.Text
  
  txtPlazo.Text = CStr(rs!Plazo)
  vCambios.vPlazo = txtPlazo.Text
  
  
    If rs!Tipo_Deduc = "M" Then
     lblPorcentajeDesc.Visible = False
     txtPorcentaje.Visible = False
     
     vCambios.vCuota = txtMonto.Text
     
     If rs!tipo_cdp = 1 Then
        lblMensualidad.Visible = False
        txtMonto.Visible = False
        
        
        txtInversion.Visible = True
        lblInversion.Visible = True
        txtInversion.Locked = False
        
        txtPlazo.Visible = False
        cboPlazo.Visible = False
        
        cboPlazoInversion.Visible = True
        
        gbCupones.Enabled = True
        
        chkCuponPaga.Value = IIf(rs!PAGO_CUPONESCDP = 1, xtpChecked, xtpUnchecked)
        chkCuponPaga.Enabled = True
        
        If IsNull(rs!TASA_PREFERENCIAL_APLICA) Then
         vTasaPreferencial = False
        Else
         vTasaPreferencial = rs!TASA_PREFERENCIAL_APLICA
        End If
        
        btnTasaPreferencial.Visible = True
        
     Else
        lblMensualidad.Visible = True
        txtMonto.Visible = True
     
        txtPlazo.Visible = True
        cboPlazo.Visible = True
        
        txtInversion.Visible = False
        lblInversion.Visible = False
        txtInversion.Locked = True
        
        cboPlazoInversion.Visible = False
        
        gbCupones.Enabled = False
        
        chkCuponPaga.Value = xtpUnchecked
        chkCuponPaga.Enabled = False
     
        btnTasaPreferencial.Visible = False
     
     End If
     
  Else
     vCambios.vCuota = txtPorcentaje.Text
    
     lblPorcentajeDesc.Left = lblMensualidad.Left
     lblPorcentajeDesc.top = lblMensualidad.top
     
     txtPorcentaje.Left = txtMonto.Left
     txtPorcentaje.top = txtMonto.top
          
     lblPorcentajeDesc.Visible = True
     txtPorcentaje.Visible = True
     
     lblMensualidad.Visible = False
     lblInversion.Visible = False
     
     txtMonto.Visible = False
     txtInversion.Visible = False
  
  End If

  
  chkTasaPreferencial.Value = vTasaPreferencial
  
  txtEstado.Text = IIf(rs!Estado = "A", "ACTIVO", "LIQUIDADO")
  txtFecha.Text = Format(rs!Fecha_Inicio, "yyyy/mm/dd")
  
  
  txtCedula.Text = Trim(rs!Cedula)
  txtNombre.Text = rs!Cliente & ""
  
  cboVendedor.Text = rs!Vendedor
  
  txtTasa.Text = Format(rs!Tasa_Referencia & "", "Standard")
  txtTasaTipo.Text = IIf(rs!Tasa_Tipo = "V", "Variante", "Fija")
  txtPtsAdd.Text = Format(rs!Tasa_PtsAdd & "", "Standard")
   
  txtCuponProximo.Text = Format(rs!cupon_proximo & "", "yyyy/mm/dd")
  txtCuponUltimo.Text = Format(rs!cupon_ultimo & "", "yyyy/mm/dd")

'   cboCuponFrecuencia.Text = fxCuponFrecuencia(rs!Cupon_Frecuencia & "")

'  Call sbCboAsignaDato(cboCuponFrecuencia, rs!Frecuencia_Cupon_Desc, True, rs!Frecuencia_Cupon_Id)
  
  
  txtAcumAportes.Text = Format(rs!APORTES, "Standard")
  txtAcumRend.Text = Format(rs!Rendimiento, "Standard")
  txtAcumTotal.Text = Format(rs!APORTES + rs!Rendimiento, "Standard")
  
  txtMontoTransito.Text = Format(IIf(IsNull(rs!Monto_Transito), 0, rs!Monto_Transito), "Standard")
  txtDisponible.Text = Format(CCur(txtAcumTotal.Text) - CCur(txtMontoTransito.Text), "Standard")
  
  txtCuentaCliente.Text = rs!CUENTA_CLIENTE & ""
  
  
  txtOperacion.Text = CStr(rs!Operacion & "")
  
  If IIf(IsNull(rs!plazo_Tipo), "M", rs!plazo_Tipo) = "D" Then
    cboPlazo.Text = "Días"
  Else
    cboPlazo.Text = "Meses"
  End If
  vCambios.vDescPlazo = cboPlazo.Text
  
  dtpCorte.Value = Format(IIf(IsNull(rs!Fecha_Corte), rs!Fecha_Inicio, rs!Fecha_Corte), "dd/mm/yyyy")
  
  
  chkDeducePlanilla.Value = rs!Ind_Deduccion
  vCambios.vDedPlanilla = rs!Ind_Deduccion
  
  chkPagoTercero.Value = rs!PERMITE_GIRO_TERCEROS
  If rs!PlanPermiteGT = 0 Then
     chkPagoTercero.Value = vbUnchecked
     chkPagoTercero.Enabled = False
  Else
     chkPagoTercero.Enabled = True
  End If

  cboRenueva.Text = IIf(rs!Renueva = "S", "SI", "NO")
  cboIncTipo.Text = IIf(rs!inc_tipo = "P", "Porcentaje", "Monto")
  txtIncAnual.Text = Format(rs!Inc_Anual, "Standard")
  
  If Not IsNull(rs!TIPO_PAGO) Then cboTipoDocumento.Text = fxgFNDTipoPago("C", rs!TIPO_PAGO)
  Call sbCboAsignaDato(cboBanco, rs!BancoDesc, False, rs!BancoID)
   
  
    
   'Asigna Cuenta Utilizada
 Call sbCboAsignaDato(cboCuenta, Trim(rs!Cuenta_Ahorros & ""), True, Trim(rs!Cuenta_Ahorros & ""))
  
  txtExc.Text = Format(rs!CapExc, "Standard")
    
  txtAlbaceaCed.Text = rs!albacea_Cedula & ""
  txtAlbaceaNom.Text = rs!albacea_nombre & ""
  
  If rs!cuenta_maestra = 1 Then
     vGrid.Enabled = True
     txtMonto.Locked = True
  Else
     vGrid.Enabled = False
  End If
'  ssTab.TabEnabled(2) = True
'  ssTab.TabEnabled(3) = True
  
  
  'Si tiene aportes y es CDP, Bloquear los plazos y tasas y tipo de plazo
  If rs!tipo_cdp = 1 Then
    vPaso = True
        Call sbCboAsignaDato(cboPlazoInversion, rs!Plazo_Desc, True, rs!Plazo_Id)
        Call sbCboAsignaDato(cboCuponFrecuencia, rs!Frecuencia_Cupon_Desc, True, rs!Frecuencia_Cupon_Id)
        chkCuponPaga.Value = rs!CDP_PAGA_CUPON
    vPaso = False
  End If
  
  cboCuponFrecuencia.Locked = True
  If vTipoCDP And rs!APORTES > 0 Then
     txtInversion.Locked = True
     cboPlazo.Locked = True
     txtPlazo.Locked = True
     dtpCorte.Enabled = False
  End If
  
  If vTipoCDP And rs!APORTES = 0 Then
     txtInversion.Locked = False
     cboPlazo.Locked = False
     txtPlazo.Locked = False
     dtpCorte.Enabled = True
     cboCuponFrecuencia.Locked = False
  End If
  
  'Esto para todos los demas contratos
  If rs!APORTES > 0 Then
     txtCedula.Locked = True
     txtNombre.Locked = True
  Else
     txtCedula.Locked = False
     txtNombre.Locked = False
  End If
 
 ' ssTabAux.TabEnabled(3) = True
 
 
  StatusBarX.Panels(1).Text = rs!Fecha_Inicio & ""
  StatusBarX.Panels(2).Text = rs!Usuario & ""
  StatusBarX.Panels(3).Text = rs!MODIFICA_FECHA & ""
  StatusBarX.Panels(4).Text = rs!MODIFICA_USUARIO & ""
    

  If Not vGrid.Enabled Then
        'Activar
  End If
  
Else
   Me.MousePointer = vbDefault
   Call sbLimpiaPantalla
   MsgBox "El Contrato No existe o Este Usuario no tiene acceso a su consulta...", vbInformation
End If
rs.Close

Me.MousePointer = vbDefault

vCarga = False

Call RefrescaTags(Me)

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 Resume

End Sub


Private Sub sbConsultaPlan(pCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select DESCRIPCION, TIPO_DEDUC, PORC_DEDUC, TIPO_CDP, PAGO_CUPONES" _
       & " from fnd_Planes" _
       & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " and cod_plan='" & Trim(pCodigo) & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  
  txtDescripcion.Text = rs!Descripcion
  
  mTipoDeduc = rs!Tipo_Deduc
  mPorcRef = rs!Porc_Deduc



  If rs!Tipo_Deduc = "M" Then
     lblPorcentajeDesc.Visible = False
     txtPorcentaje.Visible = False
     
     If rs!tipo_cdp = 1 Then
        lblMensualidad.Visible = False
        txtMonto.Visible = False
        
        
        txtInversion.Visible = True
        lblInversion.Visible = True
        txtInversion.Locked = False
        
        txtPlazo.Visible = False
        cboPlazo.Visible = False
        
        cboPlazoInversion.Visible = True
        
        gbCupones.Enabled = True
        
        If rs!PAGO_CUPONES = True Then
            chkCuponPaga.Value = xtpUnchecked
            chkCuponPaga.Enabled = True
        Else
            chkCuponPaga.Value = xtpUnchecked
            chkCuponPaga.Enabled = False
        End If
        
        
     Else
        lblMensualidad.Visible = True
        txtMonto.Visible = True
     
        txtPlazo.Visible = True
        cboPlazo.Visible = True
        
        txtInversion.Visible = False
        lblInversion.Visible = False
        txtInversion.Locked = True
        
        cboPlazoInversion.Visible = False
        
        gbCupones.Enabled = False
     
     End If
     
  Else
    
     lblPorcentajeDesc.Left = lblMensualidad.Left
     lblPorcentajeDesc.top = lblMensualidad.top
     
     txtPorcentaje.Left = txtMonto.Left
     txtPorcentaje.top = txtMonto.top
          
     lblPorcentajeDesc.Visible = True
     txtPorcentaje.Visible = True
     
     lblMensualidad.Visible = False
     lblInversion.Visible = False
     
     txtMonto.Visible = False
     txtInversion.Visible = False
  
  End If
  
  
End If
rs.Close

If cboPlazoInversion.Visible Then

 vPaso = True
    strSQL = "exec spFnd_Inversion_Plazos '" & txtCodigo.Text & "'"
    Call sbCbo_Llena_New(cboPlazoInversion, strSQL, False, True)
 vPaso = False
    
    Call cboPlazoInversion_Click
End If


Me.MousePointer = vbDefault


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbConsulta()
       
gBusquedas.Convertir = "S"
gBusquedas.Columna = "CEDULA"
gBusquedas.Orden = "CEDULA"

gBusquedas.Consulta = "select COD_PLAN,COD_CONTRATO,CEDULA,FECHA_INICIO from fnd_Contratos"
gBusquedas.Filtro = " and cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
                  & " and cod_plan='" & Trim(txtCodigo) & "'"

frmBusquedas.Show vbModal

txtContrato.SetFocus

If Trim(gBusquedas.Resultado2) <> "" Then
   txtContrato = Trim(gBusquedas.Resultado2)
End If

If txtCedula.Enabled Then txtCedula.SetFocus
gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""

End Sub


Private Sub sbGuardar()
Dim lngContrato As Long, vFecha As Date
Dim strSQL As String, vCuponFrecuencia As String
Dim vCuponProximo As String

If Not fxVerificaDatos Then
  Exit Sub
End If

On Error GoTo vError

vFecha = fxFechaServidor

Dim pCuponFrecuencia As String, pPlazoInversionId As String, pCuponPaga As String, pCuponFrecuenciaId As String

If txtInversion.Visible Then
    pPlazoInversionId = cboPlazoInversion.ItemData(cboPlazoInversion.ListIndex)
        
    If chkCuponPaga.Value Then
      pCuponPaga = "1"
      pCuponFrecuencia = Mid(cboCuponFrecuencia.Text, 1, 1)
      pCuponFrecuenciaId = cboCuponFrecuencia.ItemData(cboCuponFrecuencia.ListIndex)
    Else
      pCuponPaga = "0"
      pCuponFrecuencia = "N"
      pCuponFrecuenciaId = "Null"
    End If

Else
    pCuponFrecuencia = "N"
    pCuponFrecuenciaId = "Null"
    pPlazoInversionId = "Null"
    pCuponPaga = "0"
End If


If vEdita Then
  strSQL = "Update FND_Contratos Set cod_Vendedor = '" & cboVendedor.ItemData(cboVendedor.ListIndex) _
         & "',Plazo = " & txtPlazo & ",fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
         & "',Monto = " & CCur(txtMonto) & ",Renueva = '" & IIf(Trim(cboRenueva) = "SI", "S", "N") & "',Inc_Anual = " & CCur(txtIncAnual) _
         & ",Inc_Tipo='" & IIf(Trim(cboIncTipo) = "Porcentaje", "P", "M") & "',Cod_Banco = " & cboBanco.ItemData(cboBanco.ListIndex) _
         & ",Cuenta_Ahorros = '" & cboCuenta.ItemData(cboCuenta.ListIndex) & "', tipo_Pago = '" & fxgFNDTipoPago("D", cboTipoDocumento.Text) _
         & "',CapExc = " & Trim(txtExc) & ", albacea_Cedula = '" & txtAlbaceaCed & "',albacea_nombre = '" & txtAlbaceaNom _
         & "', plazo_tipo = '" & Mid(cboPlazo.Text, 1, 1) & "', inversion = " & CCur(txtInversion) & ", tasa_referencia = " & CCur(txtTasa) _
         & ",modifica_fecha = dbo.MyGetdate(), modifica_usuario = '" & glogon.Usuario & "'" _
         & ",cupon_frecuencia = '" & pCuponFrecuencia & "',cupon_proximo = " _
         & IIf(Trim(txtCuponProximo.Text) = "", "Null", "'" & txtCuponProximo.Text & "'") _
         & ",ind_deduccion = " & chkDeducePlanilla.Value & ", PERMITE_GIRO_TERCEROS = " & chkPagoTercero.Value _
         & ", Tipo_Deduc = '" & mTipoDeduc & "', PORC_DEDUC = " & CCur(txtPorcentaje.Text) _
         & ", IDCUPON_FRECUENCIA = " & pCuponFrecuenciaId _
         & ", PAGO_CUPONESCDP = " & pCuponPaga _
         & ", ID_PER_TASA = dbo.fxFnd_ReglaId_Tasa(cod_plan, Fecha_Inicio)" _
         & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
         & " and cod_plan = '" & Trim(txtCodigo) & "' and cod_Contrato = " & vCodigo
  Call ConectionExecute(strSQL)
  
  
  Call Bitacora("Modifica", "Contrato:" & vCodigo & " Plan:" & Trim(txtCodigo) & " Oper:" & cboOperadora.ItemData(cboOperadora.ListIndex))
   
  If vCambios.vCuota <> txtMonto Then
     Call sbGuardaCambios("01", "Anterior " & Format(vCambios.vCuota, "Standard") & " Nueva " & txtMonto.Text)
  End If

   If vCambios.vPlazo <> txtPlazo Then
     Call sbGuardaCambios("02", "Anterior  " & vCambios.vPlazo & " " & vCambios.vDescPlazo & " Nuevo " & txtPlazo.Text & " " & cboPlazo)
  End If
  
  If vCambios.vInversion <> txtInversion Then
     Call sbGuardaCambios("03", "Anterior " & Format(vCambios.vInversion, "Standard") & " Nueva " & txtInversion.Text)
  End If
 
  If vCambios.vDedPlanilla <> chkDeducePlanilla.Value Then
     Call sbGuardaCambios("06", "Anterior " & IIf((vCambios.vDedPlanilla = 1), "Sí", "No") & " Nuevo " & IIf((chkDeducePlanilla.Value = 1), "Sí", "No"))
  End If
 
Else

   lngContrato = fxConsecutivoContrato
   
   strSQL = "insert FND_Contratos(Cod_operadora,Cod_plan,Cod_Contrato,Cedula,Cod_Vendedor, Tipo_Deduc, PORC_DEDUC" _
          & ", Estado, Fecha_Inicio, Plazo, Monto, Renueva, Inc_Anual, Inc_Tipo, Ind_comision" _
          & ", Cod_Banco, Cuenta_Ahorros, Tipo_Pago, CapExc, rend_corte, rend_saldo, fecha_corte, usuario" _
          & ", albacea_cedula, albacea_nombre, plazo_tipo, inversion, tasa_referencia, Tasa_Tipo, Tasa_PtsAdd" _
          & ", Cupon_Frecuencia, Cupon_Proximo, Cupon_Consec, ind_deduccion, PERMITE_GIRO_TERCEROS" _
          & ", IDCUPON_FRECUENCIA, PAGO_CUPONESCDP, ID_PER_TASA)" _
          & "  values(" & cboOperadora.ItemData(cboOperadora.ListIndex) _
          & ", '" & Trim(txtCodigo) & "', " & lngContrato & ", '" & Trim(txtCedula) & "', '" & cboVendedor.ItemData(cboVendedor.ListIndex) _
          & "', '" & mTipoDeduc & "', " & CCur(txtPorcentaje.Text) & ", 'A', '" & Format(vFecha, "yyyy/mm/dd") & "'," & txtPlazo & "," & CCur(txtMonto) _
          & ", '" & IIf(Trim(cboRenueva) = "SI", "S", "N") & "'," & CCur(txtIncAnual) & ",'" & IIf(Trim(cboIncTipo) = "Porcentaje", "P", "M") _
          & "', 0, " & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & cboCuenta.ItemData(cboCuenta.ListIndex) & "','" & fxgFNDTipoPago("D", cboTipoDocumento.Text) _
          & "', " & Trim(txtExc) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',0,'" & Format(dtpCorte.Value, "yyyy/mm/dd") _
          & "', '" & glogon.Usuario & "','" & txtAlbaceaCed & "','" & txtAlbaceaNom & "','" & Mid(cboPlazo.Text, 1, 1) _
          & "', " & CCur(txtInversion) & "," & CCur(txtTasa) & ",'" & Mid(txtTasaTipo.Text, 1, 1) & "'," & CCur(txtPtsAdd.Text) _
          & ", '" & pCuponFrecuencia & "'," & IIf(Trim(txtCuponProximo.Text) = "", "Null", "'" & txtCuponProximo.Text & "'") _
          & ", 0," & chkDeducePlanilla.Value & "," & chkPagoTercero.Value _
          & ", " & pCuponFrecuenciaId & ", " & pCuponPaga _
          & ", dbo.fxFnd_ReglaId_Tasa('" & Trim(txtCodigo) & "', '" & Format(vFecha, "yyyy/mm/dd") & "')   )"
   Call ConectionExecute(strSQL)
   
   
   txtContrato = lngContrato
   
     If fxAplicaBeneficiarios Then
      If Not fxBeneficiariosNoIncluidos(lngContrato) Then
         MsgBox "No estan incluidos los beneficiarios o el porcentaje es inferior al 100%..." & _
         vbCrLf & "Por Favor Incluirlos", vbInformation
'         ssTab.Tab = 3
      End If
   End If
    
   Call Bitacora("Registra", "Contrato:" & lngContrato & " Plan:" & Trim(txtCodigo) & " Oper:" & cboOperadora.ItemData(cboOperadora.ListIndex))
   Call sbSIFRegistraTags(txtCodigo.Text, "S09", "Fondos", txtContrato.Text, "FND", txtCodigo.Text, txtContrato.Text, txtCedula.Text)
   Call sbGuardaCambios("05", "Mensualidad: " & txtMonto.Text & " ¦ Inversión: " & txtInversion.Text)
      
'   ssTabAux.TabEnabled(3) = True
    
End If

vCodigo = txtContrato

If vTipoCDP Then
  strSQL = "exec spFndCDPCupones " & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & Trim(txtCodigo.Text) & "'," & vCodigo & ",'" & glogon.Usuario & "'"
  Call ConectionExecute(strSQL)
End If

vEdita = True
Call sbToolBar(Me.tlb, "activo")

txtContrato.SetFocus

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer
Dim strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este contrato", vbYesNo)

If i = vbYes Then
  
  strSQL = "delete FND_contratos where cod_contrato = " & vCodigo _
         & " and cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
         & " and cod_plan='" & Trim(txtCodigo) & "'"
  
  Call ConectionExecute(strSQL)
  Call Bitacora("Borra", "Contrato:" & vCodigo & " Plan:" & Trim(txtCodigo) & " Oper:" & cboOperadora.ItemData(cboOperadora.ListIndex))
  
  Call sbLimpiaPantalla
  Call sbToolBar(Me.tlb, "nuevo")
  
  txtContrato.SetFocus
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub





Public Sub sbNuevaSubCuenta()

vGrid.MaxRows = vGrid.MaxRows + 1
vGrid.Row = vGrid.MaxRows
vGrid.Col = 2
vGrid.Text = Trim(txtCedula) & "-" & Format(vGrid.MaxRows, "00")
vGrid.Col = 4
vGrid.Text = "0"
vGrid.Col = 5
vGrid.Text = "0"


End Sub

Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 0
vGrid.MaxRows = 1

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End Sub





Private Function fxDestino_Guardar() As Long

On Error GoTo vError

fxDestino_Guardar = 0
gDestinos.Row = gDestinos.ActiveRow
gDestinos.Col = 1

If gDestinos.Text = "" Then 'Insertar
  
  gDestinos.Col = 2
  strSQL = "insert into FND_CONTRATOS_DESTINOS_AHORRO(ID_DESTINO, COD_PLAN, COD_CONTRATO, OBSERVACIONES, FEC_REGISTRO, USU_REGISTRO) values(" _
         & gDestinos.Text & ", '" & Trim(txtCodigo.Text) & "', " & txtContrato.Text & ", '"
  gDestinos.Col = 4
  strSQL = strSQL & gDestinos.Text & "', dbo.MyGetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)
  
  strSQL = "select isnull(max(ID_REGISTRO),0) as IdSeQ from FND_CONTRATOS_DESTINOS_AHORRO"
  Call OpenRecordSet(rs, strSQL)
  
  gDestinos.Col = 1
  gDestinos.Text = CStr(rs!IdSeQ)
        
    rs.Close
  
'
'  Call Bitacora("Registra", "Categoría para Apremiantes Id: " & gDestinos.Text)

Else 'Actualizar

 strSQL = "update FND_CONTRATOS_DESTINOS_AHORRO set OBSERVACIONES = '"
 gDestinos.Col = 4
 strSQL = strSQL & gDestinos.Text & ", Modifica_Fecha = dbo.MyGetdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "' where ID_REGISTRO = "
 gDestinos.Col = 1
 strSQL = strSQL & gDestinos.Text
 
 Call ConectionExecute(strSQL)

' gDestinos.Col = 1
' Call Bitacora("Modifica", "Categoría para Apremiantes Id: " & gDestinos.Text)

End If

fxDestino_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub gDestinos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If gDestinos.ActiveCol = gDestinos.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxDestino_Guardar
  If i = 0 Then Exit Sub
  gDestinos.Row = gDestinos.ActiveRow
End If

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lsw.ListItems.Count = 0 Then Exit Sub

lblRetLiq.Caption = "Retiro o Liquidación : " & Item
lblRetLiq.Tag = Item

If Mid(Item.SubItems(5), 1, 1) = "P" Then
   cmdReversion.Enabled = True
Else
   cmdReversion.Enabled = False
End If

Call RefrescaTags(Me)
End Sub


Private Sub lswDestinos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert fnd_contratos_destinos(cod_plan,cod_operadora,cod_contrato,cod_destino,registro_usuario,registro_fecha)" _
          & " values('" & txtCodigo.Text & "'," & cboOperadora.ItemData(cboOperadora.ListIndex) & "," & txtContrato.Text & ",'" & Item.Tag _
          & "','" & glogon.Usuario & "',dbo.MyGetdate())"
   Call ConectionExecute(strSQL)
   Call Bitacora("Aplica", "Asignación Destino: " & Item.Tag & " P.: " & txtCodigo.Text & " Cnt: " & txtContrato.Text)

Else
   strSQL = "delete fnd_contratos_destinos where cod_destino = '" & Item.Tag _
          & "' and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) & " and cod_plan = '" & txtCodigo.Text _
          & "' and cod_contrato = " & txtContrato.Text
   Call ConectionExecute(strSQL)
   Call Bitacora("Elimina", "Asignación Destino: " & Item.Tag & " P.: " & txtCodigo.Text & " Cnt: " & txtContrato.Text)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbConsulta_Detalle(pDetalle As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, i As Integer, curCuota As Currency

On Error GoTo vError


If Not IsNumeric(txtContrato.Text) Then Exit Sub


Select Case pDetalle
  Case "General"
    
    If vGrid.Enabled Then
       curCuota = 0
       For i = 1 To vGrid.MaxRows
          vGrid.Row = i
          vGrid.Col = 1
          If vGrid.Text <> "" Then
             vGrid.Col = 4
             curCuota = curCuota + CCur(vGrid.Text)
          End If
       Next i
           
       txtMonto = Format(curCuota, "Standard")
    
    End If
    
    tcMain.Item(0).Selected = True
  
  Case "Sub Cuentas"
     tcMain.Item(4).Selected = True
     
     ',aportes+rendimiento as Acumulado,parentesco
     strSQL = "select idx,cedula,nombre,cuota,0" _
            & " from fnd_subCuentas where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
            & " and cod_plan = '" & txtCodigo & "' and cod_contrato = " & txtContrato
     Call sbCargaGridLocal(vGrid, 5, strSQL)
     vGrid.MaxRows = vGrid.MaxRows - 1
     Call sbNuevaSubCuenta
  
  Case "Retiros" 'Retiros
     tcMain.Item(5).Selected = True
    
    strSQL = "select consec,fecha,aportes_liq,rendi_liq,estado ,usuario" _
           & " from fnd_liquidacion where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " and cod_plan = '" & txtCodigo & "' and cod_contrato = " & txtContrato
    Call OpenRecordSet(rs, strSQL)
    
    lsw.ListItems.Clear
    
    Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!consec)
         itmX.SubItems(1) = Format(rs!fecha, "dd/mm/yyyy")
         itmX.SubItems(2) = rs!Usuario & ""
         itmX.SubItems(3) = Format(rs!Aportes_Liq, "Standard")
         itmX.SubItems(4) = Format(rs!Rendi_Liq, "Standard")
         If rs!Estado = "P" Then
             itmX.SubItems(5) = "Procesado"
         Else
             itmX.SubItems(5) = "Reversado"
         End If
     rs.MoveNext
    Loop
    rs.Close
    

Case "Beneficiarios" ' Beneficiarios
     
    tcMain.Item(3).Selected = True
    
    lswBeneficiarios.ListItems.Clear
    lswBeneficiarios.ColumnHeaders.Clear
    
    lswBeneficiarios.ColumnHeaders.Add 1, , "Cedula", 1400
    lswBeneficiarios.ColumnHeaders.Add 2, , "Nombre", 3300
    lswBeneficiarios.ColumnHeaders.Add 3, , "Porcentaje", 1200
    lswBeneficiarios.ColumnHeaders.Add 4, , "Parentesco", 1700
    
    strSQL = "Select CedulaBn,Nombre,Porcentaje,parentesco From FND_CONTRATOS_BENEFICIARIOS where " _
           & " Cedula = '" & Trim(txtCedula) & "' and cod_contrato = " & txtContrato & "" _
           & " and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " and cod_plan='" & Trim(txtCodigo) & "'"
           
           
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswBeneficiarios.ListItems.Add(, , rs!cedulaBn)
           itmX.SubItems(1) = Trim(rs!Nombre)
           itmX.SubItems(2) = Trim(rs!Porcentaje) & "%"
           itmX.SubItems(3) = fxParentesco(rs!parentesco)
       rs.MoveNext
    Loop
    rs.Close
  

  Case "Cupones" 'Cupones
    tcMain.Item(6).Selected = True
    
    Call sbFnd_Contratos_Cupones(cboOperadora.ItemData(cboOperadora.ListIndex), txtCodigo.Text, txtContrato.Text, lswCupones)
  
  
  Case "Bitacora" 'Bitacora
    tcMain.Item(7).Selected = True
    
    Call sbFnd_Contratos_Bitacora(cboOperadora.ItemData(cboOperadora.ListIndex), txtCodigo.Text, txtContrato.Text, lswBitacora)

  Case "Destinos" 'Destinos
    Call sbDestinos_Load
End Select

vError:
End Sub


Private Sub sbDestinos_Load()
     tcMain.Item(2).Selected = True
     
     
'     strSQL = "select D.cod_destino,D.descripcion,A.cod_contrato " _
'            & " from fnd_destinos D left join fnd_contratos_destinos A on D.cod_destino = A.cod_destino" _
'            & " and A.cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
'            & " and A.cod_plan = '" & Trim(txtCodigo.Text) & "' and A.cod_contrato = " & txtContrato.Text _
'            & " Where D.cod_destino in(select cod_destino from fnd_planes_destinos where cod_plan = '" _
'            & txtCodigo.Text & "')"
'     lswDestinos.ListItems.Clear
'     Call OpenRecordSet(rs, strSQL)
'     Do While Not rs.EOF
'      Set itmX = lswDestinos.ListItems.Add(, , rs!Descripcion)
'          itmX.Tag = rs!cod_destino
'
'      If IsNull(rs!COD_CONTRATO) Then
'         itmX.Checked = False
'      Else
'         itmX.ForeColor = vbBlue
'         itmX.Checked = True
'      End If
'      rs.MoveNext
'     Loop
'     rs.Close


Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

With gDestinos

.MaxRows = 0

strSQL = "exec spFnd_Contrato_Destinos_List " & cboOperadora.ItemData(cboOperadora.ListIndex) & ", '" & Trim(txtCodigo.Text) & "', " & txtContrato.Text
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 
 .MaxRows = .MaxRows + 1
 .Row = .MaxRows
 
 .Col = 1
 If rs!Id_Registro > 0 Then
     .Text = rs!Id_Registro
 End If
 
 .Col = 2
 .Text = rs!Id_Destino
 .Col = 3
 .Text = rs!Descripcion
 .Col = 4
 .Text = rs!observaciones
 
 rs.MoveNext
Loop
rs.Close

End With

Exit Sub

vError:


End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      Call sbToolBar(Me.tlb, "edicion")
      txtCodigo.SetFocus
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
      Call sbToolBar(Me.tlb, "edicion")
      
      cboOperadora.Enabled = False
      txtCodigo.Locked = True
      tcMain.Item(0).Selected = True
      cboVendedor.SetFocus
    
    Case "BORRAR"
      Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
'      txtCodigo_LostFocus
'      txtCedula_LostFocus
      If vGuardar Then Call sbGuardar
      
    Case "DESHACER"
      Call sbToolBar(Me.tlb, "nuevo")
      Call sbLimpiaPantalla
      
      txtContrato.SetFocus
      vEdita = True
      
    Case "CONSULTAR"
      Select Case vBusqueda
        Case "1"
         txtCodigo_KeyDown vbKeyF4, 0
        Case "2"
         txtDescripcion_KeyDown vbKeyF4, 0
        Case "3"
         txtContrato_KeyDown vbKeyF4, 0
        Case "4"
         txtCedula_KeyDown vbKeyF4, 0
        Case "5"
         txtNombre_KeyDown vbKeyF4, 0
      End Select
       
    
    Case "REPORTES"
        If IsNumeric(txtContrato.Text) Then
            gFondos.Operadora = cboOperadora.ItemData(cboOperadora.ListIndex)
            gFondos.Plan = txtCodigo.Text
            gFondos.Contrato = txtContrato.Text
            Call sbFormsCall("frmFNDContratosInformes", 1, , , False, Me)
        End If
        
        
    Case "CERRAR"
      Unload Me
      

End Select

Call RefrescaTags(Me)

End Sub



Private Sub tpMain_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Call sbTaskPanel_Accion(Item.Id)
End Sub

Private Sub txtAlbaceaCed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAlbaceaNom.SetFocus
End Sub

Private Sub txtCedula_GotFocus()
vBusqueda = "4"
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cedula"
   gBusquedas.Orden = "cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "select cedula,nombre from socios"
   frmBusquedas.Show vbModal
   txtNombre.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCedula = Trim(gBusquedas.Resultado)
      txtNombre = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

End Sub

Private Sub txtCedula_LostFocus()
 txtNombre = fxNombre(txtCedula)
End Sub

Private Sub txtCodigo_Change()
If Trim(txtContrato) <> "" Then Call sbConsultaContrato(Trim(txtContrato))
End Sub

Private Sub txtCodigo_GotFocus()
vBusqueda = "1"
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 And txtCodigo.Locked = False Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   gBusquedas.Filtro = " And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
                     & " and dbo.fxFnd_Seguridad_Acceso_Planes('" & glogon.Usuario & "', cod_operadora, cod_plan) = 1"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
      Call sbConsultaPlan(txtCodigo.Text)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    Call sbConsultaPlan(txtCodigo.Text)
    txtDescripcion.SetFocus
End If

End Sub


Private Function fxTasaPtsAdd(xPlazo As Long, xTipo As String, xPlan As String _
                         , xOperadora As Integer) As Currency
Dim strSQL As String, rs As New ADODB.Recordset
Dim xTasa As Currency

xTasa = 0
If xTipo = "M" Then
   xPlazo = xPlazo * 30
End If

strSQL = "select tasa_base,UTILIZA_TBP,TIPO_CDP" _
       & ",dbo.fxFNDTasaPlus(cod_operadora,cod_plan," & xPlazo & ") as PlusTasa" _
       & " from fnd_planes where cod_operadora = " & xOperadora & " and cod_plan = '" & xPlan & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
   xTasa = rs!PlusTasa
End If
rs.Close

End Function

Private Function fxTasaRef(xPlazo As Long, xTipo As String, xPlan As String _
                         , xOperadora As Integer) As Currency
Dim strSQL As String, rs As New ADODB.Recordset
Dim xTasa As Currency

On Error GoTo vError

If chkCuponPaga.Value = xtpUnchecked Or Not txtInversion.Visible Then
    strSQL = "select dbo.fxFNDCalcularTasaRefContrato(" & xOperadora & ", '" & xPlan & "', " & txtPlazo.Text & ", '" & xTipo & "', Null, Null, 0) as 'Tasa'"
Else
    If cboCuponFrecuencia.ListIndex = -1 Then
      If IsNumeric(txtTasa.Text) Then
          fxTasaRef = CCur(txtTasa.Text)
      Else
          fxTasaRef = 0
      End If
      Exit Function
    End If
    
    strSQL = "exec dbo.spFnd_Inversion_Tasas_Condiciones " & xOperadora & ", '" & xPlan & "', " & cboPlazoInversion.ItemData(cboPlazoInversion.ListIndex) & ", " & cboCuponFrecuencia.ItemData(cboCuponFrecuencia.ListIndex)
'     spFnd_Inversion_Tasas_Condiciones(@Operadora int, @Plan varchar(10),  @PlazoId int, @fCuponId int)
End If

Call OpenRecordSet(rs, strSQL)
    xTasa = rs!Tasa
rs.Close

'xTasa = 0
'If xTipo = "M" Then
'   xPlazo = xPlazo * 30
'End If
'
'strSQL = "select tasa_base,UTILIZA_TBP,TIPO_CDP" _
'       & ",dbo.fxFNDTasaPlus(cod_operadora,cod_plan," & xPlazo & ") as PlusTasa" _
'       & ",dbo.fxFndTasaReferencia( dbo.fxFndTipoTasa(cod_operadora,cod_plan," & xPlazo & ") ) as TasaRef" _
'       & " from fnd_planes where cod_operadora = " & xOperadora & " and cod_plan = '" & xPlan & "'"
'
'Call OpenRecordSet(rs, strSQL)
'
'If Not rs.EOF And Not rs.BOF Then
'
'      If rs!utiliza_tbp = 1 Then
'         xTasa = rs!TasaRef + rs!PlusTasa
'      Else
'         xTasa = IIf(IsNull(rs!tasa_base), 0, rs!tasa_base) + rs!PlusTasa
'      End If
'
'
''   If rs!tipo_cdp = 1 Then
''      If rs!utiliza_tbp = 1 Then
''         xTasa = rs!TasaRef + rs!PlusTasa
''      Else
''         xTasa = IIf(IsNull(rs!tasa_base), 0, rs!tasa_base) + rs!PlusTasa
''      End If
''   End If
'End If
'rs.Close

'Si está en modo de carga de información conservar la tasa de registro
If vCarga Then
    fxTasaRef = CCur(txtTasa.Text)
Else
    fxTasaRef = xTasa
End If

Exit Function

vError:
    fxTasaRef = CCur(txtTasa.Text)

End Function



Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

txtCodigo.Text = UCase(txtCodigo.Text)


If Trim(txtCodigo) <> "" Then
   strSQL = "Select Descripcion, Monto_Minimo, Plazo_Minimo, cuenta_maestra, Inversion_minimo" _
          & ", Tipo_CDP, WEB_VENCE, isnull(PERMITE_GIRO_TERCEROS,0) as 'PlanPermiteGT', cod_moneda" _
          & ", PAGO_CUPONES, TASA_MARGEN_NEGOCIACION, dbo.MyGetDate() as 'FechaServidor'" _
          & " from fnd_planes" _
          & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
          & " And cod_plan = '" & Trim(txtCodigo) & "'"
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
       txtDescripcion = Trim(rs!Descripcion)
       vGuardar = True
       
       vMontoMin = rs!MONTO_MINIMO
       vPlazoMin = rs!Plazo_Minimo
       vInversionMin = rs!INVERSION_MINIMO
       vTipoCDP = IIf((rs!tipo_cdp = 1), True, False)
       
       vCDPCuponesAplica = rs!PAGO_CUPONES 'IIf((rs!PAGO_CUPONES = 1), True, False)
       vTasaMargenNegociacion = rs!TASA_MARGEN_NEGOCIACION
       
       chkCuponPaga.Value = xtpUnchecked
       chkCuponPaga.Enabled = False

       
       If vTipoCDP Then
          cboCuponFrecuencia.Locked = False
          txtTasaTipo.Text = "Fija"
       
          chkCuponPaga.Enabled = vCDPCuponesAplica
       Else
          txtTasaTipo.Text = "Variable"
       End If
       
       txtDivisa.Text = Trim(rs!Cod_Moneda & "")
       
       If rs!cuenta_maestra = 1 Then
          txtMonto.Locked = True
          vGrid.Enabled = True
       Else
          txtMonto.Locked = False
          vGrid.Enabled = False
       End If
       
       
        If rs!PlanPermiteGT = 0 Then
           chkPagoTercero.Enabled = False
        Else
           chkPagoTercero.Enabled = True
        End If
       
       
       If vEdita = False Then
        If Not IsNull(rs!web_vence) Then
            txtPlazo.Text = DateDiff("m", rs!FechaServidor, rs!web_vence)
            cboPlazo.Text = "Meses"
        End If
       End If

        Call sbConsultaPlan(txtCodigo.Text)
    Else
       MsgBox "Código de Plan incorrecto", vbExclamation
       txtCodigo = ""
       txtDescripcion = ""
       txtCodigo.SetFocus
       vGuardar = False
       vMontoMin = 0
       vPlazoMin = 0
       vInversionMin = 0
       vTasaMargenNegociacion = 0
       vCDPCuponesAplica = False
       
    End If
    rs.Close
Else
  txtDescripcion = ""
End If

End Sub

Private Sub txtContrato_GotFocus()
vBusqueda = "3"
End Sub

Private Sub txtContrato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbConsulta
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   Call sbTaskPanel_Accion(Id_TaskItem_General)
   If txtCedula.Enabled Then
      txtCedula.SetFocus
   Else
      cboVendedor.SetFocus
   End If
End If
End Sub

Private Sub txtContrato_LostFocus()

If IsNumeric(txtContrato) Then
   Call sbConsultaContrato(Trim(txtContrato))
Else
  If vEdita Then
   Call sbToolBar(Me.tlb, "nuevo")
   Call RefrescaTags(Me)
   Call sbLimpiaPantalla
  End If
End If

End Sub



Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAlbaceaCed.SetFocus
End Sub

Private Sub txtDescripcion_GotFocus()
vBusqueda = "2"
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = " And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContrato.SetFocus

End Sub


Private Sub txtExc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboBanco.SetFocus
End Sub

Private Sub txtExc_LostFocus()
On Error GoTo vError

If IsNumeric(txtExc) Then
   If CCur(txtExc) > 100 Then
      MsgBox "Porcentaje debe ser menor o igual a 100"
      txtExc.SetFocus
   End If
End If

vError:
End Sub



Private Sub txtIncAnual_GotFocus()
On Error GoTo vError
  txtIncAnual = CCur(txtIncAnual)
vError:
End Sub

Private Sub txtIncAnual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtExc.SetFocus
End Sub

Private Sub txtIncAnual_LostFocus()
On Error GoTo vError
  txtIncAnual = Format(CCur(txtIncAnual), "###.00")
vError:
End Sub

Private Sub txtInversion_GotFocus()
On Error GoTo vError
  txtInversion = CCur(txtInversion)
vError:
End Sub

Private Sub txtInversion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  cboPlazoInversion.SetFocus
End If
End Sub

Private Sub txtInversion_LostFocus()
On Error GoTo vError

If IsNumeric(txtInversion) Then
   If CCur(txtInversion) < vInversionMin Then
      MsgBox "La Inversión mínima para este plan debe ser mayor o igual a " & Format(vInversionMin, "Standard")
      txtInversion.SetFocus
   Else
      txtInversion = Format(CCur(txtInversion), "Standard")
   End If
End If
vError:
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
  txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    If txtInversion.Visible Then
     If txtInversion.Enabled Then
        txtInversion.SetFocus
     Else
        cboIncTipo.SetFocus
     End If
    Else
       txtPlazo.SetFocus
    End If
End If
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

If IsNumeric(txtMonto) Then
   If CCur(txtMonto) < vMontoMin Then
      MsgBox "La cuota Mínima para este plan debe ser mayor o igual a " & Format(vMontoMin, "Standard")
'      txtMonto.SetFocus
   Else
      txtMonto = Format(CCur(txtMonto), "Standard")
   End If
End If
vError:

End Sub


Private Sub txtNombre_GotFocus()
vBusqueda = "5"
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "nombre"
   gBusquedas.Orden = "nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "select cedula,nombre from socios"
   frmBusquedas.Show vbModal
   txtNombre.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCedula = Trim(gBusquedas.Resultado)
      txtNombre = Trim(gBusquedas.Resultado2)
   End If
   
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboVendedor.SetFocus

End Sub



Private Sub txtPlazo_GotFocus()
On Error GoTo vError

If CCur(txtTasa) > 0 Then
   txtIntereses = CCur(txtInversion) * IIf((Mid(cboPlazo.Text, 1, 1) = "D"), CLng(txtPlazo), CLng(txtPlazo) * 30) * CCur(txtTasa) / 36500
   txtIntereses = Format(txtIntereses, "Standard")
End If

vError:
End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub txtPlazo_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If Mid(cboPlazo.Text, 1, 1) = "D" Then
    dtpCorte.Value = DateAdd("d", CDbl(txtPlazo), CDate(txtFecha))
Else
    dtpCorte.Value = DateAdd("m", CDbl(txtPlazo), CDate(txtFecha))
End If

txtTasa.Text = Format(fxTasaRef(txtPlazo.Text, Mid(cboPlazo.Text, 1, 1), txtCodigo, cboOperadora.ItemData(cboOperadora.ListIndex)), "##0.00")

If CCur(txtTasa) > 0 Then
   txtIntereses.Text = CCur(txtInversion.Text) * IIf((Mid(cboPlazo.Text, 1, 1) = "D"), CLng(txtPlazo), CLng(txtPlazo) * 30) * CCur(txtTasa.Text) / 36500
   txtIntereses.Text = Format(txtIntereses.Text, "Standard")
End If

vError:

End Sub


Private Function fxSubCuentaContrato()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) + 1 as Ultimo from fnd_subCuentas" _
       & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " and cod_plan = '" & txtCodigo & "' and cod_contrato = " & txtContrato
Call OpenRecordSet(rs, strSQL)
 fxSubCuentaContrato = rs!ultimo
rs.Close


End Function

Private Sub sbActualizaCuotaContrato()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(sum(cuota),0) as Cuota from fnd_subCuentas where cod_operadora = " _
       & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " and cod_plan = '" & txtCodigo & "' and cod_contrato = " & txtContrato
Call OpenRecordSet(rs, strSQL)

strSQL = "update fnd_contratos set monto = " & rs!Cuota _
       & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " and cod_plan = '" & txtCodigo & "' and cod_contrato = " & txtContrato
Call ConectionExecute(strSQL)

rs.Close

End Sub

Private Function fxGuardar() As Integer
Dim strSQL As String
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If vGrid.Text = "" Then

   vGrid.Text = CStr(fxSubCuentaContrato)
   strSQL = "insert fnd_subCuentas(cod_operadora,cod_plan,cod_contrato,idX,cedula,nombre,cuota,estado,aportes,rendimiento" _
          & ",telefono1,telefono2,notas,email,apto_postal,direccion,parentesco,cod_beneficiario)  values(" & cboOperadora.ItemData(cboOperadora.ListIndex) _
          & ",'" & txtCodigo & "'," & txtContrato & "," & vGrid.Text & ",'"
   vGrid.Col = 2
   strSQL = strSQL & vGrid.Text & "','"
   vGrid.Col = 3
   strSQL = strSQL & vGrid.Text & "',"
   vGrid.Col = 4
   strSQL = strSQL & CCur(vGrid.Text) & ",'A',0,0,'','','','','','','O',0)"
   Call ConectionExecute(strSQL)
   
   vGrid.Col = 1
   
   Call Bitacora("Registra", "SubCuenta " & vGrid.Text & " Plan : " & txtCodigo & " Contrato : " & txtContrato)
   
 Else 'Actualizar
    vGrid.Col = 2
    strSQL = "update fnd_subCuentas set cedula = '" & vGrid.Text & "',nombre = '"
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & "',cuota = "
    vGrid.Col = 4
    strSQL = strSQL & CCur(vGrid.Text) & ""
    vGrid.Col = 1
    strSQL = strSQL & " where idx = " & vGrid.Text & " and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " and cod_plan = '" & txtCodigo & "' and cod_contrato = " & txtContrato
   
    Call ConectionExecute(strSQL)
    
    Call sbGuardaCambios("04", "Modifica a SubCuenta " & vGrid.Text)
    
    Call Bitacora("Modifica", "SubCuenta " & vGrid.Text & " Plan : " & txtCodigo & " Contrato : " & txtContrato)
    
End If

vGrid.Col = 1
fxGuardar = vGrid.Text

Call sbActualizaCuotaContrato

Exit Function
   
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Function


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

  
If vPaso Then Exit Sub
  
  vGrid.Col = 1
  vGrid.Row = Row
  
  If vGrid.Text = "" Then
      gFondos.SubCuenta = 0
  Else
      gFondos.SubCuenta = vGrid.Text
  End If
  
  gFondos.Contrato = txtContrato.Text
  gFondos.Operadora = cboOperadora.ItemData(cboOperadora.ListIndex)
  gFondos.Plan = txtCodigo.Text
  gFondos.Cedula = txtCedula.Text
  
  frmFNDSubCuentas.Show vbModal
  
  Call sbConsulta_Detalle("General")

End Sub

Private Sub vGrid_DblClick(ByVal Col As Long, ByVal Row As Long)

If Col = 1 Then
  
  vGrid.Col = Col
  vGrid.Row = Row
  
  If vGrid.Text = "" Then
      gFondos.SubCuenta = 0
  Else
      gFondos.SubCuenta = vGrid.Text
  End If
  
  gFondos.Contrato = txtContrato
  gFondos.Operadora = cboOperadora.ItemData(cboOperadora.ListIndex)
  gFondos.Plan = txtCodigo
  gFondos.Cedula = txtCedula
  
  frmFNDSubCuentas.Show vbModal
  
  Call sbConsulta_Detalle("General")

End If


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
     sbNuevaSubCuenta
  End If
End If


If vGrid.ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = vGrid.Text
End If

End Sub


Private Sub sbGuardaCambios(vMovimiento As String, vDetalle As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "insert fnd_contratos_cambios(cod_operadora,cod_plan,cod_contrato,usuario,fecha,movimiento,detalle)values(" & _
       cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & txtCodigo.Text & "'," & txtContrato & " ,'" & glogon.Usuario & "',dbo.MyGetdate(),'" & vMovimiento & "','" & vDetalle & "')"
Call ConectionExecute(strSQL)

End Sub


Private Function fxAplicaBeneficiarios() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select REQUIERE_BENEFICIARIOS from fnd_planes where cod_plan = '" & txtCodigo & "' and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
   If rs!REQUIERE_BENEFICIARIOS = 0 Then
      fxAplicaBeneficiarios = False
   Else
      fxAplicaBeneficiarios = True
   End If
   
Else
fxAplicaBeneficiarios = False
End If
rs.Close
End Function



Private Function fxBeneficiariosNoIncluidos(iContrato As Long) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(sum(porcentaje),0) as porcentaje from FND_CONTRATOS_BENEFICIARIOS where cod_plan = '" & txtCodigo & "' and cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
   If rs!Porcentaje = 0 Or rs!Porcentaje < 100 Then
      fxBeneficiariosNoIncluidos = False
   Else
      fxBeneficiariosNoIncluidos = True
   End If
   
Else
fxBeneficiariosNoIncluidos = False
End If
rs.Close

End Function



