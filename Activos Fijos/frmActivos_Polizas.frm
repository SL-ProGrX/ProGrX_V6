VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmActivos_Polizas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro y Asignación de Polizas"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   9615
      _Version        =   1572864
      _ExtentX        =   16960
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
      ItemCount       =   2
      Item(0).Caption =   "Póliza"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "dtpInicia"
      Item(0).Control(1)=   "dtpVence"
      Item(0).Control(2)=   "txtObservacion"
      Item(0).Control(3)=   "cbo"
      Item(0).Control(4)=   "txtDescripcion"
      Item(0).Control(5)=   "txtDocumento"
      Item(0).Control(6)=   "txtNumPoliza"
      Item(0).Control(7)=   "txtMonto"
      Item(0).Control(8)=   "Label1(8)"
      Item(0).Control(9)=   "Label1(7)"
      Item(0).Control(10)=   "Label1(6)"
      Item(0).Control(11)=   "Label1(5)"
      Item(0).Control(12)=   "Label1(4)"
      Item(0).Control(13)=   "Label1(3)"
      Item(0).Control(14)=   "Label1(1)"
      Item(0).Control(15)=   "Label1(0)"
      Item(1).Caption =   "Asignación"
      Item(1).ControlCount=   12
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "btnBuscar"
      Item(1).Control(2)=   "chkLsw"
      Item(1).Control(3)=   "cboTipo"
      Item(1).Control(4)=   "txtActivoPlaca"
      Item(1).Control(5)=   "txtActivoDesc"
      Item(1).Control(6)=   "txtAsPoliza"
      Item(1).Control(7)=   "txtAsPolDesc"
      Item(1).Control(8)=   "Label1(11)"
      Item(1).Control(9)=   "scTitulo"
      Item(1).Control(10)=   "Label1(10)"
      Item(1).Control(11)=   "Label1(9)"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3375
         Left            =   -70000
         TabIndex        =   4
         Top             =   2160
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1572864
         _ExtentX        =   16960
         _ExtentY        =   5953
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
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   315
         Left            =   -61600
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
         _Version        =   1572864
         _ExtentX        =   656
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkLsw 
         Height          =   195
         Left            =   -69880
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   195
         _Version        =   1572864
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   79
         BackColor       =   -2147483633
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
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   315
         Left            =   -68320
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtActivoPlaca 
         Height          =   330
         Left            =   -68320
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
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
      Begin XtremeSuiteControls.FlatEdit txtActivoDesc 
         Height          =   330
         Left            =   -66760
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   5055
         _Version        =   1572864
         _ExtentX        =   8916
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
      Begin XtremeSuiteControls.FlatEdit txtAsPoliza 
         Height          =   330
         Left            =   -68320
         TabIndex        =   10
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
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
      Begin XtremeSuiteControls.FlatEdit txtAsPolDesc 
         Height          =   330
         Left            =   -66760
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
         _Version        =   1572864
         _ExtentX        =   8916
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicia 
         Height          =   315
         Left            =   5040
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
         _Version        =   1572864
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
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   315
         Left            =   7680
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtObservacion 
         Height          =   1995
         Left            =   1680
         TabIndex        =   18
         Top             =   2640
         Width           =   7455
         _Version        =   1572864
         _ExtentX        =   13144
         _ExtentY        =   3514
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
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   960
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
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
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   330
         Left            =   1680
         TabIndex        =   20
         Top             =   600
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
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
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   330
         Left            =   1680
         TabIndex        =   21
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1800
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtNumPoliza 
         Height          =   330
         Left            =   1680
         TabIndex        =   22
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1320
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   330
         Left            =   1680
         TabIndex        =   23
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2160
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
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
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Número"
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
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Inicia"
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
         Left            =   4200
         TabIndex        =   28
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Vence"
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
         Left            =   6840
         TabIndex        =   27
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
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
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Observación"
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
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
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
         Index           =   8
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Póliza"
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
         Height          =   255
         Index           =   9
         Left            =   -69760
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   855
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   -69760
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Left            =   -70000
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1572864
         _ExtentX        =   16960
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seleccione los Activos que tiene cobertura con esta póliza"
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Placa/Nombre"
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
         Height          =   255
         Index           =   11
         Left            =   -69760
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   435
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   2175
      _Version        =   1572864
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   375
      Left            =   360
      TabIndex        =   32
      Top             =   600
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "  xxx"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   4935
   End
End
Attribute VB_Name = "frmActivos_Polizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vEdita As Boolean, vCodigo As String, vPaso As Boolean

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNumPoliza.SetFocus
End Sub


Private Sub btnBuscar_Click()
Call sbPolizas_Asignacion
End Sub

Private Sub cboTipo_Click()
If Not vPaso Then Exit Sub
Call sbPolizas_Asignacion
End Sub


Private Sub chkLsw_Click()
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

For i = 1 To lsw.ListItems.Count
  If chkLsw.Value = vbChecked Then
     If Not lsw.ListItems.Item(i).Checked Then
        strSQL = "insert Activos_polizas_asigna(cod_poliza,num_placa) values(" _
               & txtAsPoliza & ",'" & lsw.ListItems.Item(i).Text & "')"
     End If
  Else
     If lsw.ListItems.Item(i).Checked Then
        strSQL = "delete Activos_polizas_asigna where cod_poliza = " _
               & txtAsPoliza & " and num_placa = '" & lsw.ListItems.Item(i).Text & "')"
     End If
  End If
  Call ConectionExecute(strSQL)
  lsw.ListItems.Item(i).Checked = chkLsw.Value
Next i

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub dtpInicia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpVence.SetFocus
End Sub

Private Sub dtpVence_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 36

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

vCodigo = ""
txtCodigo = ""

tcMain.Item(0).Selected = True

txtDescripcion = ""
lblEstado.Caption = ""

strSQL = "select rtrim(tipo_poliza) as 'Idx', rtrim(descripcion) as 'ItmX'" _
       & " from Activos_polizas_tipos order by tipo_poliza"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

txtMonto.Text = 0
txtNumPoliza.Text = ""
txtDocumento.Text = ""
txtObservacion.Text = ""
dtpInicia.Value = fxFechaServidor
dtpVence.Value = dtpInicia

End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert Activos_polizas_asg(cod_poliza,num_placa,registro_fecha,registro_usuario) values('" & txtAsPoliza _
          & "','" & Item.Text & "',getdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete Activos_polizas_asg where cod_poliza = '" & txtAsPoliza _
          & "' and num_placa = '" & Item.Text & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
  txtAsPoliza = ""
  txtAsPolDesc = ""
  lsw.ListItems.Clear
  
  vPaso = False
    strSQL = "select rtrim(tipo_activo) as 'IdX',  rtrim(descripcion) as 'ItmX'" _
       & " from Activos_tipo_activo order by tipo_activo"
    Call sbCbo_Llena_New(cboTipo, strSQL, False, True)
  vPaso = True

End If

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
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
       gBusquedas.Consulta = "select cod_poliza,descripcion from Activos_polizas"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(xCodigo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select P.*,rtrim(T.descripcion) as 'TipoDesc',getdate() as FechaX" _
       & " from Activos_polizas_tipos T inner join Activos_polizas P on T.tipo_poliza = P.tipo_poliza" _
       & " where P.cod_poliza = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  tcMain.Item(0).Selected = True
  
  vCodigo = rs!cod_poliza
  txtCodigo = rs!cod_poliza
 
  txtDescripcion = rs!DESCRIPCION
  txtObservacion = rs!observacion
  dtpInicia.Value = rs!fecha_inicio
  dtpVence.Value = rs!fecha_vence
  
  Call sbCboAsignaDato(cbo, rs!TipoDesc, True, rs!Tipo_Poliza)
  
  txtMonto = Format(rs!monto, "Standard")
  txtNumPoliza = rs!num_poliza
  txtDocumento = rs!Documento
  
  If rs!fecha_vence < rs!fechaX Then
    lblEstado.Caption = "Poliza Vencida"
    lblEstado.ForeColor = vbRed
  Else
    lblEstado.Caption = "Poliza Activa"
    lblEstado.ForeColor = vbGrayText
  End If

  
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
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtCodigo.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Código Interno de la poliza no es válido ..."
If txtDescripcion.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Descripción de la poliza no es válido ..."
If dtpVence.Value < dtpInicia.Value Then vMensaje = vMensaje & vbCrLf & " - La fecha de vencimiento no puede ser menor a la inicial ..."
If Not IsNumeric(txtMonto.Text) Then vMensaje = vMensaje & vbCrLf & " - Monto no es válido ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()

On Error GoTo vError

If vEdita Then
  strSQL = "update Activos_polizas set descripcion = '" & txtDescripcion.Text _
         & "', observacion = '" & txtObservacion & "',monto = " & CCur(txtMonto) _
         & ", fecha_sistema = getdate(), fecha_inicio = '" & Format(dtpInicia.Value, "yyyy/mm/dd") _
         & "', fecha_vence = '" & Format(dtpVence.Value, "yyyy/mm/dd") & "', num_poliza = '" _
         & txtNumPoliza & "', documento = '" & txtDocumento & "',tipo_poliza = '" _
         & cbo.ItemData(cbo.ListIndex) _
         & "', modifica_fecha = getdate(), modifica_usuario = '" & glogon.Usuario _
         & "' where cod_poliza = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Póliza: " & vCodigo)

Else
   
   vCodigo = txtCodigo.Text
   
   strSQL = "insert Activos_polizas(cod_poliza,tipo_poliza,descripcion,observacion,fecha_sistema,fecha_inicio,fecha_vence" _
          & ",monto,num_poliza,documento,registro_fecha,registro_usuario) values('" & vCodigo & "','" _
          & cbo.ItemData(cbo.ListIndex) & "','" & txtDescripcion.Text _
          & "','" & txtObservacion & "',getdate(),'" & Format(dtpInicia.Value, "yyyy/mm/dd") _
          & "','" & Format(dtpVence.Value, "yyyy/mm/dd") & "'," & CCur(txtMonto) & ",'" _
          & txtNumPoliza & "','" & txtDocumento & "',getdate(),'" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Póliza: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbConsulta(txtCodigo.Text)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete Activos_polizas where cod_poliza = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Póliza: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbPolizas_Asignacion()

On Error GoTo vError


If txtAsPoliza = "" Then Exit Sub

Me.MousePointer = vbHourglass


lsw.ListItems.Clear
lsw.Checkboxes = True
With lsw.ColumnHeaders
    .Clear
    .Add , , "Placa", 1800
    .Add , , "Nombre", 5700
    .Add , , "Estado", 1800, vbCenter
End With

strSQL = "select A.num_placa,A.nombre,A.estado,P.cod_poliza" _
       & " from Activos_Principal A left join Activos_polizas_asg P" _
       & " on A.num_placa = P.num_placa and P.cod_poliza = '" & txtAsPoliza & "'" _
       & " where A.tipo_activo = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
       
       
If Len(txtActivoPlaca.Text) > 0 Then
    strSQL = strSQL & " and A.Num_Placa like '%" & txtActivoPlaca.Text & "%'"
End If

If Len(txtActivoDesc.Text) > 0 Then
    strSQL = strSQL & " and A.nombre like '%" & txtActivoDesc.Text & "%'"
End If
       
strSQL = strSQL & " order by A.num_placa"

vPaso = True

    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!NUM_PLACA)
         itmX.SubItems(1) = rs!Nombre
         itmX.SubItems(2) = IIf((rs!Estado = "A"), "VIGENTE", "RETIRADO")
         itmX.Checked = IIf(IsNull(rs!cod_poliza), False, True)
     rs.MoveNext
    Loop
    rs.Close

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtActivoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Call sbPolizas_Asignacion
End If
End Sub

Private Sub txtActivoPlaca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call sbPolizas_Asignacion
End If
End Sub

Private Sub txtAsPolDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "P.cod_poliza"
  gBusquedas.Orden = "P.cod_poliza"
  gBusquedas.Consulta = "select P.cod_poliza,P.descripcion,T.descripcion as Tipo" _
                      & " from Activos_polizas P inner join Activos_polizas_Tipos T on P.tipo_poliza = T.tipo_poliza"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtAsPoliza = gBusquedas.Resultado
  txtAsPolDesc = gBusquedas.Resultado2
  Call sbPolizas_Asignacion
End If
End Sub

Private Sub txtAsPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "P.cod_poliza"
  gBusquedas.Orden = "P.cod_poliza"
  gBusquedas.Consulta = "select P.cod_poliza,P.descripcion,T.descripcion as Tipo" _
                      & " from Activos_polizas P inner join Activos_polizas_Tipos T on P.tipo_poliza = T.tipo_poliza"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtAsPoliza = gBusquedas.Resultado
  txtAsPolDesc = gBusquedas.Resultado2
  Call sbPolizas_Asignacion
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" Then Call sbConsulta(txtCodigo)
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_poliza,descripcion from Activos_polizas"
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
  gBusquedas.Consulta = "select cod_poliza,descripcion from Activos_polizas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub


Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
 txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
 txtMonto = Format(CCur(txtMonto), "Standard")
vError:
End Sub


Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservacion.SetFocus
End Sub

Private Sub txtNumPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicia.SetFocus
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub
