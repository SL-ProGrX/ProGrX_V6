VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmActivos_TiposActivo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Activos"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3735
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   9855
      _Version        =   1441793
      _ExtentX        =   17378
      _ExtentY        =   6583
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
      ItemCount       =   1
      Item(0).Caption =   "General"
      Item(0).ControlCount=   22
      Item(0).Control(0)=   "Label6(6)"
      Item(0).Control(1)=   "Label6(5)"
      Item(0).Control(2)=   "Label6(4)"
      Item(0).Control(3)=   "Label6(3)"
      Item(0).Control(4)=   "Label6(2)"
      Item(0).Control(5)=   "Label6(1)"
      Item(0).Control(6)=   "Label6(0)"
      Item(0).Control(7)=   "txtDescripcion"
      Item(0).Control(8)=   "cbo"
      Item(0).Control(9)=   "cboVidaUtil"
      Item(0).Control(10)=   "txtVidaUtil"
      Item(0).Control(11)=   "Label6(7)"
      Item(0).Control(12)=   "txtAsiento"
      Item(0).Control(13)=   "txtAsientoDesc"
      Item(0).Control(14)=   "txtCtaActivo"
      Item(0).Control(15)=   "txtCtaActivoDesc"
      Item(0).Control(16)=   "txtCtaDepAcum"
      Item(0).Control(17)=   "txtCtaDepAcumDesc"
      Item(0).Control(18)=   "txtCtaGastos"
      Item(0).Control(19)=   "txtCtaGastosDesc"
      Item(0).Control(20)=   "txtCtaTransitoria"
      Item(0).Control(21)=   "txtCtaTransitoriaDesc"
      Begin XtremeSuiteControls.FlatEdit txtCtaTransitoriaDesc 
         Height          =   330
         Left            =   3960
         TabIndex        =   2
         Top             =   3240
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaGastosDesc 
         Height          =   330
         Left            =   3960
         TabIndex        =   25
         Top             =   2880
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaDepAcumDesc 
         Height          =   330
         Left            =   3960
         TabIndex        =   23
         Top             =   2520
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaActivoDesc 
         Height          =   330
         Left            =   3960
         TabIndex        =   21
         Top             =   2160
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   330
         Left            =   1800
         TabIndex        =   12
         Top             =   480
         Width           =   7812
         _Version        =   1441793
         _ExtentX        =   13779
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   1800
         TabIndex        =   14
         Top             =   960
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.ComboBox cboVidaUtil 
         Height          =   312
         Left            =   7800
         TabIndex        =   15
         Top             =   960
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtVidaUtil 
         Height          =   330
         Left            =   6720
         TabIndex        =   16
         Top             =   960
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAsientoDesc 
         Height          =   330
         Left            =   3960
         TabIndex        =   19
         Top             =   1800
         Width           =   5772
         _Version        =   1441793
         _ExtentX        =   10181
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAsiento 
         Height          =   330
         Left            =   1800
         TabIndex        =   18
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1800
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaActivo 
         Height          =   330
         Left            =   1800
         TabIndex        =   20
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2160
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaDepAcum 
         Height          =   330
         Left            =   1800
         TabIndex        =   22
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2520
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaGastos 
         Height          =   330
         Left            =   1800
         TabIndex        =   24
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2880
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaTransitoria 
         Height          =   330
         Left            =   1800
         TabIndex        =   13
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3240
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   7
         Left            =   5160
         TabIndex        =   17
         Top             =   960
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Vida Util"
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
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Descripción"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Depreciación"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo Asiento"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta Activo"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   4
         Left            =   360
         TabIndex        =   7
         Top             =   2520
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta Dep.Acum."
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   2880
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta Gastos"
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
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   6
         Left            =   360
         TabIndex        =   5
         Top             =   3240
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Transitoria Adq."
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
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4950
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario que Registra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario que Modifica"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Fecha de Actualización"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   435
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4254
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   4320
      TabIndex        =   26
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   1800
      TabIndex        =   27
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_TiposActivo.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   2880
      TabIndex        =   28
      ToolTipText     =   "Editar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_TiposActivo.frx":0632
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   3240
      TabIndex        =   29
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_TiposActivo.frx":0C2D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   3840
      TabIndex        =   30
      ToolTipText     =   "Guardar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_TiposActivo.frx":11D1
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   4200
      TabIndex        =   31
      ToolTipText     =   "Deshacer"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_TiposActivo.frx":1902
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   4680
      TabIndex        =   32
      ToolTipText     =   "Reporte"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_TiposActivo.frx":2002
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.Label LabelX 
      Height          =   435
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   762
      _StockProps     =   79
      Caption         =   "Tipo Activo"
      ForeColor       =   0
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmActivos_TiposActivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean

Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub btnBarra_Click(Index As Integer)
Dim strSQL As String


Select Case Index
    Case 0 'NUEVO
        vEdita = False
        Call sbLimpiaPantalla
        txtCodigo.SetFocus

        Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
        vEdita = True
        txtDescripcion.SetFocus
        Call sbBarra_Accion("Editar")
      
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
      Call sbBarra_Accion("Editar")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
    
    Case 5 'REPORTES
   
End Select


End Sub



Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVidaUtil.SetFocus
End Sub


Private Sub cboVidaUtil_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAsiento.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not vScroll Then Exit Sub

strSQL = "select Top 1 TIPO_ACTIVO from ACTIVOS_TIPO_ACTIVO"

If FlatScrollBar.Value = 1 Then
   strSQL = strSQL & " where TIPO_ACTIVO > '" & txtCodigo.Text & "' order by TIPO_ACTIVO asc"
Else
   strSQL = strSQL & " where TIPO_ACTIVO < '" & txtCodigo.Text & "' order by TIPO_ACTIVO desc"
End If

Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtCodigo.Text = rs!tipo_activo
  Call sbConsulta(txtCodigo)
End If
rs.Close

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 36


End Sub

Private Sub Form_Load()

On Error GoTo vError
 
vModulo = 36


vScroll = False
FlatScrollBar.Value = 0
vScroll = True
 
cboVidaUtil.Clear
cboVidaUtil.AddItem "Años"
cboVidaUtil.AddItem "Meses"
cboVidaUtil.Text = "Años"
 

 vEdita = False
 
 Call sbBarra_Accion("Nuevo")
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

Call sbActivos_MetodosDepreciacion(cbo)
cboVidaUtil.Text = "Años"

txtDescripcion = ""
txtVidaUtil = ""
txtAsiento = ""
txtAsientoDesc.Text = ""

txtCtaActivo = ""
txtCtaActivoDesc = ""

txtCtaDepAcum = ""
txtCtaDepAcumDesc = ""

txtCtaGastos = ""
txtCtaGastosDesc = ""

With StatusBarX.Panels
  .Item(1).Text = ""
  .Item(2).Text = ""
  .Item(3).Text = ""
  .Item(4).Text = ""
End With

End Sub


Public Sub sbConsultaExterna(pTipo As String)
If pTipo <> "" Then
 Call sbConsulta(pTipo)
End If
End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from Activos_tipo_activo where tipo_activo = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
   Call sbBarra_Accion("activo")
  vEdita = True
  
  vCodigo = rs!tipo_activo
  txtCodigo = rs!tipo_activo
 
  txtDescripcion.Text = rs!Descripcion
  txtDescripcion.SetFocus
  
    
  cbo.Text = fxActivos_MetodoDepreciacion(rs!met_depreciacion)
  txtVidaUtil = rs!Vida_Util
  If rs!tipo_vida_util = "A" Then
    cboVidaUtil.Text = "Años"
  Else
    cboVidaUtil.Text = "Meses"
  End If
    
  txtAsiento.Text = rs!Asiento_Genera
  txtCtaActivo.Text = fxgCntCuentaFormato(True, rs!cod_cuenta_activo, 0)
  txtCtaGastos.Text = fxgCntCuentaFormato(True, rs!cod_cuenta_gastos, 0)
  txtCtaDepAcum.Text = fxgCntCuentaFormato(True, rs!cod_cuenta_DepAcum, 0)
  txtCtaTransitoria.Text = fxgCntCuentaFormato(True, rs!cod_cuenta_Transitoria, 0)
  
  txtAsientoDesc.Text = fxgCntTipoAsientoDesc(rs!Asiento_Genera)
  txtCtaActivoDesc.Text = fxgCntCuentaDesc(rs!cod_cuenta_activo)
  txtCtaGastosDesc.Text = fxgCntCuentaDesc(rs!cod_cuenta_gastos)
  txtCtaDepAcumDesc.Text = fxgCntCuentaDesc(rs!cod_cuenta_DepAcum)
  txtCtaTransitoriaDesc.Text = fxgCntCuentaDesc(rs!cod_cuenta_Transitoria)
    
    With StatusBarX.Panels
      .Item(1).Text = rs!registro_usuario & ""
      .Item(2).Text = rs!registro_fecha & ""
      .Item(3).Text = rs!Modifica_Usuario & ""
      .Item(4).Text = rs!Modifica_Fecha & ""
    End With
  
  
Else
  If vEdita Then
      MsgBox "No se encontró registro verifique...", vbInformation
  End If
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

'Validar Cuentas Aqui

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion del tipo de Activo no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update Activos_tipo_activo set descripcion = '" & UCase(txtDescripcion) _
         & "',met_depreciacion = '" & fxActivos_MetodoDepreciacion(cbo.Text) & "',asiento_genera = '" _
         & txtAsiento & "',cod_cuenta_activo = '" & fxgCntCuentaFormato(False, txtCtaActivo) _
         & "',cod_cuenta_DepAcum = '" & fxgCntCuentaFormato(False, txtCtaDepAcum) _
         & "',cod_cuenta_gastos = '" & fxgCntCuentaFormato(False, txtCtaGastos) _
         & "',cod_cuenta_Transitoria = '" & fxgCntCuentaFormato(False, txtCtaTransitoria) _
         & "',Tipo_Vida_Util = '" & Mid(cboVidaUtil.Text, 1, 1) & "',Vida_Util = " & txtVidaUtil _
         & ", modifica_usuario = '" & glogon.Usuario & "', modifica_fecha = getdate()" _
         & " where tipo_Activo = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Tipo Activo : " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert into Activos_tipo_activo(tipo_activo,descripcion,asiento_genera,met_depreciacion" _
          & ",tipo_vida_util,vida_util,cod_cuenta_activo,cod_cuenta_DepAcum,cod_cuenta_Gastos,cod_cuenta_Transitoria,registro_usuario,registro_fecha) values('" _
          & vCodigo & "','" & UCase(txtDescripcion) & "','" & txtAsiento & "','" & fxActivos_MetodoDepreciacion(cbo.Text) _
          & "','" & Mid(cboVidaUtil.Text, 1, 1) & "'," & txtVidaUtil & ",'" & fxgCntCuentaFormato(False, txtCtaActivo) _
          & "','" & fxgCntCuentaFormato(False, txtCtaDepAcum) & "','" & fxgCntCuentaFormato(False, txtCtaGastos) _
          & "','" & fxgCntCuentaFormato(False, txtCtaTransitoria) & "','" & glogon.Usuario & "',getdate())"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Tipo de Activo: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(txtCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete Activos_tipo_activo where tipo_activo = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
'  Call sbBitacora("Elimina", "Tipo Activo : " & vCodigo)
  Call sbLimpiaPantalla
   Call sbBarra_Accion("nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaActivo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
  gBusquedas.Filtro = " and COD_CONTABILIDAD = " & GLOBALES.gEnlace
  gBusquedas.Columna = "tipo_asiento"
  gBusquedas.Orden = "tipo_asiento"
  frmBusquedas.Show vbModal
  txtAsiento.Text = gBusquedas.Resultado
  txtAsientoDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtAsiento_LostFocus()
txtAsientoDesc.Text = fxgCntTipoAsientoDesc(txtAsiento.Text)
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Tipo_Activo"
  gBusquedas.Orden = "Tipo_Activo"
  gBusquedas.Consulta = "select Tipo_Activo,descripcion from Activos_Tipo_Activo"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()
If txtCodigo.Text <> "" Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtCtaActivo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaActivoDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaActivo = gCuenta
End If
End Sub

Private Sub txtCtaActivo_LostFocus()
If txtCtaActivo.Text <> "" Then
    txtCtaActivo.Text = fxgCntCuentaFormato(True, txtCtaActivo.Text)
    txtCtaActivoDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaActivo.Text, 0))
End If
End Sub

Private Sub txtCtaActivoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaDepAcum.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaActivo = gCuenta
   txtCtaActivo.SetFocus
End If
End Sub

Private Sub txtCtaDepAcum_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaDepAcumDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaDepAcum = gCuenta
End If
End Sub


Private Sub txtCtaDepAcum_LostFocus()
If txtCtaDepAcum.Text <> "" Then
    txtCtaDepAcum.Text = fxgCntCuentaFormato(True, txtCtaDepAcum.Text)
    txtCtaDepAcumDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaDepAcum, 0))
End If
End Sub

Private Sub txtCtaDepAcumDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaGastos.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaDepAcum = gCuenta
   txtCtaDepAcum.SetFocus
End If
End Sub


Private Sub txtCtaGastos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaGastosDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaGastos = gCuenta
End If
End Sub

Private Sub txtCtaGastos_LostFocus()
If txtCtaGastos.Text <> "" Then
    txtCtaGastos.Text = fxgCntCuentaFormato(True, txtCtaGastos.Text)
    txtCtaGastosDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaGastos.Text, 0))
End If
End Sub

Private Sub txtCtaGastosDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaTransitoria.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaGastos = gCuenta
   txtCtaGastos.SetFocus
End If
End Sub







Private Sub txtCtaTransitoria_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaTransitoriaDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaTransitoria.Text = gCuenta
End If
End Sub


Private Sub txtCtaTransitoria_LostFocus()
If txtCtaTransitoria.Text <> "" Then
    txtCtaTransitoria.Text = fxgCntCuentaFormato(True, txtCtaTransitoria.Text)
    txtCtaTransitoriaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaTransitoria.Text, 0))
End If
End Sub

Private Sub txtCtaTransitoriaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaActivo.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaTransitoria.Text = gCuenta
   txtCtaTransitoria.SetFocus
End If
End Sub



Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select tipo_activo,descripcion from Activos_tipo_activo"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub

Private Sub txtVidaUtil_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboVidaUtil.SetFocus
End Sub
