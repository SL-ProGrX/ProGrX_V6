VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCO_Incobrables 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Incobrables"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9240
      Top             =   240
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   2892
      Left            =   0
      TabIndex        =   38
      Top             =   4080
      Width           =   10932
      _Version        =   1441792
      _ExtentX        =   19283
      _ExtentY        =   5101
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
      Item(0).Caption =   "Registro de Incobrable"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "txtNotas"
      Item(0).Control(1)=   "chkIncobrable"
      Item(0).Control(2)=   "chkSaldos"
      Item(0).Control(3)=   "Label2(5)"
      Item(0).Control(4)=   "btnPrincipal(0)"
      Item(1).Caption =   "Reversión de Incobrable"
      Item(1).ControlCount=   14
      Item(1).Control(0)=   "btnPrincipal(1)"
      Item(1).Control(1)=   "txtRegistroFecha"
      Item(1).Control(2)=   "txtRegistroUsuario"
      Item(1).Control(3)=   "txtRegistroDocumento"
      Item(1).Control(4)=   "Label2(4)"
      Item(1).Control(5)=   "Label2(2)"
      Item(1).Control(6)=   "txtReversionFecha"
      Item(1).Control(7)=   "txtReversionUsuario"
      Item(1).Control(8)=   "txtReversionDocumento"
      Item(1).Control(9)=   "txtReversionRecargo"
      Item(1).Control(10)=   "Label2(10)"
      Item(1).Control(11)=   "Label2(3)"
      Item(1).Control(12)=   "txtReversionNotas"
      Item(1).Control(13)=   "Label2(6)"
      Begin XtremeSuiteControls.PushButton btnPrincipal 
         Height          =   732
         Index           =   0
         Left            =   2880
         TabIndex        =   43
         Top             =   2160
         Width           =   2652
         _Version        =   1441792
         _ExtentX        =   4678
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Registar Incobrable"
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
         Picture         =   "frmCO_Incobrables.frx":0000
      End
      Begin XtremeSuiteControls.CheckBox chkIncobrable 
         Height          =   252
         Left            =   2880
         TabIndex        =   40
         Top             =   360
         Width           =   3612
         _Version        =   1441792
         _ExtentX        =   6371
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Activar Categoria Crediticia de Incobrable"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkSaldos 
         Height          =   252
         Left            =   2880
         TabIndex        =   41
         Top             =   720
         Width           =   3612
         _Version        =   1441792
         _ExtentX        =   6371
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplicar únicamente el Saldo sin los intereses"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnPrincipal 
         Height          =   732
         Index           =   1
         Left            =   -62080
         TabIndex        =   45
         Top             =   2040
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1441792
         _ExtentX        =   4678
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Reversar Incobrable"
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
         Picture         =   "frmCO_Incobrables.frx":07D8
      End
      Begin XtremeSuiteControls.FlatEdit txtRegistroDocumento 
         Height          =   312
         Left            =   -67720
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441792
         _ExtentX        =   4254
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtRegistroFecha 
         Height          =   312
         Left            =   -65320
         TabIndex        =   53
         Top             =   960
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441792
         _ExtentX        =   4254
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtRegistroUsuario 
         Height          =   312
         Left            =   -67720
         TabIndex        =   54
         Top             =   960
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441792
         _ExtentX        =   4254
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtReversionFecha 
         Height          =   312
         Left            =   -65320
         TabIndex        =   55
         Top             =   1680
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441792
         _ExtentX        =   4254
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtReversionUsuario 
         Height          =   312
         Left            =   -67720
         TabIndex        =   56
         Top             =   1680
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441792
         _ExtentX        =   4254
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtReversionDocumento 
         Height          =   312
         Left            =   -65320
         TabIndex        =   57
         Top             =   1320
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441792
         _ExtentX        =   4254
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtReversionRecargo 
         Height          =   312
         Left            =   -67720
         TabIndex        =   58
         Top             =   1320
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441792
         _ExtentX        =   4254
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   912
         Left            =   2880
         TabIndex        =   39
         Top             =   1080
         Width           =   7812
         _Version        =   1441792
         _ExtentX        =   13779
         _ExtentY        =   1609
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtReversionNotas 
         Height          =   792
         Left            =   -67720
         TabIndex        =   50
         Top             =   2040
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1441792
         _ExtentX        =   8488
         _ExtentY        =   1397
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
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
         Index           =   6
         Left            =   -69760
         TabIndex        =   51
         Top             =   2160
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Recargo y Documento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   3
         Left            =   -69760
         TabIndex        =   49
         Top             =   1320
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Reversa"
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
         Index           =   10
         Left            =   -69760
         TabIndex        =   48
         Top             =   1680
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento Generado"
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
         Left            =   -69760
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Genera"
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
         Index           =   4
         Left            =   -69760
         TabIndex        =   46
         Top             =   960
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
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
         Index           =   5
         Left            =   1440
         TabIndex        =   42
         Top             =   1080
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   432
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1812
      _Version        =   1441792
      _ExtentX        =   3196
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   432
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   2052
      _Version        =   1441792
      _ExtentX        =   3619
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   432
      Left            =   6720
      TabIndex        =   2
      Top             =   120
      Width           =   1092
      _Version        =   1441792
      _ExtentX        =   1926
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   1812
      _Version        =   1441792
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2880
      TabIndex        =   4
      Top             =   960
      Width           =   1812
      _Version        =   1441792
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4680
      TabIndex        =   5
      Top             =   960
      Width           =   6012
      _Version        =   1441792
      _ExtentX        =   10604
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   4680
      TabIndex        =   6
      Top             =   600
      Width           =   6012
      _Version        =   1441792
      _ExtentX        =   10604
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbDeuda 
      Height          =   2532
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   10812
      _Version        =   1441792
      _ExtentX        =   19071
      _ExtentY        =   4466
      _StockProps     =   79
      Caption         =   "Deuda"
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
      Begin XtremeSuiteControls.DateTimePicker dtpCalculoIntCorte 
         Height          =   312
         Left            =   2400
         TabIndex        =   11
         Top             =   4560
         Width           =   1212
         _Version        =   1441792
         _ExtentX        =   2138
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtIntMor 
         Height          =   312
         Left            =   2880
         TabIndex        =   12
         Top             =   1080
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   312
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIntCor 
         Height          =   312
         Left            =   2880
         TabIndex        =   14
         Top             =   720
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPoliza 
         Height          =   312
         Left            =   2880
         TabIndex        =   15
         Top             =   2160
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAmortizacion 
         Height          =   312
         Left            =   2880
         TabIndex        =   16
         Top             =   1440
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCargos 
         Height          =   312
         Left            =   2880
         TabIndex        =   17
         Top             =   1800
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCbrIntereses 
         Height          =   312
         Left            =   2400
         TabIndex        =   18
         Top             =   4080
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalMora 
         Height          =   312
         Left            =   6960
         TabIndex        =   19
         Top             =   1440
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalMoraLegal 
         Height          =   312
         Left            =   6960
         TabIndex        =   20
         Top             =   1800
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalAtrasado 
         Height          =   312
         Left            =   6960
         TabIndex        =   21
         Top             =   2160
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   312
         Left            =   8760
         TabIndex        =   33
         Top             =   2160
         Width           =   612
         _Version        =   1441792
         _ExtentX        =   1080
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
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   312
         Left            =   6960
         TabIndex        =   34
         Top             =   360
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtGarantia 
         Height          =   312
         Left            =   6960
         TabIndex        =   35
         Top             =   720
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   8760
         TabIndex        =   44
         Top             =   360
         Width           =   972
         _Version        =   1441792
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.FlatEdit txtEstadoMoroso 
         Height          =   312
         Left            =   6960
         TabIndex        =   59
         Top             =   1080
         Width           =   1812
         _Version        =   1441792
         _ExtentX        =   3196
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Antiguedad"
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
         Left            =   5040
         TabIndex        =   60
         Top             =   1080
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Index           =   23
         Left            =   5040
         TabIndex        =   37
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Garantia"
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
         Index           =   24
         Left            =   5040
         TabIndex        =   36
         Top             =   720
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Left            =   960
         TabIndex        =   32
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Interes Corriente"
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
         Index           =   22
         Left            =   960
         TabIndex        =   31
         Top             =   720
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Interes Moratorio"
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
         Left            =   960
         TabIndex        =   30
         Top             =   1080
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pólizas atrasadas"
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
         Index           =   19
         Left            =   960
         TabIndex        =   29
         Top             =   2160
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargos registrados"
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
         Left            =   960
         TabIndex        =   28
         Top             =   1800
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Principal atrasado"
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
         Index           =   11
         Left            =   960
         TabIndex        =   27
         Top             =   1440
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Mora Legal"
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
         Index           =   20
         Left            =   5040
         TabIndex        =   26
         Top             =   1800
         Width           =   2052
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Mora Financiera"
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
         Left            =   5040
         TabIndex        =   25
         Top             =   1440
         Width           =   1932
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deuda"
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
         Index           =   17
         Left            =   5040
         TabIndex        =   24
         Top             =   2160
         Width           =   1932
      End
      Begin VB.Label Label2 
         Caption         =   "Corte Intereses"
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
         Left            =   480
         TabIndex        =   23
         Top             =   4572
         Width           =   1812
      End
      Begin VB.Label Label2 
         Caption         =   "Intereses a Hoy"
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
         Index           =   18
         Left            =   480
         TabIndex        =   22
         Top             =   4080
         Width           =   1212
      End
      Begin VB.Image imgCalculoInt 
         Height          =   252
         Index           =   0
         Left            =   3600
         Picture         =   "frmCO_Incobrables.frx":1165
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   252
      End
      Begin VB.Image imgCalculoInt 
         Height          =   252
         Index           =   1
         Left            =   3960
         Picture         =   "frmCO_Incobrables.frx":1AE2
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   252
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
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
      Left            =   1440
      TabIndex        =   9
      Top             =   600
      Width           =   1572
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      TabIndex        =   8
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   11
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   1572
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12612
   End
End
Attribute VB_Name = "frmCO_Incobrables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mOperacion As Long, mcurIntCor As Currency, mcurIntMor As Currency, mcurPrincipalMora As Currency, vPaso As Boolean
Dim mGarantia As String, mcurCargos As Currency, mcurPoliza As Currency



Private Sub btnPrincipal_Click(Index As Integer)
Dim i As Byte

Call sbSIFCleanTxtInject(txtNotas)
Call sbSIFCleanTxtInject(txtReversionNotas)


Select Case Index
  Case 0 'Aplicar"
           i = MsgBox("Esta Seguro de realizar el Registro del Incobrable?", vbYesNo)
           If i = vbYes Then
              Call sbAplica
           End If
  Case 1 ' "Reversar"
           i = MsgBox("Esta Seguro de realizar la reversión del Incobrable?", vbYesNo)
           If i = vbYes Then
              Call sbReversa
           End If
 
End Select
 
End Sub

Private Sub cbo_Click()

If vPaso Then Exit Sub

Call sbCambiaCod

End Sub


Private Sub sbConsulta()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Actualizar la informacion de la ventana segun la operacion seleccionada
'REFERENCIAS   : sbMoraActiva - (Carga Datos de Mora Activa de la Operacion)
'                fxDescribeCodigo - (Devuelve la descripcion de el código del crédito)
'                sbBoletaAfiliacion - (Carga los datos personales)
'OBSERVACIONES : Ver Traspaso de Deudas
'-------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset, strSQL As String, vFecha As Date


On Error GoTo vError

If Not IsNumeric(txtOperacion.Text) Then Exit Sub

Me.MousePointer = vbHourglass


txtSaldo.Text = "0"

 strSQL = "select R.id_solicitud, R.Codigo, R.cedula, dbo.MyGetdate() as 'FechaServer',R.saldo,R.Estado,R.Proceso,R.Opex" _
        & ",C.Descripcion as 'LineaDesc',S.nombre, G.Descripcion as 'GarantiaDesc'" _
        & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
        & " inner join Socios S on R.cedula = S.cedula" _
        & "  left join Crd_Garantia_Tipos G on R.Garantia = G.Garantia" _
        & " where R.id_solicitud = " & txtOperacion
Call OpenRecordSet(rs, strSQL)
 
 If rs.EOF And rs.BOF Then
   Exit Sub
 Else
    txtCodigo.Text = rs!Codigo
    txtDescripcion.Text = rs!LineaDesc
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
        
    txtSaldo.Text = Format(rs!Saldo, "Standard")
        
    vFecha = rs!FechaServer
   
    txtEstado.Text = fxDescribeEstado(IIf(IsNull(rs!Estado), "N", rs!Estado))
    
    Select Case UCase(IIf(IsNull(rs!Proceso), "N", rs!Proceso))
     Case "N"
      txtProceso.Text = "NORMAL"
     Case "T"
      txtProceso.Text = "TRASPASO DEUDAS"
     Case "J"
      txtProceso.Text = "COBRO JUDICIAL"
     Case Else
      txtProceso.Text = "NORMAL"
    End Select
    
    If IIf(IsNull(rs!opex), 0, rs!opex) = 1 Then
       txtOpex.Text = "SI"
    Else
       txtOpex.Text = "NO"
    End If
    
    txtGarantia.Text = rs!GarantiaDesc & ""
    
 End If
 rs.Close


strSQL = "exec spCbrCobroJudicialInteresesHoy " & txtOperacion.Text & ",'" & Format(vFecha, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)

mcurIntCor = rs!RegIntCor
mcurIntMor = rs!RegIntMor
mcurPoliza = rs!Poliza
mcurCargos = rs!Cargos
mcurPrincipalMora = rs!RegPrincipal

txtIntCor.Text = Format(rs!RegIntCor, "Standard")
txtIntMor.Text = Format(rs!RegIntMor, "Standard")
txtAmortizacion.Text = Format(rs!RegPrincipal, "Standard")
txtCargos.Text = Format(rs!Cargos, "Standard")
txtPoliza.Text = Format(rs!Poliza, "Standard")

txtTotalMora.Text = Format(rs!RegIntCor + rs!RegIntMor + rs!Cargos + rs!Poliza + rs!RegPrincipal, "Standard")
txtTotalMoraLegal.Text = Format(rs!RegIntCor + rs!RegIntMor + rs!Cargos + rs!Poliza + CCur(txtSaldo.Text), "Standard")


txtTotalAtrasado.Text = txtTotalMoraLegal.Text

txtEstadoMoroso.Text = rs!Antiguedad

rs.Close


 Me.MousePointer = vbDefault

Exit Sub


vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mOperacion = GLOBALES.gTag
txtOperacion.Text = GLOBALES.gTag

Call sbConsulta


End Sub

Private Sub sbConsultaInicial()
Dim strSQL As String, rs As New ADODB.Recordset

txtOperacion = mOperacion

vPaso = True
cbo.Clear
cbo.AddItem "0"
cbo.Text = "0"

strSQL = "select cod_incobrable from cbr_incobrables where id_solicitud = " & mOperacion
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  cbo.AddItem CStr(rs!cod_incobrable)
  rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveLast
   cbo.Text = CStr(rs!cod_incobrable)
End If
rs.Close

vPaso = False

Call cbo_Click

End Sub


Private Sub sbCambiaCod()
Dim strSQL As String, rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim vUltimaCuota As Currency, vProceso As Currency, curInteres As Currency
Dim i As Integer, iMeses As Integer, vFecha As Date


''Desactiva barra
'For i = 1 To btnPrincipal.Count - 2
'  btnPrincipal.Item(i).Enabled = False
'Next i
'
tcMain.Item(0).Visible = True
tcMain.Item(1).Visible = True


txtReversionNotas.Text = ""
txtReversionFecha.Text = ""
txtReversionUsuario.Text = ""
txtReversionRecargo.Text = "0.00"
txtReversionDocumento.Text = ""


txtRegistroUsuario.Text = ""
txtRegistroFecha.Text = ""
txtRegistroDocumento.Text = ""

txtCargos.Text = "0.00"
txtPoliza.Text = "0.00"
txtAmortizacion.Text = "0.00"
txtTotalAtrasado.Text = "0.00"


strSQL = "select I.*,R.garantia " _
      & " from cbr_incobrables I inner join reg_Creditos R on I.id_solicitud = R.id_solicitud" _
      & " Where I.id_solicitud = " & mOperacion & " And I.cod_incobrable = " & cbo.Text
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
    'Consulta el estado de la operación
    rs.Close
    
    If GLOBALES.SysPlanPagos = 1 Then
          strSQL = "select R.cedula,S.nombre,R.saldo,R.proceso,R.prideduc,R.fecult,dbo.MyGetdate() as FechaActual,R.interesv,R.cod_Divisa" _
               & ",R.codigo,C.descripcion,R.Opex,R.garantia,isnull(sum(V.Principal),0) as MoraAmortiza,dbo.MyGetdate() as 'FechaServer'" _
               & " from Socios S inner join reg_creditos R on S.cedula = R.cedula inner join Catalogo C on R.codigo = C.codigo" _
               & " left join crd_operacion_Transac V on R.id_solicitud = V.id_solicitud and V.estado = 'A' and V.mora_dias > 0" _
               & " Where R.id_solicitud = " & mOperacion _
               & " Group by R.cedula,S.nombre,R.saldo,R.proceso,R.codigo,C.descripcion,R.Opex,R.interesv,R.prideduc,R.fecUlt,R.garantia,R.cod_Divisa"
    
    Else
        strSQL = "select R.cedula,S.nombre,R.saldo,R.proceso,R.prideduc,R.fecult,dbo.MyGetdate() as FechaActual,R.interesv,R.cod_Divisa" _
               & ",isnull(sum(V.intc + V.intm),0) as Intereses,R.codigo,C.descripcion,R.Opex,R.garantia" _
               & ",isnull(sum(V.intc),0) as MoraIntC,isnull(sum(V.intm),0) as MoraIntM,isnull(sum(V.amortiza),0) as MoraAmortiza" _
               & ",isnull(sum(V.Cargo),0) as 'Cargos', 0 as 'Poliza'" _
               & " from Socios S inner join reg_creditos R on S.cedula = R.cedula inner join Catalogo C on R.codigo = C.codigo" _
               & " left join morosidad V on R.id_solicitud = V.id_solicitud and V.estado = 'A'" _
               & " Where R.id_solicitud = " & mOperacion _
               & " Group by R.cedula,S.nombre,R.saldo,R.proceso,R.codigo,C.descripcion,R.Opex,R.interesv,R.prideduc,R.fecUlt,R.garantia,R.cod_Divisa"
    End If 'SysPlanPagos = 1
    
    Call OpenRecordSet(rs, strSQL)
    
    mcurPrincipalMora = rs!MoraAmortiza
    mGarantia = rs!Garantia
    vUltimaCuota = rs!FecUlt
    vFecha = rs!FechaActual
    vProceso = Year(rs!FechaActual) & Format(Month(rs!FechaActual), "00")
    
    If GLOBALES.SysPlanPagos = 1 Then
        'Actualiza Int.Moratorio Plan de Pagos.
        strSQL = "exec spCrdPlanPagosMoraActualizaOp " & mOperacion & ",'" & Format(rs!FechaServer, "yyyy/mm/dd") & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = "exec spCrdPlanPagosInfoCancelacion " & mOperacion & ",'" & Format(rs!FechaServer, "yyyy/mm/dd") & "'"
        Call OpenRecordSet(rsTmp, strSQL, 0)
          mcurIntCor = rsTmp!IntCor
          mcurIntMor = rsTmp!IntMor
          mcurCargos = rsTmp!Cargos
          mcurPoliza = rsTmp!Poliza
        rsTmp.Close
    Else
        'Modelo sin Plan de Pagos
        mcurIntCor = rs!MoraIntC
        mcurIntMor = rs!MoraIntM
        mcurCargos = rs!Cargos
        mcurPoliza = rs!Poliza
        'Si existe morosidad, Preguntar si la ultima cuota en mora en igual o mayor al proceso
        ' en caso de ser afirmativo entonces no registrar los dias transcurridos.
        ' si no hay mora el proceso es igual al proceso de mora
        
        If rs!MoraAmortiza + rs!MoraIntC + rs!MoraIntM > 0 Then
            strSQL = "select max(fechap) as Proceso from morosidad where estado = 'A' and id_solicitud = " & mOperacion
            Call OpenRecordSet(rsTmp, strSQL, 0)
               If rsTmp!Proceso > vUltimaCuota Then vUltimaCuota = rsTmp!Proceso
            rsTmp.Close
        End If
    
        Select Case True
          Case vProceso < rs!PriDeduc And vUltimaCuota < rs!PriDeduc
               curInteres = 0
               
          Case vProceso = rs!PriDeduc And vUltimaCuota = rs!PriDeduc
               curInteres = 0
               
          Case vProceso > rs!PriDeduc And vUltimaCuota > vProceso
               curInteres = 0
          
          Case vProceso = rs!PriDeduc And vUltimaCuota < vProceso 'Dias
                curInteres = (rs!Saldo * rs!interesv / 36000) * (Day(vFecha))
          
          Case (vProceso > rs!PriDeduc And vUltimaCuota = rs!PriDeduc)
                 iMeses = -1
                 Do While vProceso > vUltimaCuota
                    iMeses = iMeses + 1
                    vUltimaCuota = fxFechaProcesoSiguiente(vUltimaCuota)
                 Loop
                 curInteres = (rs!Saldo * rs!interesv / 36000) * (Day(vFecha) + (iMeses * 30))
          
          Case (vProceso > rs!PriDeduc And vProceso > vUltimaCuota)  'Idem Anterior
                 
                 iMeses = -1
                 Do While vProceso > vUltimaCuota
                    iMeses = iMeses + 1
                    vUltimaCuota = fxFechaProcesoSiguiente(vUltimaCuota)
                 Loop
                 curInteres = (rs!Saldo * rs!interesv / 36000) * (Day(vFecha) + (iMeses * 30))
          
          Case Else
               curInteres = 0
        End Select
        
        mcurIntCor = rs!MoraIntC + curInteres
    
    End If 'Plan de Pagos
    
    
    txtIntCor.Text = Format(mcurIntCor, "Standard")
    txtIntMor.Text = Format(mcurIntMor, "Standard")
    txtAmortizacion.Text = Format(mcurPrincipalMora, "Standard")
    txtPoliza.Text = Format(mcurPoliza, "Standard")
    txtCargos.Text = Format(mcurCargos, "Standard")
    txtSaldo.Text = Format(rs!Saldo, "Standard")
    
    txtNotas.Text = ""
    txtDivisa.Text = rs!cod_Divisa & ""
    
    
    'Activa boton de registro de incobrable
    tcMain.Item(0).Visible = False
    tcMain.Item(1).Visible = True

    txtCedula = rs!Cedula
    txtNombre = rs!Nombre
    
    txtOperacion.Tag = rs!opex
    txtCodigo.Text = rs!Codigo
    txtDescripcion.Text = rs!Descripcion


Else 'rs.EOF And rs.BOF

  
    mcurPrincipalMora = IIf(IsNull(rs!Principal), 0, rs!Principal)
    mcurIntCor = rs!IntCor
    mcurIntMor = rs!IntMor
    mGarantia = rs!Garantia
    mcurCargos = IIf(IsNull(rs!Cargos), 0, rs!Cargos)
    mcurPoliza = IIf(IsNull(rs!Poliza), 0, rs!Poliza)
    
    
    txtAmortizacion.Text = Format(mcurPrincipalMora, "Standard")
    txtPoliza.Text = Format(mcurPoliza, "Standard")
    txtCargos.Text = Format(mcurCargos, "Standard")
    
    txtSaldo.Text = Format(rs!Saldo, "Standard")
    txtIntCor.Text = Format(mcurIntCor, "Standard")
    txtIntMor.Text = Format(mcurIntMor, "Standard")
    
    txtNotas.Text = rs!Notas_Registro
    If IsNull(rs!TIPO_DOCUMENTO) Then
        txtRegistroDocumento.Text = "NC." & rs!genera_documento
    Else
        txtRegistroDocumento.Text = rs!TIPO_DOCUMENTO & "." & rs!Cod_Transaccion
    End If
    txtRegistroFecha.Text = rs!Registro_Fecha
    txtRegistroUsuario.Text = rs!Registro_Usuario
           
    Select Case rs!Estado
      Case "I" 'Inactivo - Condonacion de Deuda
        'Solo Boleta
        chkIncobrable.Value = vbUnchecked
        tcMain.Item(0).Visible = False
        tcMain.Item(1).Visible = False
        tcMain.Item(0).Selected = True
        
      Case "A" 'Activo
        'Activa reversion
        chkIncobrable.Value = vbChecked
        tcMain.Item(0).Visible = False
        tcMain.Item(1).Visible = True
        tcMain.Item(1).Selected = True
        
      Case "R" 'Reactivado
        
        chkIncobrable.Value = vbUnchecked
        txtReversionNotas.Text = rs!NOTAS_REVERSION
        txtReversionFecha.Text = rs!MODIFICA_FECHA
        txtReversionDocumento.Text = rs!REVERSA_DOCUMENTO
        txtReversionUsuario.Text = rs!MODIFICA_USUARIO
        txtReversionRecargo.Text = Format(rs!REACTIVACION_RECARGO, "Standard")
    
        tcMain.Item(0).Visible = True
        tcMain.Item(1).Visible = True
        tcMain.Item(0).Selected = True
    
    End Select
    
    rs.Close
    
    
    strSQL = "select R.cedula,S.nombre,R.saldo,R.proceso,R.codigo,C.descripcion,R.Opex,R.cod_Divisa" _
           & " from Socios S inner join reg_creditos R on S.cedula = R.cedula inner join Catalogo C on R.codigo = C.codigo" _
           & " Where R.id_solicitud = " & mOperacion
    Call OpenRecordSet(rs, strSQL)
    
    
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
    
    txtOperacion.Tag = rs!opex
    txtCodigo.Text = rs!Codigo
    txtDescripcion.Text = rs!Descripcion
    txtDivisa.Text = rs!cod_Divisa & ""

    'Activa Tabs
    tcMain.Item(0).Visible = True
    tcMain.Item(1).Visible = False

End If 'rs.EOF And rs.BOF
rs.Close
  
txtTotalAtrasado.Text = Format(CCur(txtSaldo.Text) + mcurCargos + mcurPoliza + mcurIntCor + mcurIntMor, "Standard")

End Sub


Private Function fxVerificar() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

On Error GoTo vError

vMensaje = ""

If Len(Trim(txtNotas)) = 0 Then
   vMensaje = vMensaje & " - Especifique una Nota para el registro..." & vbCrLf
End If


If CCur(txtTotalAtrasado.Text) = 0 Then
   vMensaje = vMensaje & " - No existe monto atrasado para procesar!..." & vbCrLf
End If

If Not IsNumeric(txtReversionRecargo.Text) Then
   vMensaje = vMensaje & " - Monto de Recargo no es válido..." & vbCrLf
End If

''Verifica que la Operacion se encuentre en Proceso Judicial / Para evitar accidentes
'strSQL = "select isnull(count(*),0) as Existe from reg_creditos where proceso in('J','C') and id_solicitud= '" & txtOperacion.Text & "'"
'Call OpenRecordSet(rs, strSQL)
'If rs!Existe = 0 Then
'   vMensaje = vMensaje & " - La Operación no se encuentra en PROCESO DE COBRO JUDICIAL para realizar la reversión..." & vbCrLf
'End If
'rs.Close


If Len(vMensaje) > 0 Then
  fxVerificar = False
  MsgBox vMensaje, vbExclamation
Else
  fxVerificar = True
End If

Exit Function

vError:
  fxVerificar = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Function



Private Sub sbDocumentoIncobrable(pTipoDoc As String, pComprobante As String, pConcepto As String, pCuenta As String)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim strCliente As String, vCuenta As String
Dim rsTmp As New ADODB.Recordset, vCuentaCargo As String, vCuentaPoliza As String
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency, curPoliza As Currency

vCuenta = pCuenta
vCuentaCargo = fxCBRParametro("23")
vAseDocDetalle = Mid(txtNotas.Text, 1, 255)
vAseDocDeposito = ""


strCliente = Trim(txtCedula) & " - " & Trim(txtNombre.Text)
strCliente = Mid(strCliente, 1, 45)

If GLOBALES.SysPlanPagos = 1 Then
    strSQL = "exec spCrdDocumentoAfectacion '" & pTipoDoc & "','" & pComprobante & "','R'"
    Call OpenRecordSet(rsTmp, strSQL, 0)
    If rsTmp.EOF And rsTmp.BOF Then
      curIntC = 0
      curIntM = 0
      curAmortiza = 0
      curCargo = 0
      curPoliza = 0
    Else
      curIntC = rsTmp!IntCor
      curIntM = rsTmp!IntMor
      curAmortiza = rsTmp!Principal
      curCargo = rsTmp!Cargos
      curPoliza = rsTmp!Polizas
    End If
    rsTmp.Close

Else
    'Sin Plan de Pagos
    curIntC = mcurIntCor
    curIntM = mcurIntMor
    curAmortiza = CCur(txtSaldo.Text)
    curCargo = mcurCargos
    curPoliza = mcurPoliza
End If


'Lineas de Comprobante
strLinea(1) = "Saldo Anterior    " & txtSaldo.Text
strLinea(2) = "Interes Corriente " & Format(curIntC, "Standard")
strLinea(3) = "Interes Atrasado  " & Format(curIntM, "Standard")
strLinea(4) = "Amortización      " & Format(curAmortiza, "Standard")
strLinea(5) = "Cargos            " & Format(curCargo, "Standard")
strLinea(6) = "Saldo Actual      " & Format(0, "Standard")
strLinea(7) = "Operacion/Línea   " & "Op.:" & txtOperacion.Text & " Lí.:" & txtCodigo.Text
strLinea(8) = ""
strLinea(9) = "REGISTRO DE INCOBRABLE"
strLinea(10) = ""
strLinea(11) = "Póliza            " & Format(curPoliza, "Standard")


'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
    strLinea(7) = "Operacion/Línea   " & "Op.:" & txtOperacion.Text & " Lí.:" & txtCodigo.Text
    strLinea(8) = "Divisa: " & rs!cod_Divisa & " / Tipo Cambio: " & rs!TipoCambio
    
    'Registro del Comprobante
    strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
             & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
             & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento)" _
             & " values('" & pComprobante & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
             & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo + curPoliza & ",'P','" & txtOperacion.Text _
             & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
             & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
             & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
             & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
             & strLinea(11) & "','" & vAseDocDetalle & "','" & vAseDocDeposito & "')"
     Call ConectionExecute(strSQL)
     
     'ASIENTO
     If curIntC > 0 Then
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & curIntC * rs!TipoCambio & ",'C','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintc _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If
     
     If curIntM > 0 Then
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & curIntM * rs!TipoCambio & ",'C','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintm _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If
     
     If curCargo > 0 Then
     'Detallar Cargos
       strSQL = "exec spCrdDocumentoAfectacionCargos '" & pTipoDoc & "','" & pComprobante & "'"
       Call OpenRecordSet(rsTmp, strSQL, 0)
       Do While Not rsTmp.EOF
             strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & rsTmp!Mov_Monto * rs!TipoCambio & ",'C','" & rs!cod_Divisa _
                    & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!cod_centro_costo & "','" & rsTmp!cod_cuenta _
                    & "','" & rsTmp!Id_Solicitud & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
             Call ConectionExecute(strSQL)
             rsTmp.MoveNext
       Loop
       rsTmp.Close
     End If
     
     If curPoliza > 0 Then
       strSQL = "select dbo.fxCrdOperacionCtaContaPolizas(" & rs!Id_Solicitud & ") as 'Cuenta'"
       Call OpenRecordSet(rsTmp, strSQL, 0)
         vCuentaPoliza = Trim(rsTmp!Cuenta)
       rsTmp.Close
       
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & curPoliza * rs!TipoCambio & ",'C','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & vCuentaPoliza _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If
     
     If curAmortiza > 0 Then
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & curAmortiza * rs!TipoCambio & ",'C','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaamortiza _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If
     
     If curIntC + curIntM + curAmortiza + curCargo + curPoliza > 0 Then
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & (curIntC + curIntM + curCargo + curAmortiza + curPoliza) * rs!TipoCambio _
              & ",'D','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & vCuenta _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If

rs.Close

End Sub


Private Sub sbDocumentoReversa(curRecargo As Currency, pTipoDoc As String, pComprobante As String _
                , vCuenta As String, vDetalle As String, vCtaRecargo As String, pConcepto As String)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim rsTmp As New ADODB.Recordset, vCuentaCargo As String, vCuentaPoliza As String
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency, curPoliza As Currency
Dim strCliente As String

vCuentaCargo = fxCBRParametro("23")
vAseDocDetalle = Mid(txtReversionNotas.Text, 1, 255)
vAseDocDeposito = ""

strCliente = Trim(txtCedula) & " - " & Trim(txtNombre.Text)
strCliente = Mid(strCliente, 1, 45)

curIntC = mcurIntCor
curIntM = mcurIntMor
curCargo = mcurCargos
curPoliza = mcurPoliza
curAmortiza = CCur(txtSaldo.Text)

strLinea(1) = "Incobrable Registrado " & Format(curAmortiza - curRecargo, "Standard")
strLinea(2) = "Recargo x Reversión   " & Format(curRecargo, "Standard")
strLinea(3) = "Total Reversión       " & Format(curAmortiza, "Standard")
strLinea(4) = "Inco.Saldo.Original   " & Format(CCur(txtSaldo.Text), "Standard")
strLinea(5) = "Inco.Interes.Original " & Format(CCur(txtIntCor.Text) + CCur(txtIntMor.Text), "Standard")
strLinea(6) = "Operación         " & txtOperacion
strLinea(7) = "Línea             " & txtCodigo.Text
strLinea(8) = "Proc.Retencion    " & "NO"
strLinea(9) = "Usuario           " & glogon.Usuario
strLinea(10) = "REVERSION DE INCOBRABLE"
strLinea(11) = "Póliza y Cargos      " & Format(curPoliza + curCargo, "Standard")


'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

    strLinea(6) = "Operacion/Línea   " & "Op.:" & txtOperacion.Text & " Lí.:" & txtCodigo.Text
    strLinea(7) = "Divisa: " & rs!cod_Divisa & " ¦ Tipo Cambio: " & rs!TipoCambio

    'Control de Documentos v2

    strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
             & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
             & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento)" _
             & " values('" & pComprobante & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
             & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curAmortiza & ",'P','" & txtOperacion.Text _
             & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
             & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
             & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
             & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
             & strLinea(11) & "','" & vAseDocDetalle & "','" & vAseDocDeposito & "')"
     Call ConectionExecute(strSQL)
     
     'ASIENTO
     If curIntC > 0 Then
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & curIntC * rs!TipoCambio & ",'D','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintc _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If
     
     If curIntM > 0 Then
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & curIntM * rs!TipoCambio & ",'D','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintm _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If
     
     If curCargo > 0 Then
     'Detallar Cargos
       strSQL = "exec spCrdDocumentoAfectacionCargos '" & pTipoDoc & "','" & pComprobante & "'"
       Call OpenRecordSet(rsTmp, strSQL, 0)
       Do While Not rsTmp.EOF
             strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & rsTmp!Mov_Monto * rs!TipoCambio & ",'D','" & rs!cod_Divisa _
                    & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!cod_centro_costo & "','" & rsTmp!cod_cuenta _
                    & "','" & rsTmp!Id_Solicitud & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
             Call ConectionExecute(strSQL)
             rsTmp.MoveNext
       Loop
       rsTmp.Close
     End If
     
     If curPoliza > 0 Then
       strSQL = "select dbo.fxCrdOperacionCtaContaPolizas(" & rs!Id_Solicitud & ") as 'Cuenta'"
       Call OpenRecordSet(rsTmp, strSQL, 0)
         vCuentaPoliza = Trim(rsTmp!Cuenta)
       rsTmp.Close
       
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & curPoliza * rs!TipoCambio & ",'D','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & vCuentaPoliza _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If
     
     If curAmortiza > 0 Then
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & curAmortiza * rs!TipoCambio & ",'D','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','','" & Trim(rs!ctaamortiza) _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If
     
     If curRecargo > 0 Then
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & curRecargo * rs!TipoCambio & ",'C','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & vCtaRecargo _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If
     
    
     If curAmortiza + curRecargo + curCargo + curIntC + curIntM > 0 Then
       strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pComprobante & "'," & (curAmortiza - curRecargo + curCargo + curIntC + curIntM) * rs!TipoCambio _
              & ",'C','" & rs!cod_Divisa _
              & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & vCuenta _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
       Call ConectionExecute(strSQL)
     End If


rs.Close

End Sub


Private Sub sbReversa()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Aplica el Registro del Incobrable
'REFERENCIAS   : fxFechaServior - (Devuelve fecha del servidor)
'                Bitacora - (Registra el Movimiento efectuado)
'-------------------------------------------------------------------------------------------

Dim strSQL As String, rs As New ADODB.Recordset, vCuenta As String
Dim vFecha As Date, vCtaCargo As String, vMonto As Currency
Dim vTipo As String, vTipoDoc As String, vDocumento As String, vConcepto As String
Dim vOTipoDoc As String, vODocumento As String

On Error GoTo vError

 If Not fxVerificar Then
    Exit Sub
 End If

Me.MousePointer = vbHourglass
    
'Selecciona la cuenta contable por Garantia de no existir aplica la cuenta por omision
strSQL = "select cod_cuenta_incobrable from crd_garantia_tipos where garantia = '" & mGarantia & "' and len(cod_cuenta_incobrable) > 6"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
    vCuenta = fxCBRParametro("20") 'Cuenta para Reversion
Else
    vCuenta = Trim(rs!cod_cuenta_incobrable)
End If
rs.Close
   
   
vCtaCargo = fxCBRParametro("21") 'Cuenta para Registro de Recargos
vConcepto = "CBR007"

'Selecciona el Tipo de Documento Aplicado Originalmente en el Incobrable
strSQL = "Select * from cbr_incobrables" _
       & " where id_solicitud = " & mOperacion & " and cod_incobrable = " & cbo.Text
Call OpenRecordSet(rs, strSQL)
 vOTipoDoc = rs!TIPO_DOCUMENTO
 vODocumento = rs!Cod_Transaccion
 If rs!Cod_Transaccion = "" Then
    vOTipoDoc = "NC"
    vODocumento = rs!genera_documento
 End If
rs.Close

''
'''Calcular Cuanto del Saldo estaba en morosidad
''strSQL = "select isnull(sum(abAmortiza),0) as 'Principal' from morosidad where estado = 'A' and id_solicitud = " & mOperacion _
''    & " and Tcon in('7','NC') and Ncon = '" & vODocumento & "'"
''Call OpenRecordSet(rs, strSQL)
''If Not rs.EOF And Not rs.BOF Then
''  curPrincipalAtrasado = rs!Principal
''End If
''rs.Close

'vMonto = CCur(txtReversionRecargo.Text) + CCur(txtIntereses.Text) + CCur(txtSaldo.Text) + CCur(txtCargos.Text) + CCur(txtPoliza.Text)
vMonto = CCur(txtReversionRecargo.Text) + CCur(txtSaldo.Text)

vFecha = fxFechaServidor

If GLOBALES.SysDocVersion = 1 Then
    vTipoDoc = "ND"
    vTipo = "8"
    vDocumento = fxDocumentoConsecutivo(vTipoDoc)
Else
    vTipoDoc = "CBJ"
    vTipo = "CBJ"
    vDocumento = fxDocumentoConsecutivo(vTipoDoc)
End If

   
''Documento
'vDocumento = fxDocumentoReversa(CCur(txtReversionRecargo.Text), CCur(txtIntereses.Text) + CCur(txtSaldo.Text), "ND" _
'            , vCuenta, "REVERSION DE INCOBRABLE", vCtaCargo)



'Registro del Maestro de Incobrables
strSQL = "update cbr_incobrables set estado = 'R',reActivacion_Recargo = " & CCur(txtReversionRecargo.Text) _
       & ",modifica_fecha = dbo.MyGetdate(), modifica_usuario = '" & glogon.Usuario & "',reversa_documento = '" & vTipoDoc & "." & vDocumento _
       & "',notas_reversion = '" & txtReversionNotas.Text & "'" _
       & " where id_solicitud = " & mOperacion & " and cod_incobrable = " & cbo.Text
Call ConectionExecute(strSQL)
  
  
If GLOBALES.SysPlanPagos = 1 Then
'    strSQL = "exec spCrdPlanPagoAnulaAbono " & lngOperacion & ",'CBR007','" & glogon.Usuario & "','" & vTipoDoc & "','" & vDocumento & "',1," & mcurIntCor _
'           & "," & mcurIntMor & "," & CCur(txtSaldo.Text) & "," & mcurCargos & "," & mcurPoliza _
'           & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
    strSQL = "exec spCrdPlanPagoAnulaAbono " & txtOperacion.Text & ",'" & vConcepto & "','" & glogon.Usuario & "','" & vTipoDoc & "','" & vDocumento & "',1," & 0 _
           & "," & 0 & "," & CCur(txtReversionRecargo.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtSaldo.Text) + CCur(txtCargos.Text) + CCur(txtPoliza.Text) & "," & 0 & "," & 0 _
           & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
    Call ConectionExecute(strSQL)
Else
    'Registro del Maestro de Credito
    strSQL = "update reg_creditos set estado = 'A', amortiza = amortiza - " & vMonto & ", saldo = saldo + " & vMonto _
           & ",saldo_mes = saldo_mes + " & vMonto & ", interesc = interesc - " & CCur(txtIntCor.Text) _
           & " where id_solicitud = " & txtOperacion.Text
    Call ConectionExecute(strSQL)
    
    'Registro Movimientos
    strSQL = "insert creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza,fechas,fechap,estado" _
           & ",tcon,ncon,cod_concepto,usuario) values('" & txtCodigo.Text & "'," & txtOperacion.Text & ",0,0," _
           & CCur(txtIntCor.Text) & "," & vMonto & ",dbo.MyGetdate()" _
           & "," & GLOBALES.glngFechaCR & ",'A','" & vTipo & "','" & vDocumento & "','" & vConcepto & "','" & glogon.Usuario & "')"
    Call ConectionExecute(strSQL)
    
    'Activacion de Morosidad Cancelada
    strSQL = "insert into morosidad(id_solicitud,codigo,fechap,fecap,estado,estadoi,IntC,IntM,Amortiza,Cargo,Fecult,Tcon,Ncon)" _
           & " (select id_solicitud,codigo,fechap,fecap,'A','C',AbIntC,AbIntM,AbAmortiza,AbCargo,dbo.MyGetdate(),'" _
           & vTipo & "','" & vDocumento & "' From morosidad" _
           & " where estado = 'C' and id_solicitud = " & mOperacion & " and Tcon in('7','NC') and Ncon = '" & vODocumento & "')"
    Call ConectionExecute(strSQL)
    
End If
    
'Registra Comprobante
Call sbDocumentoReversa(CCur(txtReversionRecargo), vTipoDoc, vDocumento, vCuenta, "REVERSION DE INCOBRABLE", vCtaCargo, vConcepto)
'Registro en Bitacora General
Call Bitacora("Reversa", "Incobrable de la Operación:" & txtOperacion & " Consec." & cbo.Text)

'Registro Historial y Expediente (07 y 08 registros y reversiones)
Call sbCBRRegTransac("08", txtCedula, txtOperacion, txtReversionNotas.Text, CCur(txtSaldo) + CCur(txtReversionRecargo.Text) _
                    , mcurIntCor, mcurIntMor, mcurCargos, mcurPoliza, mcurPrincipalMora, vTipoDoc, vDocumento)

Me.MousePointer = vbDefault

MsgBox "- Incobrable REVERSADO Satisfactoriamente", vbInformation

If vDocumento > 0 Then Call sbImprimeRecibo(vDocumento, vTipoDoc)

Call sbConsultaInicial

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub


Private Sub sbAplica()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Aplica el Registro del Incobrable
'REFERENCIAS   : fxFechaServior - (Devuelve fecha del servidor)
'                Bitacora - (Registra el Movimiento efectuado)
'-------------------------------------------------------------------------------------------

Dim strSQL As String, vCuenta As String
Dim vFecha As Date, i As Long, rs As New ADODB.Recordset
Dim vTipo As String, vTipoDoc As String, vDocumento As String, vConcepto As String

On Error GoTo vError

 If Not fxVerificar Then
    Exit Sub
 End If

Me.MousePointer = vbHourglass
    
'Selecciona la cuenta contable por Garantia de no existir aplica la cuenta por omision
strSQL = "select cod_cuenta_incobrable from crd_garantia_tipos where garantia = '" & mGarantia & "' and len(cod_cuenta_incobrable) > 6"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
    vCuenta = fxCBRParametro("19") 'Cuenta de Incobrables
Else
    vCuenta = Trim(rs!cod_cuenta_incobrable)
End If
rs.Close

'Consecutivo de Incobrable para la Operacion
i = cbo.Text
If i = 0 Then
   strSQL = "select isnull(max(cod_incobrable),0) + 1 as Consec from cbr_incobrables where id_solicitud = " & mOperacion
   Call OpenRecordSet(rs, strSQL)
    i = rs!consec
   rs.Close
Else
  Me.MousePointer = vbDefault
  MsgBox "Ya se procesó un registro de incobrable con este consecutivo para esta operación...!", vbInformation
  Exit Sub
End If


vFecha = fxFechaServidor
vConcepto = "CBR006"

If GLOBALES.SysDocVersion = 1 Then
    vTipoDoc = "NC"
    vTipo = "7"
    vDocumento = fxDocumentoConsecutivo(vTipoDoc)
Else
    vTipoDoc = "CBJ"
    vTipo = "CBJ"
    vDocumento = fxDocumentoConsecutivo(vTipoDoc)
End If

'Registro del Maestro de Incobrables
strSQL = "insert cbr_incobrables(cod_incobrable,id_solicitud,registro_usuario,registro_fecha,saldo,intCor,intMor" _
       & ",Estado,Notas_registro,Genera_documento,cargos,poliza,principal,tipo_documento,cod_transaccion)" _
       & " values(" & i & "," & txtOperacion.Text & ",'" & glogon.Usuario & "',dbo.MyGetdate()," & CCur(txtSaldo.Text) & "," & mcurIntCor _
       & "," & mcurIntMor & ",'" & IIf((chkIncobrable.Value = vbChecked), "A", "I") & "','" & txtNotas.Text & "','" _
       & vTipoDoc & "." & vDocumento & "'," & mcurCargos & "," & mcurPoliza & "," & mcurPrincipalMora & ",'" & vTipoDoc & "','" & vDocumento & "')"
Call ConectionExecute(strSQL)

If GLOBALES.SysPlanPagos = 1 Then
    'Cancelación en Plan de Pagos
'    strSQL = "exec spCrdPlanPagoAbonoCancelacion " & txtOperacion.Text & ",'" & vConcepto & "','" & glogon.Usuario & "','" & vTipoDoc _
'           & "','" & vDocumento & "'," & CCur(txtTotalAtrasado.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
'    Call ConectionExecute(strSQL)
    
     If chkSaldos.Value = xtpChecked Then
     strSQL = "exec spCrdPlanPagoAbonoEC " & txtOperacion.Text & ",'" & vConcepto & "','" & glogon.Usuario & "','" & vTipoDoc _
            & "','" & vDocumento & "',0,0," & CCur(txtSaldo.Text) _
            & ",0,'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "','',1,1,1"
     Else
     strSQL = "exec spCrdPlanPagoAbonoEC " & txtOperacion.Text & ",'" & vConcepto & "','" & glogon.Usuario & "','" & vTipoDoc _
            & "','" & vDocumento & "',0," & CCur(txtIntCor.Text) & "," & CCur(txtSaldo.Text) _
            & ",0,'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "','',1,1,1"
     End If
     
     strSQL = strSQL & Space(10) & "exec spCrdPlanPagos " & txtOperacion.Text
     Call ConectionExecute(strSQL)
    
Else
    'Registro del Maestro de Credito
    strSQL = "update reg_creditos set estado = 'C', amortiza = amortiza + saldo, saldo = 0,interesc = interesc + " & mcurIntCor + mcurIntMor _
           & " where id_solicitud = " & txtOperacion.Text
    Call ConectionExecute(strSQL)
    
    'Registro Movimientos
    strSQL = "update morosidad set estado = 'C', AbIntc = Intc, AbIntm = Intm, AbAmortiza = Amortiza, abCargo = Cargo" _
           & ",Tcon = '" & vTipo & "',Ncon = '" & vDocumento & "',FecUlt = dbo.MyGetdate(), cod_concepto = '" & vConcepto & "', Usuario = '" & glogon.Usuario _
           & "' where id_solicitud = " & txtOperacion.Text & " and estado = 'A'"
    Call ConectionExecute(strSQL)
      
    strSQL = "insert creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza,fechas,fechap,estado" _
           & ",tcon,ncon,cod_concepto,usuario,cod_caja) values('" & txtCodigo.Text & "'," & txtOperacion.Text & ",0,0,0," _
           & CCur(txtSaldo.Text) - CCur(txtAmortizacion.Text) & ",dbo.MyGetdate()" _
           & "," & GLOBALES.glngFechaCR & ",'A','" & vTipo & "','" & vDocumento & "','" & vConcepto & "','" & glogon.Usuario & "','')"
    Call ConectionExecute(strSQL)
End If 'Plan de Pagos
    
'Registra el Comprobante
Call sbDocumentoIncobrable(vTipoDoc, vDocumento, vConcepto, vCuenta)
    
'Registro en Bitacora General
Call Bitacora("Aplica", "Registro de Incobrable de la Operación:" & txtOperacion & " Consec: " & i)

'Registro Historial y Expediente (07 y 08 registros y reversiones)
Call sbCBRRegTransac("07", txtCedula, txtOperacion, txtNotas.Text, CCur(txtSaldo), mcurIntCor, mcurIntMor, mcurCargos, mcurPoliza, mcurPrincipalMora, vTipoDoc, vDocumento)

Me.MousePointer = vbDefault

MsgBox "- Incobrable Registrado Satisfactoriamente", vbInformation

If vDocumento > 0 Then Call sbImprimeRecibo(vDocumento, vTipoDoc)

Call sbConsultaInicial

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0
Call sbConsultaInicial

End Sub

Private Sub txtReversionRecargo_GotFocus()
On Error GoTo vError
txtReversionRecargo.Text = CCur(txtReversionRecargo.Text)
vError:
End Sub

Private Sub txtReversionRecargo_LostFocus()
On Error GoTo vError
txtReversionRecargo.Text = Format(CCur(txtReversionRecargo.Text), "Standard")
vError:
End Sub
