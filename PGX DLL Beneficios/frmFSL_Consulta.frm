VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmFSL_Consulta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Consulta de Casos en Trámite"
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17340
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   17340
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   195
      Left            =   2880
      TabIndex        =   43
      Top             =   7440
      Width           =   195
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ListView lswEnfermedades 
      Height          =   2055
      Left            =   120
      TabIndex        =   41
      Top             =   4320
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   3625
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
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin VB.Frame fraFiltros 
      Caption         =   "Filtros adicionales"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   3480
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   7455
      Begin VB.ComboBox cboEstadoPersona 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox cboDesembolso 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox cboApelacion 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3720
         Width           =   4695
      End
      Begin VB.ComboBox cboGestiones 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   4200
         Width           =   4695
      End
      Begin VB.ComboBox cboComite 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2280
         Width           =   5295
      End
      Begin VB.ComboBox cboMiembros 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2760
         Width           =   3735
      End
      Begin MSComctlLib.Toolbar tlbFiltros 
         Height          =   330
         Left            =   6000
         TabIndex        =   6
         Top             =   5160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   1561
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "Cerrar"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   330
         Left            =   1560
         TabIndex        =   47
         Top             =   360
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9340
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtExpediente 
         Height          =   330
         Left            =   1560
         TabIndex        =   48
         Top             =   1080
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuarioRegistra 
         Height          =   330
         Left            =   4200
         TabIndex        =   49
         Top             =   1080
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Persona"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   21
         Left            =   4200
         TabIndex        =   34
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   37
         Left            =   4200
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Expediente"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   35
         Left            =   1560
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Desembolso"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   33
         Left            =   1560
         TabIndex        =   12
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7200
         X2              =   120
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apelación Registrada"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestiones Registradas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Comité"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Resuelto por miembro?"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Index           =   12
         Left            =   1800
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
   End
   Begin VB.Frame fraReportes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   7800
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   5895
      Begin VB.OptionButton optReportes 
         Appearance      =   0  'Flat
         Caption         =   "Agrupado por Comité"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   19
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton optReportes 
         Appearance      =   0  'Flat
         Caption         =   "Agrupado por Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   18
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optReportes 
         Appearance      =   0  'Flat
         Caption         =   "Listado General"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.ComboBox cboTipoReporte 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComctlLib.Toolbar tlbReportes 
         Height          =   330
         Left            =   3240
         TabIndex        =   20
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         ButtonWidth     =   1799
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "Imprimir"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "Cerrar"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes ...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   28
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo ...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   27
         Left            =   600
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   6000
         X2              =   120
         Y1              =   2760
         Y2              =   2760
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   9615
      Width           =   17340
      _ExtentX        =   30586
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Casos Encontrados..:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Total Registrado..:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   330
      Left            =   9120
      TabIndex        =   24
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      ButtonWidth     =   1879
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte"
            Key             =   "Reporte"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exportar"
            Key             =   "Exportar"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Excel"
                  Text            =   "Microsoft Excel"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "HTML"
                  Text            =   "HTML"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6375
      Left            =   3480
      TabIndex        =   25
      Top             =   960
      Width           =   12615
      _Version        =   524288
      _ExtentX        =   22251
      _ExtentY        =   11245
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
      MaxCols         =   20
      SpreadDesigner  =   "frmFSL_Consulta.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14040
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Consulta.frx":158D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Consulta.frx":7DEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Consulta.frx":E651
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Consulta.frx":14EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Consulta.frx":14FEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   360
      Left            =   1200
      TabIndex        =   35
      Top             =   7680
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   635
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   360
      Left            =   1200
      TabIndex        =   36
      Top             =   8040
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   635
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
   Begin XtremeSuiteControls.ComboBox cboPlan 
      Height          =   330
      Left            =   1200
      TabIndex        =   37
      Top             =   6840
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   1200
      TabIndex        =   38
      Top             =   6480
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   3480
      TabIndex        =   39
      Top             =   480
      Width           =   3375
      _Version        =   1441793
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.FlatEdit txtBuscarPor 
      Height          =   330
      Left            =   6960
      TabIndex        =   40
      Top             =   480
      Width           =   6615
      _Version        =   1441793
      _ExtentX        =   11668
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ListView lswCausas 
      Height          =   3375
      Left            =   120
      TabIndex        =   42
      Top             =   480
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
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
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkCausas 
      Height          =   195
      Left            =   2880
      TabIndex        =   44
      Top             =   120
      Width           =   195
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkEnfermedades 
      Height          =   195
      Left            =   2880
      TabIndex        =   45
      Top             =   3960
      Width           =   195
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFiltros 
      Height          =   195
      Left            =   5760
      TabIndex        =   46
      Top             =   120
      Width           =   1035
      _Version        =   1441793
      _ExtentX        =   1826
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "+ Filtros"
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
      Appearance      =   17
      Alignment       =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   31
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enfermedades...:"
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
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   30
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Causas ...:"
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
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
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
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   28
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   27
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por ...:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   3480
      TabIndex        =   26
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image imgBanner 
      Height          =   9396
      Left            =   0
      Picture         =   "frmFSL_Consulta.frx":150EA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3200
   End
End
Attribute VB_Name = "frmFSL_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String
Dim vTipoDocu As String
Dim vScroll As Boolean
Dim strSQLReporte As String
Dim strDetalle As String
Dim vPaso As Boolean

Private Sub cboComite_Click()
Dim strSQL As String

cboMiembros.Clear
cboMiembros.AddItem "TODOS"
cboMiembros.Text = "TODOS"

If vPaso Then Exit Sub
If cboComite.ListCount = 0 Then Exit Sub
   

If cboComite.Text <> "TODOS" Then

    strSQL = "select rtrim(cedula) + ' - ' + rtrim(Nombre) as ItmX from FSL_COMITES_MIEMBROS" _
           & " where cod_comite = '" & SIFGlobal.fxCodText(cboComite.Text) & "' and Activo = 1 order by Nombre"
    
    Call sbLlenaCbo(cboMiembros, strSQL, True, False)
    
End If


End Sub

Private Sub cboPlan_Click()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListItem

If vPaso Then Exit Sub
If cboPlan.ListCount = 0 Then Exit Sub
   
chkCausas.Value = vbChecked
lswCausas.ListItems.Clear

If cboPlan.Text <> "TODOS" Then
   
   lswCausas.ListItems.Clear
    strSQL = "select cod_causa as IdX, rtrim(descripcion) as ItmX from FSL_PLANES_CAUSAS" _
           & " where cod_plan = '" & cboPlan.ItemData(cboPlan.ListIndex) & "' and Activa = 1 order by Descripcion"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     Set itmX = lswCausas.ListItems.Add(, , rs!itmX)
         itmX.Tag = rs!IdX
         itmX.Checked = chkCausas.Value
     rs.MoveNext
    Loop
    rs.Close
End If

End Sub



Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
   dtpCorte.Enabled = False
   dtpInicio.Enabled = False
Else
   dtpCorte.Enabled = True
   dtpInicio.Enabled = True
End If
End Sub

Private Sub chkFiltros_Click()
If chkFiltros.Value = vbChecked Then
    fraFiltros.Visible = True
    fraFiltros.Top = 480
    fraFiltros.Left = 3480
Else
    fraFiltros.Visible = False
End If
End Sub

Private Sub chkEnfermedades_Click()
Dim i As Integer

For i = 1 To lswEnfermedades.ListItems.Count
  lswEnfermedades.ListItems.Item(i).Checked = chkEnfermedades.Value
Next i
End Sub

Private Sub chkCausas_Click()
Dim i As Integer

For i = 1 To lswCausas.ListItems.Count
  lswCausas.ListItems.Item(i).Checked = chkCausas.Value
Next i

End Sub

Private Sub sbBuscar()
Dim strSQL As String, i As Integer
Dim vCadena As String, iCantidad As Integer
Dim vEstado As String

On Error GoTo vError

Me.MousePointer = vbHourglass
iCantidad = 0

Select Case Mid(cboEstado.Text, 1, 3)
 Case "Pen"
   vEstado = "'P'"
 Case "Apr"
   vEstado = "'A'"
 Case "Apl"
   vEstado = "'X'"
 Case "Rec"
   vEstado = "'R'"
 Case Else
       vEstado = "'A','R','P','X','Y'"
 
End Select
    
strSQL = "select '','','',COD_EXPEDIENTE,CEDULA,NOMBRE,EDAD,REGISTRO_USUARIO,REGISTRO_FECHA,ESTADO_DESC" _
       & ", PLAN_DESC,CAUSA_DESC,ENFERMEDAD_DESC,COMITE_DESC,RESOLUCION_FECHA" _
       & ",TOTAL_DISPONIBLE, TOTAL_APLICADO,TOTAL_SOBRANTE, PRESENTA_CEDULA, PRESENTA_NOMBRE" _
       & " from vFSL_CasosLista" _
       & " Where Estado in(" & vEstado & ")"

If Mid(cboEstado.Text, 1, 3) = "Ape" Then
    strSQL = strSQL & " and RESOLUCION_ESTADO = 'Y'"
End If

If cboPlan.Text <> "TODOS" Then
    strSQL = strSQL & " and cod_Plan = '" & cboPlan.ItemData(cboPlan.ListIndex) & "'"

    'Lista de Causas
    If chkCausas.Value = vbChecked Then
        vCadena = " cod_causa in('"
        For i = 1 To lswCausas.ListItems.Count
          If lswCausas.ListItems.Item(i).Checked Then
            vCadena = vCadena & "','" & lswCausas.ListItems.Item(i).Tag
            iCantidad = iCantidad + 1
          End If
        Next i
        
        If iCantidad > 0 Then strSQL = strSQL & " and " & vCadena & "')"
        
    End If

End If

    
iCantidad = 0
''Lista de Enfermedades
If chkEnfermedades.Value = vbChecked Then
    vCadena = " cod_enfermedad in('"
    For i = 1 To lswEnfermedades.ListItems.Count
      If lswEnfermedades.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswEnfermedades.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    

        strSQL = strSQL & " and " & vCadena & "')"
End If
    
'Validación De las Fechas
If chkFechas.Value = vbUnchecked Then
   Select Case vEstado
      Case "A", "R"
        strSQL = strSQL & " and Resolucion_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                        & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
      
      Case Else
        strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                        & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
   End Select
 
End If
    
    

'Filtros Adicionales
    If Trim(cbo.Text) <> "" Then
        Select Case Mid(cbo.Text, 1, 2)
            Case "01" 'Cedula
              strSQL = strSQL & " and cedula like '%" & Trim(txtBuscarPor.Text) & "%'"
            Case "02" 'Presenta - Cedula
              strSQL = strSQL & " and Presenta_Cedula like '%" & Trim(txtBuscarPor.Text) & "%'"
            Case "03" 'Presenta - Nombre
              strSQL = strSQL & " and Presenta_Nombre like '%" & Trim(txtBuscarPor.Text) & "%'"
        End Select
    End If

   
  'Filtros a Nivel General
  If Trim(txtNombre.Text) <> "" Then
      strSQL = strSQL & " and Nombre like '%" & Trim(txtNombre.Text) & "%'"
  End If
    
  If cboDesembolso.Text <> "TODOS" Then
      strSQL = strSQL & " and Tipo_Desembolso = '" & Mid(cboDesembolso.Text, 1, 1) & "'"
  End If
  
  If cboEstadoPersona.Text <> "TODOS" Then
      strSQL = strSQL & " and EstadoActual = '" & SIFGlobal.fxCodText(cboEstadoPersona.Text) & "'"
  End If
  
  If cboComite.Text <> "TODOS" Then
      strSQL = strSQL & " and COD_COMITE = '" & SIFGlobal.fxCodText(cboComite.Text) & "'"
      
      
      If cboMiembros.Text <> "TODOS" Then
          strSQL = strSQL & " and dbo.fxFSL_Expediente_ComiteMiembro(Cod_Expediente,'" & SIFGlobal.fxCodText(cboMiembros.Text) & "') >= 1"
      End If
      
  End If
  

  
  'Filtros a nivel de Tramite
  If Trim(txtExpediente.Text) <> "" Then
      strSQL = strSQL & " and cod_expediente like '%" & Trim(txtExpediente.Text) & "%'"
  End If
  
  If Trim(txtUsuarioRegistra.Text) <> "" Then
      strSQL = strSQL & " and Registro_Usuario like '%" & Trim(txtUsuarioRegistra.Text) & "%'"
  End If
  
  If cboGestiones.Text <> "TODOS" Then
      strSQL = strSQL & " and dbo.fxFSL_Expediente_GestionRegistrada(cod_Expediente,'" & SIFGlobal.fxCodText(cboGestiones.Text) & "') >= 1"
  End If

  If cboApelacion.Text <> "TODOS" Then
      strSQL = strSQL & " and dbo.fxFSL_Expediente_ApelacionRegistrada(cod_Expediente,'" & SIFGlobal.fxCodText(cboApelacion.Text) & "') >= 1"
  End If
  

  
    

Call sbCargaGridLocal(vGrid, 20, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim curMonto As Currency

On Error GoTo vError

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.Col = i
 vGrid.Text = ""
Next i

curMonto = 0

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i

    If rs.Fields(i - 1).Type = 135 Then
        If Year(rs.Fields(i - 1).Value) > 1900 Then
           vGrid.Text = Format((rs.Fields(i - 1).Value & ""), "dd/mm/yyyy")
        End If
    Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
    End If

  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  curMonto = curMonto + rs!Total_Disponible
  rs.MoveNext
Loop
rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

StatusBarX.Panels(1).Text = "Total Registros " & vGrid.MaxRows
StatusBarX.Panels(2).Text = "Monto ..: " & Format(curMonto, "Standard")

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxFechaReportes(vTipo As Integer) As String

fxFechaReportes = " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")" _
                & " to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

End Function

Private Function fxUsuarioNombre(vUsuario As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select descripcion from usuarios where nombre = '" & vUsuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
 fxUsuarioNombre = "[SIN DESCRIPCION]"
Else
 fxUsuarioNombre = "[" & UCase(Trim(rs!Descripcion)) & "]"

End If
rs.Close
End Function

Private Sub Form_Activate()
 vModulo = 7
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vModulo = 7


With lswEnfermedades.ColumnHeaders
    .Clear
    .Add , , "", lswEnfermedades.Width - 100
End With

With lswCausas.ColumnHeaders
    .Clear
    .Add , , "", lswCausas.Width - 100
End With


Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.AppearanceStyle = fxGridStyle


cbo.Clear
cbo.AddItem "01 - Cédula"
cbo.AddItem "02 - Presenta - Cédula"
cbo.AddItem "03 - Presenta - Nombre"
cbo.Text = "01 - Cédula"

cboTipoReporte.Clear
cboTipoReporte.AddItem "Detalle"
cboTipoReporte.AddItem "Resumen"
cboTipoReporte.Text = "Detalle"

cboEstado.Clear
cboEstado.AddItem "Pendiente"
cboEstado.AddItem "Aprobado"
cboEstado.AddItem "Rechazado"
cboEstado.AddItem "Aplicado"
cboEstado.AddItem "Apelación Pendiente"
cboEstado.AddItem "TODOS"
cboEstado.Text = "TODOS"


cboDesembolso.Clear
cboDesembolso.AddItem "Fondos"
cboDesembolso.AddItem "Tesorería"
cboDesembolso.AddItem "TODOS"
cboDesembolso.Text = "TODOS"


dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -30, dtpCorte.Value)

vGrid.MaxRows = 0

vPaso = True

strSQL = "select COD_PLAN as  'IdX', rtrim(Descripcion) as 'itmx' from FSL_PLANES  where activo = 1 order by descripcion"
Call sbCbo_Llena_New(cboPlan, strSQL, True, True)

vPaso = False
Call cboPlan_Click



lswEnfermedades.ListItems.Clear
strSQL = "select cod_enfermedad,DESCRIPCION from FSL_TIPOS_ENFERMEDADES where Activa = 1 order by descripcion"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswEnfermedades.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!cod_enfermedad
     itmX.Checked = chkEnfermedades.Value
 rs.MoveNext
Loop
rs.Close


vPaso = True
strSQL = "select RTRIM(cod_Comite) + ' - ' + DESCRIPCION AS 'ItmX' from FSL_COMITES" _
       & " Where ACTIVO = 1 ORDER BY COD_COMITE"
Call sbLlenaCbo(cboComite, strSQL, True)
vPaso = False
Call cboComite_Click

strSQL = " select RTRIM(COD_ESTADO ) + ' - ' + DESCRIPCION AS 'ItmX' from AFI_ESTADOS_PERSONA   where Activo = 1 ORDER BY COD_ESTADO "
Call sbLlenaCbo(cboEstadoPersona, strSQL, True)


strSQL = "select RTRIM(COD_GESTION) + ' - ' + DESCRIPCION AS 'ItmX' from FSL_TIPOS_GESTIONES where Activa = 1 order by COD_GESTION"
Call sbLlenaCbo(cboGestiones, strSQL, True)

strSQL = " select RTRIM(COD_APELACION ) + ' - ' + DESCRIPCION AS 'ItmX' from FSL_TIPOS_APELACIONES  where Activa = 1 ORDER BY COD_APELACION "
Call sbLlenaCbo(cboApelacion, strSQL, True)


Me.Height = 8565
Me.Width = 14070
End Sub

Private Sub Form_Resize()
On Error Resume Next


vGrid.Width = Me.Width - (vGrid.Left + 350)
vGrid.Height = Me.Height - 2100
lswEnfermedades.Height = Me.Height - 7200


Label1(6).Top = lswEnfermedades.Top + lswEnfermedades.Height + 180
cboEstado.Top = Label1(6).Top
Label1(1).Top = Label1(6).Top + Label1(6).Height + 60
cboPlan.Top = Label1(1).Top

Label1(4).Top = cboPlan.Top + cboPlan.Height + 60
dtpInicio.Top = Label1(4).Top
Label1(5).Top = dtpInicio.Top + dtpInicio.Height + 60
dtpCorte.Top = Label1(5).Top
chkFechas.Top = dtpCorte.Top + dtpInicio.Height + 60

imgBanner.Height = Me.Height

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
    vPaso = True
        Call sbBuscar
    vPaso = False

  Case "Reporte"
    fraReportes.Top = tlb.Top + 390
    fraReportes.Left = tlb.Left - 1300
    fraReportes.Visible = True
End Select

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim vHeaders As vGridHeaders, i As Integer

    vHeaders.Columnas = 20
    vHeaders.Headers(1) = ""
    vHeaders.Headers(2) = ""
    vHeaders.Headers(3) = ""
    vHeaders.Headers(4) = "No. Expediente"
    vHeaders.Headers(5) = "Cedula"
    vHeaders.Headers(6) = "Nombre"
    vHeaders.Headers(7) = "Edad"
    vHeaders.Headers(8) = "Usuario"
    vHeaders.Headers(9) = "Fecha"
    vHeaders.Headers(10) = "Estado"
    vHeaders.Headers(11) = "Plan"
    vHeaders.Headers(12) = "Causa"
    vHeaders.Headers(13) = "Enfermedad"
    vHeaders.Headers(14) = "Comité"
    vHeaders.Headers(15) = "Fec.Resolución"
    vHeaders.Headers(16) = "Total.Fosol"
    vHeaders.Headers(17) = "Total.Aplicado"
    vHeaders.Headers(18) = "Total.Sobrante"
    vHeaders.Headers(19) = "Presenta.Ced."
    vHeaders.Headers(20) = "Presenta.Nom."
    
    
Select Case ButtonMenu.Key
  Case "Excel"
      Call sbSIFGridExportar(vGrid, vHeaders, "FOSOL_Consulta")
  Case "HTML"
      Call sbSIFGridExportar(vGrid, vHeaders, "FOSOL_Consulta", "HTML")
End Select

End Sub


Private Function fxReportesFiltros() As String
Dim vFiltro As String
Dim vCadena As String, iCantidad As Integer, i As Integer


On Error GoTo vError

iCantidad = 0
vFiltro = ""
strDetalle = ""

If cbo.Text <> "" Then
    Select Case Mid(cbo, 1, 2)
        Case "01"
            vFiltro = vFiltro & "{vSIFDocumentos.cliente_Identificacion}  like '*" & txtBuscarPor.Text & "*' "
            strDetalle = "Id. cliente .: " & txtBuscarPor.Text
        Case "02"

            vFiltro = vFiltro & "{vSIFDocumentos.Cliente_Nombre}  like '*" & txtBuscarPor.Text & "*' "
            strDetalle = "Nombre cliente .: " & txtBuscarPor.Text
        Case "03"

            vFiltro = vFiltro & "{vSIFDocumentos.Cod_Transaccion} = '" & txtBuscarPor.Text & "' "
            strDetalle = "Transacción .: " & txtBuscarPor.Text
        Case "04"

             vFiltro = vFiltro & "{vSIFDocumentos.Documento} = '" & txtBuscarPor.Text & "' "
             strDetalle = "Documento .: " & txtBuscarPor.Text
        Case "05"

            vFiltro = vFiltro & "{vSIFDocumentos.Registro_Usuario}  like '" & txtBuscarPor.Text & "' "
            strDetalle = "Usuario .: " & txtBuscarPor.Text
    End Select
End If 'xx

'Lista de Documentos
vCadena = ""
For i = 1 To lswCausas.ListItems.Count
  If lswCausas.ListItems.Item(i).Checked Then
    If vCadena <> "" Then
        vCadena = vCadena & ","
    End If
    vCadena = vCadena & "'" & lswCausas.ListItems.Item(i).Tag & "'"
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad > 2 Then
  strDetalle = strDetalle & " - Doc .: Filtrados"
ElseIf iCantidad = 0 Then
  strDetalle = strDetalle & " - Doc .: Todos"
Else
   strDetalle = strDetalle & " - Doc.: " & Mid(vCadena, 28, Len(vCadena))
End If

iCantidad = 0
  If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
  vFiltro = vFiltro & "{vSIFDocumentos.Tipo_Documento} in [" & vCadena & "] "

vCadena = ""
For i = 1 To lswEnfermedades.ListItems.Count
  If lswEnfermedades.ListItems.Item(i).Checked Then
    If vCadena <> "" Then
        vCadena = vCadena & ","
    End If
    vCadena = vCadena & "'" & lswEnfermedades.ListItems.Item(i).Tag & "'"
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad > 2 Then
  strDetalle = strDetalle & " - Conceptos .: Filtrados"
ElseIf iCantidad = 0 Then
  strDetalle = strDetalle & " - Conceptos .: Todos"
Else
   strDetalle = strDetalle & " - Concepto.:" & Mid(vCadena, 28, Len(vCadena))
End If

If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
 vFiltro = vFiltro & "{vSIFDocumentos.Cod_Concepto} in [" & vCadena & "] "

If chkFechas.Value = vbUnchecked Then
    Select Case cboEstado.Text
      Case "Registro"
    
        If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
        vFiltro = vFiltro & "cdate({vSIFDocumentos.registro_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
        vFiltro = vFiltro & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    
        strDetalle = strDetalle & " - Fecha Registro.: desde " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpCorte.Value, "dd/mm/yyyy")
    
      Case "Anulación"
    
        If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
        vFiltro = vFiltro & "cdate({vSIFDocumentos.anulacion_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
        vFiltro = vFiltro & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    
        strDetalle = strDetalle & " - Fecha Anulación.: desde " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpCorte.Value, "dd/mm/yyyy")
    
      Case "Traslado"
        If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
        vFiltro = vFiltro & "cdate({vSIFDocumentos.traspaso_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
        vFiltro = vFiltro & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    
        strDetalle = strDetalle & " - Fecha Traslado.: desde " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case Else
        strDetalle = strDetalle & " - Todas las Fechas"
    End Select
End If

Select Case cboEstado.Text
  Case "Impreso"
     If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
     vFiltro = vFiltro & "{vSIFDocumentos.estado}  in ['I','E'] "

  Case "Pendiente"
     If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
     vFiltro = vFiltro & "{vSIFDocumentos.estado}  = 'P' "
  Case Else

End Select
strDetalle = strDetalle & " - Estado..:" & cboEstado.Text


fxReportesFiltros = vFiltro


Exit Function

vError:
  fxReportesFiltros = ""
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical




End Function

Private Sub tlbFiltros_ButtonClick(ByVal Button As MSComctlLib.Button)
fraFiltros.Visible = False
chkFiltros.Value = vbUnchecked
End Sub

Private Sub tlbReportes_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFiltros As String, vFiltrosEtiquetas As String

MsgBox "Opción Prevista para Posible desarrollo - Consulte a su administrador!", vbExclamation

With frmContenedor.Crt
'    .Reset
'    .WindowTitle = "Reportes del Módulo: Control de Documentos"
'    .WindowState = crptMaximized
'    .WindowShowGroupTree = True
'
'    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'    '.Formulas(2) = "Detalle = '" & strDetalle & "'"
'    .Formulas(2) = "Usuario = 'Usuario..:" & glogon.Usuario & "'"
'    .Formulas(3) = "Fecha = 'Fecha ...:" & fxFechaServidor & "'"
'    .Connect = glogon.ConectRPT
'
'
'    vFiltros = fxReportesFiltros


Select Case Button.Key
    Case "Imprimir"
'        Select Case True
'            Case optReportes.Item(0).Value 'Reporte General
''                If Mid(cboTipoReporte, 1, 1) = "D" Then
'                   .ReportFileName = SIFGlobal.fxPathReportes("Cbr_CJTramiteDetalle.rpt")
''                Else
''                   .ReportFileName = SIFGlobal.fxPathReportes("SIFDocGeneralRsm.rpt")
''                End If
'            Case optReportes.Item(1).Value 'Agrupado por Usuario
'                If Mid(cboTipoReporte, 1, 1) = "D" Then
'                    .ReportFileName = SIFGlobal.fxPathReportes("Cbr_CJTramiteDetallexUsuario.rpt")
'                Else
'                    .ReportFileName = SIFGlobal.fxPathReportes("Cbr_CJTramiteResumenxUsuario.rpt")
'                End If
'
'            Case optReportes.Item(3).Value 'Agrupado por Oficina
'                If Mid(cboTipoReporte, 1, 1) = "D" Then
'                    .ReportFileName = SIFGlobal.fxPathReportes("Cbr_CJTramiteDetallexOficina.rpt")
'                Else
'                    .ReportFileName = SIFGlobal.fxPathReportes("Cbr_CJTramiteResumenxOficina.rpt")
'                End If
'
'        End Select
''
''        'Si son los reportes Especiales no aplicar filtros default
''        If Not (optReportes.Item(5).Value Or optReportes.Item(6).Value) Then
''            .SelectionFormula = vFiltros
''        End If
'        .PrintReport
''.Action = 1

 Case "Cerrar"
      fraReportes.Visible = False


End Select

End With

End Sub








Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim frm As Form
Dim pExpediente As String


If vPaso Then Exit Sub


vGrid.Row = Row
vGrid.Col = 4
pExpediente = vGrid.Text

Select Case Col
  Case 1 'Expediente
    Call sbFormsCall("frmFSL_Expediente")
    
    For Each frm In Forms
        If (UCase(frm.Name) = UCase("frmFSL_Expediente")) Then
            Call frm.sbConsultaExterna(pExpediente)
            Exit For
        End If
    Next frm
  
  Case 2 'Gestiones
    GLOBALES.gTag = pExpediente
    Call sbFormsCall("frmFSL_ExpedienteGestiones")
    
  Case 3 'Apelaciones
    GLOBALES.gTag = pExpediente
    Call sbFormsCall("frmFSL_ExpedienteApelaciones")
End Select


End Sub

Private Function fxFiltro(vTipo As String) As String
Dim strSQL As String

If vTipo = "T" Then
    strSQL = "select '',T.COD_TRAMITE,R.ID_SOLICITUD,R.CODIGO" _
            & ",R.CEDULA,S.NOMBRE,dbo.fxCrdPlazoRestante(R.PLAZO,R.PRIDEDUC," & GLOBALES.glngFechaCR & " )as 'Plazo'" _
            & ",isnull(T.TOTAL_DEUDA, R.Saldo) as 'Monto' ,T.PROCESO_USUARIO,T.PROCESO_FECHA" _
            & ",R.ProcesoDesc,A.NOMBRE as 'Abogado'" _
            & ",J.NOMBRE as 'Juzgado',Tj.DESCRIPCION as 'Juicio'" _
            & " from vCbrCjOperaciones R left JOIN CBR_CJ_TRAMITE  T ON R.ID_SOLICITUD = T.ID_SOLICITUD" _
            & " left JOIN SOCIOS S ON R.CEDULA = S.CEDULA" _
            & " left join CBR_CJ_ABOGADOS A On T.COD_ABOGADO = A.COD_ABOGADO" _
            & " left join FSL_PLANES_CAUSAS J on T.COD_JUZGADO = J.COD_JUZGADO" _
            & " left join FSL_TIPOS_ENFERMEDADES Tj on T.cod_enfermedad = Tj.cod_enfermedad"
Else

    strSQL = "select '','',R.ID_SOLICITUD,R.CODIGO" _
            & ",R.CEDULA,S.NOMBRE,dbo.fxCrdPlazoRestante(R.PLAZO,R.PRIDEDUC," & GLOBALES.glngFechaCR & " )as 'Plazo'" _
            & ",R.Saldo as 'Monto' ,'','',R.ProcesoDesc,'','',''" _
            & " from vCbrCjOperacionesPendientes R inner JOIN SOCIOS S ON R.CEDULA = S.CEDULA"
            
End If
        
fxFiltro = strSQL

End Function


