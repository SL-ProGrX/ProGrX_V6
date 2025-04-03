VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmActivos_WizardOP 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asistente para Detallar Activos derivados de Obras en Proceso"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "&Siguiente >>"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1452
   End
   Begin VB.CommandButton cmdAtras 
      Caption         =   "<< &Atras"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1452
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   3612
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6732
      _ExtentX        =   11880
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tipo de Movimiento"
      TabPicture(0)   =   "frmActivos_WizardOP.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "opt(1)"
      Tab(0).Control(1)=   "opt(0)"
      Tab(0).Control(2)=   "Label5"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Activo D01"
      TabPicture(1)   =   "frmActivos_WizardOP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtNotas"
      Tab(1).Control(1)=   "txtCodigo"
      Tab(1).Control(2)=   "txtDescripcion"
      Tab(1).Control(3)=   "txtDocCompra"
      Tab(1).Control(4)=   "txtProveedor"
      Tab(1).Control(5)=   "cboTipo"
      Tab(1).Control(6)=   "Label20"
      Tab(1).Control(7)=   "Label14(1)"
      Tab(1).Control(8)=   "Label13"
      Tab(1).Control(9)=   "Label21"
      Tab(1).Control(10)=   "Label19"
      Tab(1).Control(11)=   "Label1(1)"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Activo D02"
      TabPicture(2)   =   "frmActivos_WizardOP.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboVU"
      Tab(2).Control(1)=   "txtVU"
      Tab(2).Control(2)=   "cbo"
      Tab(2).Control(3)=   "txtValorRescate"
      Tab(2).Control(4)=   "txtValorHistorico"
      Tab(2).Control(5)=   "txtUDProducidas"
      Tab(2).Control(6)=   "txtUDAnio"
      Tab(2).Control(7)=   "txtModelo"
      Tab(2).Control(8)=   "txtSerie"
      Tab(2).Control(9)=   "txtMarca"
      Tab(2).Control(10)=   "txtOtrasSenas"
      Tab(2).Control(11)=   "dtpAdquisicion"
      Tab(2).Control(12)=   "dtpInstalacion"
      Tab(2).Control(13)=   "Line2"
      Tab(2).Control(14)=   "Line1"
      Tab(2).Control(15)=   "Label4(0)"
      Tab(2).Control(16)=   "Label7"
      Tab(2).Control(17)=   "Label8(0)"
      Tab(2).Control(18)=   "Label17(0)"
      Tab(2).Control(19)=   "Label3"
      Tab(2).Control(20)=   "Label2"
      Tab(2).Control(21)=   "Label8(1)"
      Tab(2).Control(22)=   "Label8(2)"
      Tab(2).Control(23)=   "Label9"
      Tab(2).Control(24)=   "Label12(0)"
      Tab(2).Control(25)=   "Label12(1)"
      Tab(2).Control(26)=   "Label18"
      Tab(2).ControlCount=   27
      TabCaption(3)   =   "Activo D03"
      TabPicture(3)   =   "frmActivos_WizardOP.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lsw"
      Tab(3).Control(1)=   "chkResponsables"
      Tab(3).Control(2)=   "txtDepartamento"
      Tab(3).Control(3)=   "txtSeccion"
      Tab(3).Control(4)=   "Label12(2)"
      Tab(3).Control(5)=   "Label16"
      Tab(3).Control(6)=   "Label15"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Mejoras"
      TabPicture(4)   =   "frmActivos_WizardOP.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label4(1)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label17(1)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label6"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label4(2)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label10"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label1(0)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label17(2)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label14(0)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label11"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "imgLibros"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "lblActivo"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "dtpFecha"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "txtMeses"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "cboVidaUtil"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "txtMDescripcion"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "txtMonto"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "cboM"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "txtMCodigo"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "txtMDocCompra"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "txtMProveedor"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "fraLibros"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).ControlCount=   21
      TabCaption(5)   =   "Finalizar"
      TabPicture(5)   =   "frmActivos_WizardOP.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdFinalizar"
      Tab(5).Control(1)=   "txtSumario"
      Tab(5).ControlCount=   2
      Begin VB.Frame fraLibros 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   2970
         TabIndex        =   70
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         Begin VB.Label Label1 
            Caption         =   "Ultimo Periodo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Total Historico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   79
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Total Rescate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   78
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblPeriodo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1560
            TabIndex        =   77
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblHistorico 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1560
            TabIndex        =   76
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblRescate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1560
            TabIndex        =   75
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Depreciación Acu"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   74
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblDepreciacion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1560
            TabIndex        =   73
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label lblLibros 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1560
            TabIndex        =   72
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Valor en Libros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   71
            Top             =   1680
            Width           =   1335
         End
      End
      Begin VB.TextBox txtMProveedor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   67
         ToolTipText     =   "Presione F4 para consultar / Cualquier Tecla"
         Top             =   2880
         Width           =   4785
      End
      Begin VB.TextBox txtMDocCompra 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   66
         Top             =   3240
         Width           =   4785
      End
      Begin VB.CommandButton cmdFinalizar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Finalizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72840
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox txtSumario 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2415
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   480
         Width           =   6375
      End
      Begin VB.TextBox txtMCodigo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1230
         TabIndex        =   55
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cboM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   1770
         Width           =   4815
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3990
         TabIndex        =   53
         Top             =   2130
         Width           =   1995
      End
      Begin VB.TextBox txtMDescripcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   1230
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   810
         Width           =   4755
      End
      Begin VB.ComboBox cboVidaUtil 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmActivos_WizardOP.frx":00A8
         Left            =   3990
         List            =   "frmActivos_WizardOP.frx":00AA
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   2490
         Width           =   2055
      End
      Begin VB.TextBox txtMeses 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2490
         Width           =   1395
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   2115
         Left            =   -73740
         TabIndex        =   47
         Top             =   1395
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   3731
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cédula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.CheckBox chkResponsables 
         BackColor       =   &H00808080&
         Caption         =   "Listado General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -73680
         TabIndex        =   46
         Top             =   1160
         Width           =   1575
      End
      Begin VB.TextBox txtDepartamento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73740
         Locked          =   -1  'True
         TabIndex        =   43
         ToolTipText     =   "Presione F4 para consultar / Cualquier Tecla"
         Top             =   480
         Width           =   5145
      End
      Begin VB.TextBox txtSeccion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73740
         Locked          =   -1  'True
         TabIndex        =   42
         ToolTipText     =   "Presione F4 para consultar / Cualquier Tecla"
         Top             =   840
         Width           =   5145
      End
      Begin VB.ComboBox cboVU 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmActivos_WizardOP.frx":00AC
         Left            =   -72990
         List            =   "frmActivos_WizardOP.frx":00B6
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1200
         Width           =   990
      End
      Begin VB.TextBox txtVU 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73740
         TabIndex        =   30
         Top             =   1200
         Width           =   705
      End
      Begin VB.ComboBox cbo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmActivos_WizardOP.frx":00C7
         Left            =   -70320
         List            =   "frmActivos_WizardOP.frx":00D1
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1200
         Width           =   1755
      End
      Begin VB.TextBox txtValorRescate 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70320
         TabIndex        =   28
         Top             =   480
         Width           =   1755
      End
      Begin VB.TextBox txtValorHistorico 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73740
         TabIndex        =   27
         Top             =   480
         Width           =   1755
      End
      Begin VB.TextBox txtUDProducidas 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70320
         TabIndex        =   26
         Top             =   1560
         Width           =   1755
      End
      Begin VB.TextBox txtUDAnio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73740
         TabIndex        =   25
         Top             =   1560
         Width           =   1755
      End
      Begin VB.TextBox txtModelo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73740
         TabIndex        =   20
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtSerie 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70320
         TabIndex        =   19
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtMarca 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73740
         TabIndex        =   18
         Top             =   2520
         Width           =   1755
      End
      Begin VB.TextBox txtOtrasSenas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -73740
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2880
         Width           =   5295
      End
      Begin VB.TextBox txtNotas 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   -73740
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1560
         Width           =   5145
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -73740
         TabIndex        =   9
         Top             =   480
         Width           =   1425
      End
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -73740
         TabIndex        =   8
         Top             =   840
         Width           =   5145
      End
      Begin VB.TextBox txtDocCompra 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -73740
         TabIndex        =   7
         Top             =   3000
         Width           =   5145
      End
      Begin VB.TextBox txtProveedor 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -73740
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Presione F4 para consultar / Cualquier Tecla"
         Top             =   2640
         Width           =   5145
      End
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmActivos_WizardOP.frx":00F3
         Left            =   -73740
         List            =   "frmActivos_WizardOP.frx":00FD
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   5175
      End
      Begin VB.OptionButton opt 
         Caption         =   "Es un Nuevo Activo"
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
         Index           =   1
         Left            =   -74160
         TabIndex        =   4
         Top             =   1920
         Width           =   4455
      End
      Begin VB.OptionButton opt 
         Caption         =   "Es una Mejora / Adición de un Activo Existente"
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
         Left            =   -74160
         TabIndex        =   3
         Top             =   1560
         Value           =   -1  'True
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker dtpAdquisicion 
         Height          =   315
         Left            =   -73740
         TabIndex        =   32
         Top             =   840
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   248446979
         CurrentDate     =   36338
      End
      Begin MSComCtl2.DTPicker dtpInstalacion 
         Height          =   315
         Left            =   -70320
         TabIndex        =   33
         Top             =   840
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   248446979
         CurrentDate     =   36338
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1230
         TabIndex        =   56
         Top             =   2130
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   248446977
         CurrentDate     =   36674
      End
      Begin VB.Label lblActivo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3000
         TabIndex        =   81
         Top             =   480
         Width           =   3315
      End
      Begin VB.Image imgLibros 
         Height          =   255
         Left            =   2640
         Picture         =   "frmActivos_WizardOP.frx":010E
         Stretch         =   -1  'True
         ToolTipText     =   "Valor en Libros"
         Top             =   525
         Width           =   285
      End
      Begin VB.Label Label11 
         Caption         =   "Doc.Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   68
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2910
         TabIndex        =   63
         Top             =   2130
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Activo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   62
         Top             =   510
         Width           =   585
      End
      Begin VB.Label Label10 
         Caption         =   "Justificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   61
         Top             =   1770
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   60
         Top             =   2130
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   59
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Vida Util"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2910
         TabIndex        =   58
         Top             =   2490
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Meses V.U."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   57
         Top             =   2490
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   $"frmActivos_WizardOP.frx":0550
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
         Height          =   1092
         Left            =   -74880
         TabIndex        =   49
         Top             =   360
         Width           =   6492
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Responsables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   -73740
         TabIndex        =   48
         Top             =   1155
         Width           =   5145
      End
      Begin VB.Label Label16 
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   -74880
         TabIndex        =   45
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label15 
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   44
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   -68520
         X2              =   -74880
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   -68520
         X2              =   -74880
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label4 
         Caption         =   "Instalación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -71670
         TabIndex        =   41
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Valor de rescate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71670
         TabIndex        =   40
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label8 
         Caption         =   "Depreciación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -71670
         TabIndex        =   39
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label17 
         Caption         =   "Valor Histórico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Adquisición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Vida útil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Ud. a Producir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -71640
         TabIndex        =   35
         Top             =   1635
         Width           =   1185
      End
      Begin VB.Label Label8 
         Caption         =   "Ud. x Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   34
         Top             =   1635
         Width           =   1185
      End
      Begin VB.Label Label9 
         Caption         =   "Modelo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   23
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -71640
         TabIndex        =   22
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Otras Señas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Placa"
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
         Left            =   -74850
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label14 
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
         Height          =   252
         Index           =   1
         Left            =   -74880
         TabIndex        =   15
         Top             =   2640
         Width           =   1092
      End
      Begin VB.Label Label13 
         Caption         =   "Doc.Compra"
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
         Left            =   -74880
         TabIndex        =   14
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label21 
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
         Left            =   -74850
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label19 
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
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
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
         Left            =   -74880
         TabIndex        =   11
         Top             =   1590
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmActivos_WizardOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbActiva(i As Integer)
Dim x As Byte

cmdAtras.Enabled = True
cmdSiguiente.Enabled = True

If i = ssTab.Tabs - 1 Then
  cmdSiguiente.Enabled = False
End If

If i = 0 Then
  cmdAtras.Enabled = False
End If

For x = 0 To ssTab.Tabs - 1
 ssTab.TabEnabled(x) = False
Next x

 ssTab.TabEnabled(i) = True
 ssTab.Tab = i

End Sub

Private Sub chkResponsables_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

  lsw.ListItems.Clear
  
  If chkResponsables.Value = vbChecked Then
     strSQL = "select R.cedula,R.nombre,A.num_placa" _
            & " from Activos_responsables R left join Activos_Responsables A" _
            & " on R.cedula = A.cedula and A.num_placa = '" & txtCodigo _
            & "' and A.estado = 'A' order by A.num_placa desc,R.nombre"
  Else
     strSQL = "select R.cedula,R.nombre,A.num_placa" _
            & " from Activos_responsables R left join Activos_Responsables A" _
            & " on R.cedula = A.cedula and A.num_placa = '" & txtCodigo _
            & "' and A.estado = 'A' where R.cod_departamento = '" & txtDepartamento.Tag _
            & "' and R.cod_seccion = '" & txtSeccion.Tag _
            & "' order by A.num_placa desc,R.nombre"
  End If
  Call OpenRecordSet(rs, strSQL, 0)
  Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!cedula)
        itmX.SubItems(1) = rs!Nombre
    If Not IsNull(rs!num_placa) Then itmX.Checked = True
    rs.MoveNext
  Loop
  rs.Close

  If lsw.ListItems.Count = 0 And chkResponsables.Value = vbUnchecked Then
      chkResponsables.Value = vbChecked
      Call chkResponsables_Click
  End If

vError:
Me.MousePointer = vbDefault

End Sub

Private Sub cmdAtras_Click()
If opt.Item(0).Value = True Then
  If ssTab.Tab = 4 Then
      Call sbActiva(0)
  Else
      Call sbActiva(ssTab.Tab - 1)
  End If

Else
 'Activo
  Select Case ssTab.Tab
     Case 0, 1, 2, 3
      Call sbActiva(ssTab.Tab - 1)
     Case 5
      Call sbActiva(3)
  End Select
End If
End Sub

Private Sub cmdFinalizar_Click()

Select Case True
  Case opt.Item(0).Value 'Mejoras
    If fxValidaM Then
      Call sbGuardarM
    End If
  
  Case opt.Item(1).Value 'Activo
    If fxValida Then
      Call sbGuardar
    End If
End Select
End Sub

Private Sub cmdSiguiente_Click()
If opt.Item(0).Value = True Then
  If ssTab.Tab = 0 Then
      Call sbActiva(4)
  Else
      Call sbActiva(ssTab.Tab + 1)
  End If

Else
 'Activo
  Select Case ssTab.Tab
     Case 0, 1, 2
      Call sbActiva(ssTab.Tab + 1)
     Case 3
      Call sbActiva(5)
  End Select

End If

End Sub


Private Sub imgLibros_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If Len(txtMCodigo) = 0 Then Exit Sub

fraLibros.Visible = IIf((fraLibros.Visible = True), False, True)

If Not fraLibros.Visible Then Exit Sub

strSQL = "select depreciacion_periodo,depreciacion_Acum,Valor_historico,Valor_desecho" _
       & " from Activos_Principal where num_placa = '" & txtMCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  lblPeriodo.Caption = rs!depreciacion_periodo
  lblDepreciacion.Caption = rs!depreciacion_acum
  lblHistorico.Caption = rs!valor_historico
  lblRescate.Caption = Format(rs!valor_desecho, "Standard")
  
  rs.Close
  strSQL = "select isnull(sum(depreciacion_Acum),0) as depreciacion_Acum,isnull(sum(monto),0) as Historico" _
         & " from Activos_retiro_adicion where tipo in('A','V') and num_placa = '" _
         & txtMCodigo & "'"
  Call OpenRecordSet(rs, strSQL, 0)
  If Not rs.EOF And Not rs.BOF Then
      lblHistorico.Caption = CCur(lblHistorico.Caption) + rs!historico
      lblDepreciacion.Caption = CCur(lblDepreciacion.Caption) + rs!depreciacion_acum
  End If
  
  lblDepreciacion.Caption = Format(lblDepreciacion.Caption, "Standard")
  lblHistorico.Caption = Format(lblHistorico.Caption, "Standard")
  lblLibros.Caption = Format(CCur(lblHistorico.Caption) - CCur(lblDepreciacion.Caption), "Standard")
  
End If
rs.Close

End Sub

Private Function fxValidaM() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""
fxValidaM = True

'1. Verificar Periodo / Si esta cerrado no puede registrarse
'2. No se puede modificar si ya se le ha calculo depreciacion
'3. Verifica que la fecha de la adicion o retiro sea mayor a la fecha de adquisicion
'4. del activo
'5. No puede Modificar un Activo Retirado


strSQL = "select fecha_adquisicion from Activos_Principal where num_placa = '" _
       & txtMCodigo & "' and estado <> 'R'"
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
  vMensaje = vMensaje & vbCrLf & " - El Activo no existe, o ya fue retirado ..."
Else
  If DateDiff("d", rs!fecha_adquisicion, dtpFecha.Value) < 1 Then
      vMensaje = vMensaje & vbCrLf & " - La fecha del Movimiento no es válida, ya que es menor a la del activo ..."
  End If
End If
rs.Close

If txtMProveedor.Tag = "" Then vMensaje = vMensaje & vbCrLf & " - Proveedor no es válido ..."


strSQL = "select estado from Activos_periodos where anio = " & Year(dtpFecha.Value) _
       & " and mes = " & Month(dtpFecha.Value)
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
 If rs!Estado <> "P" Then
      vMensaje = vMensaje & vbCrLf & " - El Periodo del Movimiento ya fue cerrado ..."
 End If
End If
rs.Close

If txtMDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion del tipo de Movimiento no es válido ..."
If Not IsNumeric(txtMonto) Then vMensaje = vMensaje & vbCrLf & " - El Monto del movimiento no es válido ..."
If cboM.ListCount <= 0 Then vMensaje = vMensaje & vbCrLf & " - No existe ninguna Justificación ..."

If IsNumeric(txtMonto) Then
 If gAsistente.Tipo <> "" Then
    If CCur(txtMonto) > gAsistente.VU Then vMensaje = vMensaje & vbCrLf & " - Valor Historico es mayor al Monto Disponible por el Asistente ..."
 End If

End If

If Len(vMensaje) > 0 Then
  fxValidaM = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Function fxMeses() As Integer
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFecha As Date, vVidaUtil As Integer
Dim iMes As Integer, lngAnio As Long

On Error GoTo vError

strSQL = "select fecha_adquisicion,Vida_Util, Vida_Util_en" _
       & " from Activos_Principal where num_placa = '" & txtMCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
  fxMeses = 1

Else
  
  If UCase(rs!vida_util_en) = "A" Then vVidaUtil = rs!Vida_Util * 12
  
  Select Case Mid(cboVidaUtil.Text, 1, 1)
    Case "R"
        'Fecha de Terminacion del activo
        vFecha = DateAdd("m", vVidaUtil, rs!fecha_adquisicion)
        iMes = DateDiff("m", dtpFecha.Value, vFecha)
        If iMes < 0 Then
           fxMeses = 1
        Else
           fxMeses = iMes
        End If
    Case "S"
        fxMeses = vVidaUtil
  End Select

End If

rs.Close


Exit Function

vError:
 fxMeses = 1

End Function

Private Sub sbGuardarM()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vUltimo As Long

On Error GoTo vError

Call imgLibros_Click
fraLibros.Visible = False
  
strSQL = "select isnull(max(id),0) + 1 as Ultimo from Activos_retiro_adicion" _
       & " where num_placa = '" & txtMCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)
    vUltimo = rs!Ultimo
rs.Close
 
 strSQL = "insert into Activos_retiro_adicion(id,num_placa,descripcion,monto,fecha,tipo,cod_justificacion" _
        & ",compra_documento,compra_proveedor,tipo_vidaUtil,Meses_Calculo,depreciacion_periodo" _
        & ",depreciacion_acum,depreciacion_mes,creacion_user,creacion_fecha,venta_cliente" _
        & ",venta_documento,valor_libros) values(" _
        & vUltimo & ",'" & txtMCodigo & "','" & UCase(txtMDescripcion) & "'," & CCur(txtMonto) & ",'" _
        & Format(dtpFecha.Value, "yyyy/mm/dd") & "','A','" & SIFGlobal.fxCodText(cboM.Text) _
        & "','" & txtMDocCompra & "','" & txtMProveedor.Tag & "','" & Mid(cboVidaUtil.Text, 1, 1) _
        & "'," & fxMeses & ",0,0,0,'" & glogon.Usuario & "',getdate(),'',''," & CCur(lblLibros.Caption) & ")"
 Call ConectionExecute(strSQL)

Select Case gAsistente.Tipo
  Case "O" 'Obras en Proceso
    strSQL = "select isnull(max(idx),0) + 1 as Ultimo from Activos_obras_resultados" _
           & " where contrato = '" & gAsistente.Documento & "'"
    Call OpenRecordSet(rs, strSQL, 0)
    
    strSQL = "insert Activos_obras_resultados(idx,contrato,num_placa,tipo,id_adicion) values(" & rs!Ultimo _
           & ",'" & gAsistente.Documento & "','" & txtMCodigo & "','M'," & vUltimo & ")"
    Call ConectionExecute(strSQL)
    
    strSQL = "update Activos_obras set distribuido = distribuido + " & CCur(txtMonto) _
           & " where contrato = '" & gAsistente.Documento & "'"
    Call ConectionExecute(strSQL)
    
    rs.Close
        
    frmActivos_ObrasProceso.Show
    frmActivos_ObrasProceso.txtCodigo = gAsistente.Documento
    frmActivos_ObrasProceso.gWizardX
  Case "C" 'Compras
    frmActivos_ComprasNR.Show
    frmActivos_ComprasNR.TimerX.Interval = 20
End Select

UnLoad frmActivos_WizardOP


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub txtmCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Num_Placa"
  gBusquedas.Orden = "Num_Placa"
  gBusquedas.Consulta = "select Num_Placa,Nombre from Activos_Principal"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtMCodigo = gBusquedas.Resultado
  If txtMCodigo <> "" Then
    lblActivo.Caption = gBusquedas.Resultado2
    Call imgLibros_Click
    fraLibros.Visible = False
  End If
End If

End Sub

Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select nombre from Activos_Principal where num_placa = '" & txtMCodigo & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
 lblActivo.Caption = rs!Nombre
End If
rs.Close
End Sub

Private Sub txtMDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboM.SetFocus
End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError
 txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMDocCompra.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
 txtMonto = Format(CCur(txtMonto), "Standard")
vError:
End Sub

Private Sub txtMProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMCodigo.SetFocus
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_proveedor,descripcion from Activos_proveedores"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtMProveedor.Tag) Then
       txtMProveedor.Tag = gBusquedas.Resultado
       txtMProveedor = gBusquedas.Resultado2
    End If
End If

End Sub





Private Sub sbLimpiaPantalla()
Dim strSQL As String, rs As New ADODB.Recordset


txtDescripcion = ""

txtCodigo = ""

Call sbActivos_MetodosDepreciacion(cbo)

strSQL = "select rtrim(cod_justificacion) + ' - ' + rtrim(descripcion) as 'ItmX'" _
       & " from Activos_justificaciones where tipo = 'A'"
Call sbLlenaCbo(cboM, strSQL, False)

txtVU = ""
cboVU.Text = "Años"

txtUDProducidas = 0
txtUDProducidas.Locked = True

txtUDAnio = 0
txtUDAnio.Locked = True

vPaso = False
  strSQL = "select rtrim(tipo_activo) + ' - ' + rtrim(descripcion) as 'ItmX'" _
       & " from Activos_tipo_activo order by tipo_activo"
  Call sbLlenaCbo(cboTipo, strSQL, False)
vPaso = True
Call cboTipo_Click

txtValorHistorico = "0"
txtValorRescate = "0"

dtpAdquisicion.Value = fxFechaServidor
dtpInstalacion.Value = dtpAdquisicion.Value

txtNotas = ""
txtDepartamento = ""
txtDepartamento.Tag = ""

txtSeccion = ""
txtSeccion.Tag = ""

txtDocCompra = ""
txtProveedor = ""
txtProveedor.Tag = ""

txtModelo = ""
txtSerie = ""
txtMarca = ""
txtOtrasSenas = ""
chkResponsables.Value = vbUnchecked
lsw.ListItems.Clear


'Mejoras
dtpFecha.Value = dtpAdquisicion.Value
txtMDescripcion = ""
txtMCodigo = ""
txtMonto = ""

cboVidaUtil.Clear
cboVidaUtil.AddItem "Restante del Activo"
cboVidaUtil.AddItem "Suplementaria del Activo"
cboVidaUtil.Text = "Restante del Activo"

If gAsistente.Tipo <> "" Then
  txtDocCompra = gAsistente.Documento
  txtMDocCompra = gAsistente.Documento
  
  txtProveedor.Tag = gAsistente.Proveedor
  txtMProveedor.Tag = gAsistente.Proveedor
  
  strSQL = "select descripcion from Activos_proveedores where cod_proveedor = " & gAsistente.Proveedor
  Call OpenRecordSet(rs, strSQL, 0)
  If Not rs.EOF And Not rs.BOF Then
    txtProveedor = rs!Descripcion & ""
    txtMProveedor = rs!Descripcion & ""
  End If
  rs.Close
  
  txtValorHistorico = Format(gAsistente.VU, "Standard")
  txtMonto = Format(gAsistente.VU, "Standard")
End If




End Sub


Private Sub Form_Load()

vModulo = 36

ssTab.Tab = 0
Call sbActiva(0)
Call sbLimpiaPantalla

End Sub


Private Sub cbo_Click()
If Mid(cbo.Text, 1, 1) = "U" Then
  txtUDProducidas.Locked = False
Else
  txtUDProducidas.Locked = True
End If

txtUDProducidas.ForeColor = IIf(txtUDProducidas.Locked, vbBlack, vbBlue)

txtUDAnio.Locked = txtUDProducidas.Locked
txtUDAnio.ForeColor = txtUDProducidas.ForeColor

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub cboTipo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If Not vPaso Then Exit Sub

'Llenar con valores por defecto, del tipo de activo
strSQL = "select * from Activos_tipo_activo where tipo_activo = '" & SIFGlobal.fxCodText(cboTipo.Text) & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  cbo.Text = fxActivos_MetodoDepreciacion(rs!met_depreciacion)
  txtVU = rs!Vida_Util
  If rs!tipo_vida_util = "A" Then
    cboVU.Text = "Años"
  Else
    cboVU.Text = "Meses"
  End If
End If
rs.Close

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub cboVU_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
End Sub

Private Sub dtpAdquisicion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInstalacion.SetFocus
End Sub

Private Sub dtpInstalacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVU.SetFocus
End Sub


Private Function fxValida() As Boolean
Dim vMensaje As String, i As Integer, x As Boolean

x = False
vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion del Activo no es válido ..."
If dtpAdquisicion.Value > dtpInstalacion.Value Then vMensaje = vMensaje & vbCrLf & " - La fecha de Adquisición no puede ser menor a la de instalacion ..."

If Not IsNumeric(txtVU) Then vMensaje = vMensaje & vbCrLf & " - Vida Util no es válida ..."
If Not IsNumeric(txtValorHistorico) Then vMensaje = vMensaje & vbCrLf & " - Valor Historico no es válido ..."
If Not IsNumeric(txtValorRescate) Then vMensaje = vMensaje & vbCrLf & " - Vida Rescate no es válido ..."
If Not IsNumeric(txtUDAnio) Or Not IsNumeric(txtUDProducidas) Then vMensaje = vMensaje & vbCrLf & " - Las unidades de producción no son validas ..."


If IsNumeric(txtUDAnio) And IsNumeric(txtUDProducidas) Then
   If CCur(txtUDAnio) > CCur(txtUDProducidas) Then
     vMensaje = vMensaje & vbCrLf & " - Las unidades de producción Anual no pueden ser mayores a las totales..."
   End If
End If

If IsNumeric(txtValorHistorico) And IsNumeric(txtValorRescate) Then
 If CCur(txtValorHistorico) < CCur(txtValorRescate) Then vMensaje = vMensaje & vbCrLf & " - Valor Historico no puede ser menor al valor de rescate (desecho) ..."
 If gAsistente.Tipo <> "" Then
    If CCur(txtValorHistorico) > gAsistente.VU Then vMensaje = vMensaje & vbCrLf & " - Valor Historico es mayor al Monto Disponible por el Asistente ..."
  End If

End If

If txtDepartamento.Tag = "" Then vMensaje = vMensaje & vbCrLf & " - Departamento no es válido ..."
If txtSeccion.Tag = "" Then vMensaje = vMensaje & vbCrLf & " - Sección no es válida ..."
If txtProveedor.Tag = "" Then vMensaje = vMensaje & vbCrLf & " - Proveedor no es válido ..."

'Responsables (debe haber almenos uno)
For i = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(i).Checked = True Then
   x = True
 End If
Next i
If Not x Then vMensaje = vMensaje & vbCrLf & " - No se ha asignado ningun responsable para el activo ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

   strSQL = "insert into Activos_Principal(num_placa,nombre,tipo_activo,descripcion,met_depreciacion" _
          & ",vida_util_en,vida_util,valor_historico,valor_desecho,fecha_adquisicion,fecha_instalacion" _
          & ",cod_departamento,cod_seccion,cod_proveedor,compra_documento,num_serie,marca,modelo" _
          & ",otras_senas,estado,depreciacion_acum,depreciacion_mes,depreciacion_periodo,ud_produccion" _
          & ",ud_anio,creacion_fecha,creacion_user) " _
          & " values('" & txtCodigo & "','" & UCase(txtDescripcion) & "','" & SIFGlobal.fxCodText(cboTipo.Text) & "','" & txtNotas _
          & "','" & fxActivos_MetodoDepreciacion(cbo.Text) & "','" & Mid(cboVU.Text, 1, 1) & "'," & txtVU & "," & CCur(txtValorHistorico) _
          & "," & CCur(txtValorRescate) & ",'" & Format(dtpAdquisicion.Value, "yyyy/mm/dd") _
          & "','" & Format(dtpInstalacion.Value, "yyyy/mm/dd") & "','" & txtDepartamento.Tag _
          & "','" & txtSeccion.Tag & "','" & txtProveedor.Tag & "','" & txtDocCompra _
          & "','" & txtSerie & "','" & txtMarca & "','" & txtModelo & "','" _
          & txtOtrasSenas & "','A',0,0,0," & CCur(txtUDProducidas) & "," & CCur(txtUDAnio) & ",getdate(),'" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)
    
   For i = 1 To lsw.ListItems.Count
     If lsw.ListItems.Item(i).Checked = True Then
         strSQL = "insert Activos_Responsables(num_placa,cedula,fecha,estado) values('" & txtCodigo _
                & "','" & lsw.ListItems.Item(i).Text & "',getdate(),'A')"
         Call ConectionExecute(strSQL)
     End If
   Next i
   

Select Case gAsistente.Tipo
  Case "O" 'Obras en Proceso
    strSQL = "select isnull(max(idx),0) + 1 as Ultimo from Activos_obras_resultados" _
           & " where contrato = '" & gAsistente.Documento & "'"
    Call OpenRecordSet(rs, strSQL, 0)
    strSQL = "insert Activos_obras_resultados(idx,contrato,num_placa,tipo) values(" & rs!Ultimo _
           & ",'" & gAsistente.Documento & "','" & txtCodigo & "','A')"
    Call ConectionExecute(strSQL)
    
    strSQL = "update Activos_obras set distribuido = distribuido + " & CCur(txtValorHistorico) _
           & " where contrato = '" & gAsistente.Documento & "'"
    Call ConectionExecute(strSQL)
    
    rs.Close
        
    frmActivos_ObrasProceso.Show
    frmActivos_ObrasProceso.txtCodigo = gAsistente.Documento
    frmActivos_ObrasProceso.gWizardX
  Case "C" 'Compras
    frmActivos_ComprasNR.Show
    frmActivos_ComprasNR.TimerX.Interval = 20
End Select

UnLoad frmActivos_WizardOP



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub



Private Sub Form_Unload(Cancel As Integer)
gAsistente.Documento = ""
gAsistente.Proveedor = ""
gAsistente.Tipo = ""
gAsistente.VU = 0
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
If ssTab.Tab < 5 Then Exit Sub

Select Case True
  Case opt.Item(0) 'Mejoras
    txtSumario = "PLACA          : " & txtMCodigo & vbCrLf _
               & "DESCRIPCION    : " & txtMDescripcion & vbCrLf _
               & "JUSTIFICACION  : " & cboM.Text & vbCrLf _
               & "FECHA          : " & dtpFecha.Value & vbCrLf _
               & "MONTO          : " & txtMonto & vbCrLf _
               & "PROVEEDOR      : " & txtMProveedor & vbCrLf _
               & "DOCUMENTO COMP : " & txtMDocCompra & vbCrLf
  
  Case opt.Item(1) 'Activo
    txtSumario = "PLACA          : " & txtCodigo & vbCrLf _
               & "DESCRIPCION    : " & txtDescripcion & vbCrLf _
               & "TIPO ACTIVO    : " & cboTipo.Text & vbCrLf _
               & "PROVEEDOR      : " & txtProveedor & vbCrLf _
               & "DOCUMENTO COMP : " & txtDocCompra & vbCrLf _
               & "VAL. HISTORICO : " & txtValorHistorico & vbCrLf _
               & "VAL. DESECHO   : " & txtValorRescate & vbCrLf _
               & "ADQUISICION    : " & dtpAdquisicion.Value & vbCrLf _
               & "VIDA UTIL(AÑOS): " & txtVU & vbCrLf _
               & "METODO DEPRECI.: " & cbo.Text & vbCrLf _
               & "MODELO/SER/MAR : " & txtModelo & "/" & txtSerie & "/" & txtMarca & vbCrLf _
               & "DEPARTAMENTO   : " & txtDepartamento & vbCrLf _
               & "SECCION        : " & txtSeccion

End Select

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub txtDepartamento_Change()
txtSeccion.Tag = ""
txtSeccion = ""
Call chkResponsables_Click
End Sub

Private Sub txtDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSeccion.SetFocus
  
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_departamento,descripcion from Activos_departamentos"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtDepartamento.Tag) Then
       txtDepartamento.Tag = gBusquedas.Resultado
       txtDepartamento = gBusquedas.Resultado2
       txtSeccion.SetFocus
    End If
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
End Sub

Private Sub txtDocCompra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call cmdSiguiente_Click
  txtValorHistorico.SetFocus
End If
End Sub

Private Sub txtMarca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOtrasSenas.SetFocus
End Sub

Private Sub txtModelo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSerie.SetFocus
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedor.SetFocus
End Sub

Private Sub txtOtrasSenas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call cmdSiguiente_Click
  txtDepartamento.SetFocus
End If
End Sub

Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocCompra.SetFocus
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
       txtDocCompra.SetFocus
    End If
End If

End Sub

Private Sub txtSeccion_Change()
Call chkResponsables_Click
End Sub

Private Sub txtSeccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedor.SetFocus
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_seccion,descripcion from Activos_secciones"
    gBusquedas.Filtro = " and cod_departamento = '" & txtDepartamento.Tag & "'"
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtSeccion.Tag) Then
       txtSeccion.Tag = gBusquedas.Resultado
       txtSeccion = gBusquedas.Resultado2
       lsw.SetFocus
    End If
End If
End Sub

Private Sub txtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMarca.SetFocus
End Sub


Private Sub txtUDAnio_GotFocus()
On Error GoTo vError
txtUDAnio = CCur(txtUDAnio)
vError:
End Sub

Private Sub txtUDAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUDProducidas.SetFocus
End Sub

Private Sub txtUDAnio_LostFocus()
On Error GoTo vError
txtUDAnio = Format(CCur(txtUDAnio), "Standard")
vError:
End Sub

Private Sub txtUDProducidas_GotFocus()
On Error GoTo vError
txtUDProducidas = CCur(txtUDProducidas)
vError:
End Sub

Private Sub txtUDProducidas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtModelo.SetFocus
End Sub

Private Sub txtUDProducidas_LostFocus()
On Error GoTo vError
txtUDProducidas = Format(CCur(txtUDProducidas), "Standard")
vError:
End Sub

Private Sub txtValorHistorico_GotFocus()
On Error GoTo vError
txtValorHistorico = CCur(txtValorHistorico)
vError:
End Sub

Private Sub txtValorHistorico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtValorRescate.SetFocus
End Sub

Private Sub txtValorHistorico_LostFocus()
On Error GoTo vError
txtValorHistorico = Format(CCur(txtValorHistorico), "Standard")
vError:
End Sub

Private Sub txtValorRescate_GotFocus()
On Error GoTo vError
txtValorRescate = CCur(txtValorRescate)
vError:
End Sub

Private Sub txtValorRescate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpAdquisicion.SetFocus
Exit Sub

vError:
  txtNotas.SetFocus
End Sub

Private Sub txtValorRescate_LostFocus()
On Error GoTo vError
txtValorRescate = Format(CCur(txtValorRescate), "Standard")
vError:
End Sub

Private Sub txtVU_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 If cboVU.Locked Then
    cbo.SetFocus
 Else
    cboVU.SetFocus
 End If
End If
End Sub

'/* mejoras  */

Private Sub cboM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFecha.SetFocus
End Sub

Private Sub cboVidaUtil_Click()
txtMeses = fxMeses
End Sub

Private Sub cboVidaUtil_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedor.SetFocus
End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

