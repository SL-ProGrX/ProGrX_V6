VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAF_CD_Asignaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comites.: Asignaciones "
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   7905
   Icon            =   "FrmAF_CD_Asignaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmAF_CD_Asignaciones.frx":08CA
   ScaleHeight     =   9225
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab1 
      Height          =   7680
      Left            =   45
      TabIndex        =   0
      Top             =   900
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   13547
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Agrupación de Comites"
      TabPicture(0)   =   "FrmAF_CD_Asignaciones.frx":711C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblComite"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdAplicar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdLimpiar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdElimina"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lswComites"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Prg1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "OptTodos"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "OptUno"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtcomite"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CmdImprimir"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ChkMostrar"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Actividades"
      TabPicture(1)   =   "FrmAF_CD_Asignaciones.frx":7138
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkMostrar1"
      Tab(1).Control(1)=   "CmdLimpia"
      Tab(1).Control(2)=   "CmdGuarda"
      Tab(1).Control(3)=   "TxtCodCom"
      Tab(1).Control(4)=   "LswAct"
      Tab(1).Control(5)=   "Prg2"
      Tab(1).Control(6)=   "LblComi"
      Tab(1).Control(7)=   "Label7"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Promotores"
      TabPicture(2)   =   "FrmAF_CD_Asignaciones.frx":7154
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtCom"
      Tab(2).Control(1)=   "ChkPromo"
      Tab(2).Control(2)=   "cmdaplicapromo"
      Tab(2).Control(3)=   "cmdLimpromo"
      Tab(2).Control(4)=   "Lswpromo"
      Tab(2).Control(5)=   "Label9"
      Tab(2).Control(6)=   "LblCom"
      Tab(2).ControlCount=   7
      Begin VB.TextBox TxtCom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73965
         TabIndex        =   35
         Top             =   600
         Width           =   840
      End
      Begin VB.CheckBox ChkPromo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Mostrar todos los promotores"
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
         Height          =   345
         Left            =   -70680
         TabIndex        =   33
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmdaplicapromo 
         Caption         =   "Asignar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -68955
         Picture         =   "FrmAF_CD_Asignaciones.frx":7170
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Asignar Promotores a los Comites"
         Top             =   6720
         Width           =   1005
      End
      Begin VB.CommandButton cmdLimpromo 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -69960
         Picture         =   "FrmAF_CD_Asignaciones.frx":72A8
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Nuevo Ingreso"
         Top             =   6720
         Width           =   1005
      End
      Begin VB.CheckBox ChkMostrar1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Mostrar todas las Actividades"
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
         Height          =   285
         Left            =   -70920
         TabIndex        =   30
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton CmdLimpia 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -69975
         Picture         =   "FrmAF_CD_Asignaciones.frx":73E2
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Nuevo Ingreso"
         Top             =   6720
         Width           =   1005
      End
      Begin VB.CommandButton CmdGuarda 
         Caption         =   "Asignar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -68970
         Picture         =   "FrmAF_CD_Asignaciones.frx":751C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Asignar Actividades a los Comites"
         Top             =   6720
         Width           =   1005
      End
      Begin VB.TextBox TxtCodCom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73920
         TabIndex        =   18
         Top             =   600
         Width           =   840
      End
      Begin VB.CheckBox ChkMostrar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Mostrar todos los comités"
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
         Height          =   405
         Left            =   4950
         TabIndex        =   17
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos de Agrupamiento"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   165
         TabIndex        =   8
         Top             =   4920
         Width           =   7200
         Begin VB.TextBox TxtNotas 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1290
            Left            =   4230
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   300
            Width           =   2865
         End
         Begin VB.Label Label6 
            Caption         =   "Notas"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4215
            TabIndex        =   16
            Top             =   -30
            Width           =   525
         End
         Begin VB.Label Label5 
            Caption         =   "Cantidad de Miembros Activos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   90
            TabIndex        =   15
            Top             =   1290
            Width           =   2280
         End
         Begin VB.Label LblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3195
            TabIndex        =   14
            Top             =   1245
            Width           =   900
         End
         Begin VB.Label LblScom 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3210
            TabIndex        =   13
            Top             =   795
            Width           =   900
         End
         Begin VB.Label Label4 
            Caption         =   "Cantidad miembros en comites Agrupados"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   90
            TabIndex        =   12
            Top             =   825
            Width           =   2985
         End
         Begin VB.Label LblPCom 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3210
            TabIndex        =   11
            Top             =   330
            Width           =   900
         End
         Begin VB.Label Label3 
            Caption         =   "Cantidad de Miembros ( Comite principal )"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   10
            Top             =   375
            Width           =   2985
         End
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Reporte"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2955
         Picture         =   "FrmAF_CD_Asignaciones.frx":7654
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6720
         Width           =   1005
      End
      Begin VB.TextBox txtcomite 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   915
         TabIndex        =   5
         ToolTipText     =   "Unidad programatica del comité"
         Top             =   1050
         Width           =   840
      End
      Begin VB.OptionButton OptUno 
         Appearance      =   0  'Flat
         Caption         =   "Comité Principal y sus UPs"
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
         Height          =   300
         Left            =   150
         TabIndex        =   2
         Top             =   6810
         Value           =   -1  'True
         Width           =   3120
      End
      Begin VB.OptionButton OptTodos 
         Appearance      =   0  'Flat
         Caption         =   "Todos los Comites "
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
         Height          =   270
         Left            =   150
         TabIndex        =   1
         Top             =   7185
         Width           =   2265
      End
      Begin MSComctlLib.ListView LswAct 
         Height          =   5025
         Left            =   -74715
         TabIndex        =   21
         Top             =   1455
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8864
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Actividad"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ProgressBar Prg1 
         Height          =   165
         Left            =   120
         TabIndex        =   22
         Top             =   4590
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ListView lswComites 
         Height          =   2685
         Left            =   150
         TabIndex        =   23
         Top             =   1905
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   4736
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   10054
         EndProperty
      End
      Begin MSComctlLib.ProgressBar Prg2 
         Height          =   165
         Left            =   -74730
         TabIndex        =   29
         Top             =   6480
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton CmdElimina 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   4320
         Picture         =   "FrmAF_CD_Asignaciones.frx":77C7
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Eliminar Unidades del Comité Principal"
         Top             =   6720
         Width           =   1005
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   5325
         Picture         =   "FrmAF_CD_Asignaciones.frx":7957
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Nuevo Ingreso"
         Top             =   6720
         Width           =   1005
      End
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "Asigna"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6330
         Picture         =   "FrmAF_CD_Asignaciones.frx":7A91
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Asignar Comites"
         Top             =   6720
         Width           =   1005
      End
      Begin MSComctlLib.ListView Lswpromo 
         Height          =   5010
         Left            =   -74745
         TabIndex        =   36
         Top             =   1500
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   8837
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Promotor"
            Object.Width           =   7673
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Comité"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74640
         TabIndex        =   38
         Top             =   615
         Width           =   495
      End
      Begin VB.Label LblCom 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -73110
         TabIndex        =   37
         Top             =   600
         Width           =   5085
      End
      Begin VB.Label LblComi 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -73080
         TabIndex        =   28
         Top             =   600
         Width           =   5040
      End
      Begin VB.Label Label7 
         Caption         =   "Comité"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74640
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.Label LblComite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1755
         TabIndex        =   26
         Top             =   1050
         Width           =   5550
      End
      Begin VB.Label Label1 
         Caption         =   "U.P."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   375
         TabIndex        =   25
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comité Principal"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   6945
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Asignaciones (Relaciones) del Comité"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   34
      Top             =   240
      Width           =   6795
   End
End
Attribute VB_Name = "FrmAF_CD_Asignaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim strSQL As String

Function FxNomComite(vUnidad As String)
   
   Dim rs As New ADODB.Recordset
   Dim strSQL As String
  
   strSQL = "select descripcion from uprogramatica where codigo = '" & vUnidad & "'"
            rs.Open strSQL, glogon.Conection, adOpenStatic
   If rs.EOF Then
      FxNomComite = "No existe unidad definida "
   Else
      FxNomComite = rs!Descripcion
   End If

End Function

Sub sbActividades()

Dim itmX As ListItem

strSQL = "select codtipo,descripcion from afi_cd_periocidadactividades"
         rs.Open strSQL, glogon.Conection, adOpenStatic
         LswAct.ListItems.Clear
        
        While Not rs.EOF
          Set itmX = LswAct.ListItems.Add(, , rs!Codtipo)
          itmX.SubItems(1) = IIf(IsNull(rs!Descripcion), "Sin Nombre", rs!Descripcion)
          rs.MoveNext
        Wend
        rs.Close

End Sub

Sub sbAplicaPromo()

Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, S As Integer
Dim CPromo As Integer

'On Error GoTo vError

Me.MousePointer = vbHourglass

For i = 1 To Lswpromo.ListItems.Count
     If Lswpromo.ListItems.Item(i).Checked = True Then
        CPromo = CPromo + 1
     End If
Next i
If CPromo >= 2 Then
      MsgBox "A seleccionado más de un promotor, seleccione el promotor correcto para esta zona", vbInformation, "Información"
      Me.MousePointer = vbDefault
      Exit Sub
End If

For i = 1 To Lswpromo.ListItems.Count
    
 If Lswpromo.ListItems.Item(i).Checked = True Then
        strSQL = "select * from afi_cd_cobertura where up = " & TxtCom.Text & "" _
                  & "and id_promotor = '" & Lswpromo.ListItems.Item(i) & "'"
                  rs.Open strSQL, glogon.Conection, adOpenStatic
    If rs.EOF Then
        strSQL = "insert into afi_cd_cobertura (id_promotor,up) " _
                & "values('" & Lswpromo.ListItems.Item(i) & "','" & TxtCom.Text & "')"
                glogon.Conection.Execute strSQL
    End If
  rs.Close
 End If
 
Next i
    
   
Me.MousePointer = vbDefault
MsgBox "El promotor fue asignado al comité satisfactoriamente", vbInformation, "Información"
ChkPromo.Value = 0

'Call Bitacora(vMovimiento, "Destino : " & Item.Text & " al codigo " & txtCodigoCorriente)
'Exit Sub

'vError:
' MsgBox Err.Description, vbCr

End Sub

Sub sbAplicar()

Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

'On Error GoTo vError
Me.MousePointer = vbHourglass
Prg1.Max = lswComites.ListItems.Count


For i = 1 To lswComites.ListItems.Count
        
    strSQL = "select * from afi_cd_agrupacomites where id_pricomite = '" & txtComite.Text & "' " _
             & "and id_segcomite = '" & lswComites.ListItems.Item(i) & "'"
              rs.Open strSQL, glogon.Conection, adOpenStatic
       
        If rs.EOF And lswComites.ListItems.Item(i).Checked = True Then
             strSQL = "insert into afi_cd_agrupacomites (id_pricomite,id_segcomite) " _
                      & "values('" & txtComite.Text & "','" & lswComites.ListItems.Item(i) & "')"
                      glogon.Conection.Execute strSQL
'          Else
'            strSql = "insert into afi_cd_agrupacomites (id_pricomite,id_segcomite) " _
'                      & "values('" & TxtComite.Text & "','111')"
'                      glogon.Conection.Execute strSql
'            MsgBox "Solo se almaceno el comité principal", vbInformation, "Información"
'            Me.MousePointer = vbDefault
'            Exit Sub
        End If
    rs.Close
   
    strSQL = "select * from afi_cd_descripcomites where id_pricomite = '" & txtComite.Text & "'"
              rs.Open strSQL, glogon.Conection, adOpenStatic
                   
              If rs.EOF Then
                  strSQL = "insert into afi_cd_descripcomites (id_pricomite,descripcion) " _
                           & "values('" & txtComite.Text & "','" & txtNotas.Text & "')"
                           glogon.Conection.Execute strSQL
              Else
                  strSQL = "update afi_cd_descripcomites set descripcion = '" & txtNotas.Text & "' " _
                           & "where id_pricomite = '" & txtComite.Text & "'"
                           glogon.Conection.Execute strSQL
              End If
     rs.Close
 
 Prg1.Value = Prg1.Value + 1
 Next i
Prg1.Value = 0
Me.MousePointer = vbDefault
MsgBox "Los Comites fueron agrupados al comite principal" & vbCrLf _
        & "" & Trim(lblComite.Caption) & " satisfactoriamente", vbInformation, "Información"
ChkMostrar.Value = 0

'Call Bitacora(vMovimiento, "Destino : " & Item.Text & " al codigo " & txtCodigoCorriente)
'Exit Sub

'vError:
' MsgBox Err.Description, vbCritical

End Sub

Sub sbAplicarAct()

Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
Prg2.Max = LswAct.ListItems.Count


For i = 1 To LswAct.ListItems.Count
        
    strSQL = "select * from afi_cd_comact where id_pricomite = '" & TxtCodCom.Text & "'" _
             & "and codtipo = '" & LswAct.ListItems.Item(i) & "'"
              rs.Open strSQL, glogon.Conection, adOpenStatic
       
        If rs.EOF And LswAct.ListItems.Item(i).Checked = True Then
             strSQL = "insert into afi_cd_comact (id_pricomite,codtipo) " _
                      & "values('" & TxtCodCom.Text & "','" & LswAct.ListItems.Item(i) & "')"
                      glogon.Conection.Execute strSQL
    
        End If
    rs.Close
   
        If LswAct.ListItems.Item(i).Checked = False Then
             strSQL = "delete afi_cd_comact where id_pricomite = '" & TxtCodCom.Text & "' " _
                      & " and codtipo = '" & LswAct.ListItems.Item(i) & "'"
                      glogon.Conection.Execute strSQL
    
        End If
                 
 Prg2.Value = Prg2.Value + 1
 Next i
Prg2.Value = 0
Me.MousePointer = vbDefault
MsgBox "La actualización de las Activideades para el comite " & vbCrLf _
        & "" & Trim(LblComi.Caption) & " se realizo satisfactoriamente", vbInformation, "Información"
ChkMostrar.Value = 0

'Call Bitacora(vMovimiento, "Destino : " & Item.Text & " al codigo " & txtCodigoCorriente)
Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 Me.MousePointer = vbDefault
End Sub

Sub sbBorraUnidades()
 Dim i As Integer

  For i = 1 To lswComites.ListItems.Count
       If lswComites.ListItems.Item(i).Checked = False Then
         strSQL = "delete afi_cd_agrupacomites where id_pricomite = '" & txtComite.Text & "' " _
                  & "and id_segcomite = '" & lswComites.ListItems.Item(i) & "'"
                  glogon.Conection.Execute strSQL
     End If
Next i
 Call sbCargaComite
 Call sbCalMiembros
End Sub

Sub sbCalMiembros()

Dim i As Integer, vUniPrinc As Integer, vSegFor As Integer
Dim vUnidades() As String
Dim vTotal As Currency
            
            strSQL = "select count(*) as Tprinc from socios where up = '" & txtComite & "'"
                     rs.Open strSQL, glogon.Conection, adOpenStatic
                     If Not rs.EOF Then
                       vUniPrinc = rs!tprinc
                     End If
            rs.Close
            
            strSQL = "select id_segcomite from afi_cd_agrupacomites where id_pricomite ='" & txtComite.Text & "'"
                      rs.Open strSQL, glogon.Conection, adOpenForwardOnly
                     If Not rs.EOF Then
                      ReDim vUnidades(rs.RecordCount)
                      vSegFor = rs.RecordCount
                          For i = 1 To rs.RecordCount
                           'Ingresa las unidades programaticas en el array
                           vUnidades(i) = rs!id_segcomite
                          rs.MoveNext
                          Next i
                     End If
            rs.Close
 
 'Calculando la cantidad de Asociados por Unidad
          For i = 1 To vSegFor
            strSQL = "select count(*) as Cantidad from socios where up = '" & vUnidades(i) & "'"
                     rs.Open strSQL, glogon.Conection, adOpenStatic
              
              'Suma Asociados
               vTotal = vTotal + rs!cantidad
                 
            rs.Close
          Next i
  LblPCom.Caption = vUniPrinc
  LblScom.Caption = vTotal
  LblTotal.Caption = CCur(LblPCom.Caption) + CCur(LblScom.Caption)

End Sub

Sub sbCargaAct()

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem
LswAct.ListItems.Clear
lblComite.Caption = FxNomComite(txtComite.Text)

strSQL = "select P.codtipo,P.descripcion from afi_cd_periocidadactividades P inner join afi_cd_comact C " _
         & "on P.codtipo = C.codtipo inner join afi_cd_agrupacomites A on A.id_pricomite = C.id_pricomite " _
         & "where A.id_pricomite = '" & TxtCodCom.Text & "' group by P.codtipo,P.descripcion"
         rs.Open strSQL, glogon.Conection, adOpenStatic


        While Not rs.EOF
             Set itmX = LswAct.ListItems.Add(, , rs!Codtipo)
             itmX.Checked = True
             itmX.SubItems(1) = Trim(rs!Descripcion)
             rs.MoveNext
        Wend
        rs.Close
                       
             If ChkMostrar1.Value = 1 Then
                    
                  strSQL = "select codtipo,descripcion from afi_cd_periocidadactividades " _
                           & "where codtipo not in " _
                           & "(select codtipo from afi_cd_comact where id_pricomite = '" & TxtCodCom.Text & "')"
                           rs.Open strSQL, glogon.Conection, adOpenStatic
                              
                              While Not rs.EOF
                                                 
                                        Set itmX = LswAct.ListItems.Add(, , rs!Codtipo)
                                        itmX.SubItems(1) = rs!Descripcion
                                
                                rs.MoveNext
                              Wend
                 rs.Close
             End If
End Sub

Sub sbCargaPromo()

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem
Lswpromo.ListItems.Clear
lblComite.Caption = FxNomComite(txtComite.Text)

strSQL = "select O.id_promotor,P.Nombre from afi_cd_cuentas C right join afi_cd_cobertura O " _
         & "on C.id_pricomite = O.up left join promotores P " _
         & "on O.id_promotor = P.id_promotor " _
         & "where O.up = " & TxtCom.Text & " and P.tipo = 'P' group by O.id_promotor,P.Nombre"
         rs.Open strSQL, glogon.Conection, adOpenStatic

        While Not rs.EOF
             Set itmX = Lswpromo.ListItems.Add(, , rs!id_promotor)
             itmX.Checked = True
             itmX.SubItems(1) = Trim(rs!Nombre)
             rs.MoveNext
        Wend
  rs.Close
                       
   If ChkPromo.Value = 1 Then
      strSQL = "select id_promotor,nombre from promotores " _
               & "where tipo = 'P' and id_promotor not in " _
               & "(select id_promotor from afi_cd_cobertura where up = '" & TxtCom.Text & "')"
               rs.Open strSQL, glogon.Conection, adOpenStatic
       
       While Not rs.EOF
          Set itmX = Lswpromo.ListItems.Add(, , rs!id_promotor)
              itmX.SubItems(1) = rs!Nombre
              rs.MoveNext
       Wend
     rs.Close
   End If

End Sub

Sub sbLimpia()
  
  txtComite.Text = ""
  lblComite.Caption = ""
  lswComites.ListItems.Clear
  ChkMostrar.Value = 0
  LblPCom.Caption = ""
  LblScom.Caption = ""
  LblTotal.Caption = ""
  txtNotas.Text = ""
  txtComite.SetFocus

End Sub

Private Sub ChkMostar1_Click()
 Call sbCargaAct
End Sub

Private Sub ChkMostrar_Click()
lswComites.SetFocus
Call sbCargaComite
End Sub

Private Sub ChkMostrar1_Click()
Call sbCargaAct
End Sub

Private Sub ChkPromo_Click()
 Call sbCargaPromo
End Sub

Private Sub cmdaplicapromo_Click()
 Call sbAplicaPromo
End Sub

Private Sub cmdAplicar_Click()
 Call sbAplicar
 Call sbCargaComite
 Call sbCalMiembros
End Sub
Private Sub CmdElimina_Click()
Dim i As Integer
Dim S As Integer

S = MsgBox("Desea eliminar los comites relacionados que desmarco", vbInformation + vbYesNo, "Información")

If S = vbYes Then

Me.MousePointer = vbHourglass
For i = 1 To lswComites.ListItems.Count
     If lswComites.ListItems.Item(i).Checked = False Then
         strSQL = "delete afi_cd_agrupacomites where id_pricomite = '" & txtComite.Text & "' " _
                  & "and id_segcomite = '" & lswComites.ListItems.Item(i) & "'"
                  glogon.Conection.Execute strSQL
     End If
Next i
 MsgBox "Unidades eliminadas satisfactoriamente", vbInformation, "Información"
 
 Call sbCargaComite
 Call sbCalMiembros
Me.MousePointer = vbDefault
End If

End Sub

Private Sub Cmdguarda_Click()
 Call sbAplicarAct
 ChkMostrar1.Value = 0
 Call sbCargaAct
End Sub

Private Sub Cmdimprimir_Click()

With Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .Connect = "pwd =" & glogon.RootKey
 .WindowTitle = "Reporte Asignación de Unidades a Comite Principal"
 .ReportFileName = App.Path & "\comitesd\Reportes\Afi_Cd_AsgComite.rpt"
 
 If OptUno.Value = True Then
  If txtComite.Text = Empty Then
  MsgBox "Debe de ingresar una unidad para poder imprimir las asignaciones de este comité", vbInformation, "Información"
  txtComite.SetFocus
  Exit Sub
  Else
   .SelectionFormula = "{afi_cd_agrupacomites.id_pricomite} = '" & txtComite.Text & "'"
  End If
 End If
 
'  strSql = strSql & "cdate({AFI_CD_CONTROLIQUIDA.FECHA}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
'  strSql = strSql & ") to Date (" & Format(DTPFinal.Value, "yyyy,mm,dd") & ")"
' .SelectionFormula = strSql
 .Formulas(0) = "fxTitulo='ASIGNACION DE UNIDADES A COMITE PRINCIPAL'"
 .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(3) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .PrintReport

End With

End Sub

Private Sub CmdLimpia_Click()

    TxtCodCom.Text = ""
    LblComi.Caption = ""
    LswAct.ListItems.Clear

End Sub

Private Sub cmdLimpiar_Click()
 Call sbLimpia

End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdLimpromo_Click()
 TxtCom.Text = ""
 LblCom.Caption = ""
 ChkPromo.Value = 0
 TxtCom.SetFocus
 Lswpromo.ListItems.Clear
End Sub

Private Sub Form_Activate()
' LswComites.SetFocus
  
End Sub

Private Sub Form_Load()
ssTab1.Tab = 0
' Call sbComites(LswComites)
 
End Sub
Sub sbCargaComite()

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem
lswComites.ListItems.Clear


strSQL = "select A.id_segcomite,U.descripcion as Comite,D.descripcion as notas from " _
         & "afi_cd_descripcomites D inner join afi_cd_agrupacomites A on D.id_pricomite = A.id_pricomite" _
         & " inner join uprogramatica U on A.id_segcomite = U.codigo " _
         & "where A.id_pricomite = '" & txtComite.Text & "' group by A.id_segcomite,U.descripcion,D.descripcion"
         rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF Then
        txtNotas.Text = rs!notas
        While Not rs.EOF
             Set itmX = lswComites.ListItems.Add(, , rs!id_segcomite)
             itmX.Checked = True
             itmX.SubItems(1) = Trim(rs!comite)
             rs.MoveNext
        Wend
        rs.Close
                       
             If ChkMostrar.Value = 1 Then
                  
                  strSQL = "select A.*,U.codigo,U.descripcion " _
                           & "from afi_cd_agrupacomites A right join uprogramatica U " _
                           & "on A.id_segcomite = U.codigo "
                           rs.Open strSQL, glogon.Conection, adOpenStatic
                              
                              While Not rs.EOF
                                      If IsNull(rs!id_pricomite) Then
                                        Set itmX = lswComites.ListItems.Add(, , rs!Codigo)
                                        itmX.SubItems(1) = rs!Descripcion
                                      End If
                              rs.MoveNext
                              Wend
                 rs.Close
             End If
                    
 Else
        rs.Close

         strSQL = "select * from uprogramatica order by codigo"
         rs.Open strSQL, glogon.Conection, adOpenStatic
         lswComites.ListItems.Clear
      
            While Not rs.EOF
              Set itmX = lswComites.ListItems.Add(, , rs!Codigo)
                  itmX.SubItems(1) = Trim(rs!Descripcion)
                 rs.MoveNext
            Wend
   rs.Close
End If

End Sub
Sub sbComites(lsw)

'Dim strSql As String, rs As New ADODB.Recordset

'Dim itMx As ListItem
'
'strSql = "select distinct N.id_promotor,P.nombre from afi_cd_nombramientos N left join promotores P" _
'         & " on N.id_promotor = P.id_promotor order by P.nombre"
'         rs.Open strSql, glogon.Conection, adOpenStatic
'lsw.ListItems.Clear
'
'While Not rs.EOF
' Set itMx = lsw.ListItems.Add(, , rs!id_promotor)
'     itMx.SubItems(1) = IIf(Not IsNull(rs!Nombre), rs!Nombre, "Sin Nombre")
' rs.MoveNext
'Wend
'rs.Close

End Sub



Private Sub lswactividades_ItemCheck(ByVal Item As MSComctlLib.ListItem)

'Dim strSql As String, vMovimiento As String
''On Error GoTo vError
'Fecha = Format(LswActividades.SelectedItem.SubItems(2), "yyyymmdd")
'
'If Item.Checked Then
'  strSql = "insert afi_cd_asignacion (id_promotor,cod_actividad,fecha,usuario) " _
'           & "values('" & LswComites.SelectedItem.Text & "','" & Item.Text & "'," _
'           & "'" & Fecha & "','" & glogon.Usuario & "')"
'Else
'  'vMovimiento = "Borrar"
'  strSql = "delete afi_cd_asignacion where id_promotor = '" & LswComites.SelectedItem.Text & "' and " _
'           & "cod_actividad = '" & Item.Text & "'"
'End If
'glogon.Conection.Execute strSql
'
''Call Bitacora(vMovimiento, "Destino : " & Item.Text & " al codigo " & txtCodigoCorriente)
'Exit Sub
'
''vError:
'' MsgBox Err.Description, vbCritical
End Sub



Private Sub LswComites_Click()
' Call sbAct
' LblComite.Caption = LswComites.SelectedItem.SubItems(1)
End Sub

Private Sub LswComites_KeyDown(KeyCode As Integer, Shift As Integer)
' Call sbAct
' LblComite.Caption = LswComites.SelectedItem.SubItems(1)
End Sub

Private Sub LswComites_KeyUp(KeyCode As Integer, Shift As Integer)
' Call sbAct
' LblComite.Caption = LswComites.SelectedItem.SubItems(1)
End Sub

Private Sub Lswpromo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim S As Integer

If Item.Checked = False Then
 S = MsgBox("Desea eliminar el Promotor relacionado a este comite", vbInformation + vbYesNo, "Información")
  If S = vbYes Then
     strSQL = "delete afi_cd_cobertura where id_promotor = " & Lswpromo.SelectedItem & " and " _
              & "up = " & TxtCom.Text & ""
              glogon.Conection.Execute strSQL
     MsgBox "Promotor Eliminado", vbInformation, "Información"
  End If
End If

End Sub


Private Sub Lswpromo_ItemClick(ByVal Item As MSComctlLib.ListItem)
 
 For i = 1 To Lswpromo.ListItems.Count
     If Lswpromo.ListItems.Item(i).Checked = True Then
        CPromo = CPromo + 1
     End If
Next i
If CPromo >= 2 Then
      MsgBox "A seleccionado más de un promotor, seleccione el promotor correcto para esta zona", vbInformation, "Información"
      Lswpromo.ListItems.Clear
      Call sbCargaPromo
      Exit Sub
End If

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

If ssTab1.Tab = 1 Then
    If txtComite.Text = Empty Then Exit Sub
    LswAct.ListItems.Clear
    TxtCodCom.Text = txtComite.Text
    LblComi.Caption = FxNomComite(TxtCodCom.Text)
    Call sbCargaAct
    
    'Carga Promotores
    TxtCom.Text = txtComite.Text
    LblCom.Caption = FxNomComite(TxtCodCom.Text)
    Call sbCargaPromo

End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
          Case 48 To 57, 8
          Case 13
            lblComite.Caption = FxNomComite(txtComite.Text)
          Case Else
           KeyAscii = 0
        End Select
End Sub


Private Sub TxtCodCom_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
        LblComi.Caption = FxNomComite(TxtCodCom.Text)
        Call sbCargaAct
      Case Else
       KeyAscii = 0
    End Select
End Sub


Private Sub TxtCom_KeyPress(KeyAscii As Integer)
 
 Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
        LblCom.Caption = FxNomComite(TxtCom.Text)
        Call sbCargaPromo
      Case Else
       KeyAscii = 0
    End Select

End Sub


Private Sub TxtComite_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
        If txtComite.Text = Empty Then Exit Sub
        lblComite.Caption = FxNomComite(txtComite.Text)
        Call sbCargaComite
        Call sbCalMiembros
      Case Else
       KeyAscii = 0
    End Select
End Sub
