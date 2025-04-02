VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmSYS_Contacto_Servicio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contact Center"
   ClientHeight    =   7404
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10176
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7404
   ScaleWidth      =   10176
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl TabInfo 
      Height          =   5292
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   9852
      _Version        =   1245187
      _ExtentX        =   17378
      _ExtentY        =   9334
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
      ItemCount       =   4
      Item(0).Caption =   "General"
      Item(0).ControlCount=   22
      Item(0).Control(0)=   "Label2(4)"
      Item(0).Control(1)=   "Label2(5)"
      Item(0).Control(2)=   "Label2(6)"
      Item(0).Control(3)=   "Label2(7)"
      Item(0).Control(4)=   "Label2(8)"
      Item(0).Control(5)=   "Label2(9)"
      Item(0).Control(6)=   "txtG_Provincia"
      Item(0).Control(7)=   "txtG_Canton"
      Item(0).Control(8)=   "txtG_Distrito"
      Item(0).Control(9)=   "txtG_Direccion"
      Item(0).Control(10)=   "txtG_Edad"
      Item(0).Control(11)=   "txtG_FecNac"
      Item(0).Control(12)=   "txtG_Ocupacion"
      Item(0).Control(13)=   "txtG_EstadoCivil"
      Item(0).Control(14)=   "txtG_Genero"
      Item(0).Control(15)=   "Label2(10)"
      Item(0).Control(16)=   "Label2(11)"
      Item(0).Control(17)=   "txtG_Salario"
      Item(0).Control(18)=   "txtG_Email_01"
      Item(0).Control(19)=   "Label2(12)"
      Item(0).Control(20)=   "txtG_Email_02"
      Item(0).Control(21)=   "txtG_Email_03"
      Item(1).Caption =   "Teléfonos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswTelefonos"
      Item(2).Caption =   "Direcciones"
      Item(2).ControlCount=   5
      Item(2).Control(0)=   "txtD_Direccion"
      Item(2).Control(1)=   "txtD_Distrito"
      Item(2).Control(2)=   "txtD_Canton"
      Item(2).Control(3)=   "txtD_Provincia"
      Item(2).Control(4)=   "lswDirecciones"
      Item(3).Caption =   "Empresas"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "lswEmpresas"
      Begin VB.TextBox txtG_Email_03 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2640
         Width           =   7572
      End
      Begin VB.TextBox txtG_Email_02 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2280
         Width           =   7572
      End
      Begin VB.TextBox txtG_Email_01 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1920
         Width           =   7572
      End
      Begin MSComctlLib.ListView lswDirecciones 
         Height          =   2652
         Left            =   -68920
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   7572
         _ExtentX        =   13356
         _ExtentY        =   4678
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Provincia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cantón"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Distrito"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Dirección Exacta"
            Object.Width           =   9596
         EndProperty
      End
      Begin VB.TextBox txtD_Provincia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68920
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3360
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.TextBox txtD_Canton 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -66400
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3360
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.TextBox txtD_Distrito 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -63880
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3360
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.TextBox txtD_Direccion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1392
         Left            =   -68920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   3720
         Visible         =   0   'False
         Width           =   7572
      End
      Begin VB.TextBox txtG_Salario 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1320
         Width           =   2052
      End
      Begin VB.TextBox txtG_Direccion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1272
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   3840
         Width           =   7572
      End
      Begin VB.TextBox txtG_Distrito 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3240
         Width           =   2532
      End
      Begin VB.TextBox txtG_Canton 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3240
         Width           =   2532
      End
      Begin VB.TextBox txtG_Provincia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3240
         Width           =   2532
      End
      Begin VB.TextBox txtG_Edad 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   2052
      End
      Begin VB.TextBox txtG_FecNac 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1320
         Width           =   2052
      End
      Begin VB.TextBox txtG_Ocupacion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   2052
      End
      Begin VB.TextBox txtG_EstadoCivil 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   720
         Width           =   2052
      End
      Begin VB.TextBox txtG_Genero 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   2052
      End
      Begin MSComctlLib.ListView lswTelefonos 
         Height          =   4572
         Left            =   -69400
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   8652
         _ExtentX        =   15261
         _ExtentY        =   8065
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Telefono"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Extensión"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Atiende?"
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComctlLib.ListView lswEmpresas 
         Height          =   4572
         Left            =   -69760
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   9372
         _ExtentX        =   16531
         _ExtentY        =   8065
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Empresa"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Ingreso"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Telefono 1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Telefono 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Email"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Salario"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Emails:"
         Height          =   252
         Index           =   12
         Left            =   1080
         TabIndex        =   37
         Top             =   1680
         Width           =   1572
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Salario"
         Height          =   252
         Index           =   11
         Left            =   6600
         TabIndex        =   27
         Top             =   1080
         Width           =   1452
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Edad"
         Height          =   252
         Index           =   10
         Left            =   3960
         TabIndex        =   26
         Top             =   1080
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección General"
         Height          =   252
         Index           =   9
         Left            =   1080
         TabIndex        =   16
         Top             =   3600
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lugar de Votación"
         Height          =   252
         Index           =   8
         Left            =   1080
         TabIndex        =   15
         Top             =   3000
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Nacimiento"
         Height          =   252
         Index           =   7
         Left            =   1080
         TabIndex        =   14
         Top             =   1080
         Width           =   1572
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ocupación"
         Height          =   252
         Index           =   6
         Left            =   6600
         TabIndex        =   13
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil"
         Height          =   252
         Index           =   5
         Left            =   3960
         TabIndex        =   12
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Genero"
         Height          =   252
         Index           =   4
         Left            =   1080
         TabIndex        =   11
         Top             =   480
         Width           =   1572
      End
   End
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   312
      Left            =   7440
      TabIndex        =   9
      Top             =   1440
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Consultar"
      Appearance      =   16
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   1812
   End
   Begin VB.TextBox txtApellido2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1812
   End
   Begin VB.TextBox txtApellido1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1440
      Width           =   1812
   End
   Begin VB.TextBox txtIdentificacion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1812
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   252
      Index           =   3
      Left            =   5520
      TabIndex        =   4
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido 2"
      Height          =   252
      Index           =   2
      Left            =   3720
      TabIndex        =   3
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido 1"
      Height          =   252
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Centro de Contactos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   4212
   End
   Begin VB.Image imgBanner 
      Height          =   996
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmSYS_Contacto_Servicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub sbLimpia()

TabInfo.Item(0).Selected = True

txtG_Genero.Text = ""
txtG_EstadoCivil.Text = ""
txtG_Ocupacion.Text = ""
txtG_FecNac.Text = ""
txtG_Edad.Text = ""
txtG_Salario.Text = "0.00"
txtG_Email_01.Text = ""
txtG_Email_02.Text = ""
txtG_Email_03.Text = ""
txtG_Provincia.Text = ""
txtG_Canton.Text = ""
txtG_Distrito.Text = ""
txtG_Direccion.Text = ""

End Sub

Private Sub btnConsulta_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
  Case TabInfo.Item(0).Selected 'General
    Call gBase_Padron(txtIdentificacion.Text, "General", rs, "CRC")
    Call sbLimpia
    If rs.RecordCount > 0 Then
        txtApellido1.Text = rs!Apellido_1
        txtApellido2.Text = rs!Apellido_2
        txtNombre.Text = rs!Nombre
    
        txtG_Genero.Text = rs!Sexo
        txtG_EstadoCivil.Text = rs!Estado_Civil
        txtG_Ocupacion.Text = rs!Profesion
        txtG_FecNac.Text = Format(rs!Fecha_Nacimiento, "dd/mm/yyyy")
        txtG_Edad.Text = ""
        txtG_Salario.Text = Format(rs!SALARIO, "Standard")
        txtG_Email_01.Text = rs!Email_01
        txtG_Email_02.Text = rs!Email_02
        txtG_Email_03.Text = rs!Email_03
        txtG_Provincia.Text = rs!Provincia
        txtG_Canton.Text = rs!Canton
        txtG_Distrito.Text = rs!Distrito
        txtG_Direccion.Text = rs!Direccion
    End If
    rs.Close
  Case TabInfo.Item(1).Selected 'Telefonos
    Call gBase_Padron(txtIdentificacion.Text, "Telefonos", rs, "CRC")
    With lswTelefonos.ListItems
      .Clear
      Do While Not rs.EOF
        Set itmX = .Add(, , rs!Telefono_Tipo)
            itmX.SubItems(1) = rs!Telefono
            itmX.SubItems(2) = rs!Extension
            itmX.SubItems(3) = rs!Atiende
        rs.MoveNext
      Loop
      rs.Close
    End With
      
  
  Case TabInfo.Item(2).Selected 'Direcciones
    Call gBase_Padron(txtIdentificacion.Text, "Direccion", rs, "CRC")
    With lswDirecciones.ListItems
      .Clear
      Do While Not rs.EOF
        Set itmX = .Add(, , rs!Provincia)
            itmX.SubItems(1) = rs!Canton
            itmX.SubItems(2) = rs!Distrito
            itmX.SubItems(3) = rs!Direccion
        rs.MoveNext
      Loop
      rs.Close
    End With
  
  
  Case TabInfo.Item(3).Selected 'Empresas
    Call gBase_Padron(txtIdentificacion.Text, "Empresas", rs, "CRC")
    With lswEmpresas.ListItems
      .Clear
      Do While Not rs.EOF
        Set itmX = .Add(, , rs!Nombre)
            itmX.SubItems(1) = rs!Canton
            itmX.SubItems(2) = Format(rs!FECHA_INGRESO, "dd/mm/yyyy")
            itmX.SubItems(3) = rs!TELEFONO_1
            itmX.SubItems(4) = rs!TELEFONO_2
            itmX.SubItems(5) = Format(rs!SALARIO, "Standard")
            itmX.SubItems(6) = rs!ACTIVO
        rs.MoveNext
      Loop
      rs.Close
    End With

End Select

Me.MousePointer = vbDefault

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 10
Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

End Sub


Private Sub lswDirecciones_DblClick()
If lswDirecciones.ListItems.Count <= 0 Then Exit Sub

With lswDirecciones.SelectedItem
    txtD_Provincia.Text = .Text
    txtD_Canton.Text = .SubItems(1)
    txtD_Distrito.Text = .SubItems(2)
    txtD_Direccion.Text = .SubItems(3)
End With

End Sub



Private Sub TabInfo_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Call btnConsulta_Click
End Sub


