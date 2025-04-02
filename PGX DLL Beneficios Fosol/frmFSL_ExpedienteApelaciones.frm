VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFSL_ExpedienteApelaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Apelaciones"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab ssTab 
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Histórico"
      TabPicture(0)   =   "frmFSL_ExpedienteApelaciones.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vgApelaciones"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Registro"
      TabPicture(1)   =   "frmFSL_ExpedienteApelaciones.frx":011B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAplicar"
      Tab(1).Control(1)=   "cboApelacion"
      Tab(1).Control(2)=   "txtPresentaCedula"
      Tab(1).Control(3)=   "txtPresentaNombre"
      Tab(1).Control(4)=   "txtPresentaNotas"
      Tab(1).Control(5)=   "Line3(2)"
      Tab(1).Control(6)=   "Label3(2)"
      Tab(1).Control(7)=   "Line3(1)"
      Tab(1).Control(8)=   "Label2(0)"
      Tab(1).Control(9)=   "Label3(0)"
      Tab(1).Control(10)=   "Label3(6)"
      Tab(1).Control(11)=   "Label3(7)"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Resolución"
      TabPicture(2)   =   "frmFSL_ExpedienteApelaciones.frx":0217
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraValidaMiembro"
      Tab(2).Control(1)=   "txtResolucionNotas"
      Tab(2).Control(2)=   "cboResolucion"
      Tab(2).Control(3)=   "lswComite"
      Tab(2).Control(4)=   "tlbResolucion"
      Tab(2).Control(5)=   "Label3(11)"
      Tab(2).Control(6)=   "Label3(12)"
      Tab(2).Control(7)=   "Label3(13)"
      Tab(2).Control(8)=   "Line3(4)"
      Tab(2).ControlCount=   9
      Begin VB.Frame fraValidaMiembro 
         Caption         =   "Validación del Miembro de Comité"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -73320
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox txtMiembroUsuario 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3480
            TabIndex        =   10
            ToolTipText     =   "Número de Tramite"
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txtMiembroClave 
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
            IMEMode         =   3  'DISABLE
            Left            =   3480
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   1080
            Width           =   2175
         End
         Begin MSComctlLib.Toolbar tlbValidaMiembro 
            Height          =   360
            Left            =   3000
            TabIndex        =   11
            Top             =   1680
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   635
            ButtonWidth     =   1640
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imgLista"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Valida"
                  Key             =   "Valida"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Cerrar"
                  Key             =   "Cerrar"
                  ImageIndex      =   3
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label lblMiembro 
            Alignment       =   1  'Right Justify
            Caption         =   "Miembro.: "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Width           =   5295
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   240
            X2              =   5760
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Clave .: "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   13
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Usuario .: "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   2040
            TabIndex        =   12
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -66600
         Picture         =   "frmFSL_ExpedienteApelaciones.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4560
         Width           =   975
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
         Height          =   330
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3840
         Width           =   4695
      End
      Begin VB.TextBox txtPresentaCedula 
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
         Left            =   -73680
         TabIndex        =   25
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtPresentaNombre 
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
         Left            =   -70680
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox txtPresentaNotas 
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
         Height          =   2175
         Left            =   -73680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   1440
         Width           =   8295
      End
      Begin VB.TextBox txtResolucionNotas 
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
         Height          =   3735
         Left            =   -74640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox cboResolucion 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -70200
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   5040
         Width           =   2895
      End
      Begin MSComctlLib.ListView lswComite 
         Height          =   3735
         Left            =   -70200
         TabIndex        =   17
         Top             =   840
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
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
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Width           =   1834
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbResolucion 
         Height          =   360
         Left            =   -66960
         TabIndex        =   18
         Top             =   5040
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   635
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgLista"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Guardar"
               Key             =   "Resolucion"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin FPSpreadADO.fpSpread vgApelaciones 
         Height          =   4935
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   9855
         _Version        =   524288
         _ExtentX        =   17383
         _ExtentY        =   8705
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         SpreadDesigner  =   "frmFSL_ExpedienteApelaciones.frx":0401
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   -74760
         X2              =   -65520
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Apelación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   -74640
         TabIndex        =   30
         Top             =   3720
         Width           =   855
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -74640
         X2              =   -65400
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Datos de la persona que apela.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   29
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Cédula"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   28
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -71640
         TabIndex        =   27
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   -74640
         TabIndex        =   26
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Notas de la Resolución.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -74640
         TabIndex        =   21
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Miembros del Comité.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -70200
         TabIndex        =   20
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   " Resolución.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   -72000
         TabIndex        =   19
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   -74640
         X2              =   -65400
         Y1              =   4800
         Y2              =   4800
      End
   End
   Begin VB.TextBox txtEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtExpediente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Número de Tramite"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtCedula 
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   6255
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   9840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_ExpedienteApelaciones.frx":0B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_ExpedienteApelaciones.frx":736A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_ExpedienteApelaciones.frx":7463
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Expediente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   10320
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmFSL_ExpedienteApelaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mComite As String, mLinea As Integer

Private Sub cmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If txtEstado.Tag <> "R" Then
   MsgBox "El expediente no se encuentra RECHAZADO para poder registrar una apelación!", vbExclamation
   Exit Sub
End If

If mLinea > 0 Then
   MsgBox "Ya se encuentra registrada una apelación (Pendiente de Resolución) a este expediente, verifique!", vbExclamation
   Exit Sub
End If

On Error GoTo vError

strSQL = "exec spFSL_ApelacionRegistra " & txtExpediente.Text & ",'" & SIFGlobal.fxSIFCodText(cboApelacion.Text) & "','" _
       & txtPresentaCedula.Text & "','" _
       & txtPresentaNombre.Text & "','" & txtPresentaNotas.Text & "','" & glogon.Usuario & "'"
glogon.Conection.Execute strSQL

MsgBox "Apelación registrada satisfactoriamente!", vbInformation
Call sbInicializa

Exit Sub

vError:
  MsgBox Err.Description, vbCritical


End Sub

Private Sub Form_Load()

vModulo = 22

txtExpediente.Text = GLOBALES.gTag

Call sbInicializa

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

SSTab.Tab = 0
strSQL = "select rtrim(cod_apelacion) + ' - ' + DESCRIPCION as ItmX from FSL_TIPOS_APELACIONES WHERE ACTIVA = 1"
Call sbLlenaCbo(cboApelacion, strSQL, False, False)

cboResolucion.Clear
cboResolucion.AddItem "Aprobado"
cboResolucion.AddItem "Rechazado"
cboResolucion.Text = "Aprobado"



strSQL = "select Soc.Nombre,Ex.*" _
       & " from FSL_Expedientes Ex inner join Socios Soc on Ex.cedula = Soc.Cedula" _
       & " Where Ex.Cod_Expediente = " & txtExpediente.Text
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF Or Not rs.BOF Then
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre

  
  txtPresentaCedula.Text = rs!Presenta_Cedula & ""
  txtPresentaNombre.Text = rs!PRESENTA_NOMBRE & ""
  txtPresentaNotas.Text = ""
  
  mComite = rs!cod_Comite
 txtEstado.Tag = rs!Estado
  Select Case rs!Estado
   Case "P" 'Pendiente
        txtEstado.Text = "PENDIENTE"
    Case "A" 'Aprobado
        txtEstado.Text = "APROBADO"
    Case "R" 'Rechazado
        txtEstado.Text = "RECHAZADO"
    Case "X" 'Aplicado
        txtEstado.Text = "APLICADO"
  End Select

End If
rs.Close

'Linea de Apelación Pendiente
mLinea = 0

strSQL = "select isnull( max(Linea),0) as 'Linea'" _
       & " from FSL_EXPEDIENTES_APELACIONES" _
       & " Where cod_Expediente = " & txtExpediente.Text & " and resolucion = 'P'"

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF Or Not rs.BOF Then
  mLinea = rs!Linea
End If
rs.Close

'Histórico
strSQL = "select Ta.Descripcion, Ea.*" _
       & " from FSL_EXPEDIENTES_APELACIONES Ea inner join FSL_TIPOS_APELACIONES Ta on Ea.COD_APELACION = Ta.COD_APELACION" _
       & " Where Ea.cod_Expediente = " & txtExpediente.Text & " order by registra_fecha desc"

rs.Open strSQL, glogon.Conection, adOpenStatic
vgApelaciones.MaxRows = 0

Do While Not rs.EOF
  vgApelaciones.MaxRows = vgApelaciones.MaxRows + 1
  vgApelaciones.Row = vgApelaciones.MaxRows
  
  vgApelaciones.Col = 1
  vgApelaciones.Text = rs!Descripcion

  vgApelaciones.CellNote = "Fecha : " & rs!Registra_Fecha & vbCrLf & "Usuario : " & rs!Registra_Usuario
  vgApelaciones.CellTag = CStr(rs!Linea)
   
    
  vgApelaciones.Col = 2
  vgApelaciones.Text = rs!notas
      
  vgApelaciones.Col = 3
  
  Select Case rs!Resolucion
     Case "P" 'Pendiente
          vgApelaciones.Text = "Pendiente"
     Case "A" 'Aprobada
          vgApelaciones.Text = "Aprobada"
     Case "R" 'Rechazada
          vgApelaciones.Text = "Rechazada"
  End Select
  
  vgApelaciones.Col = 4
  vgApelaciones.Text = rs!Registra_Fecha
      
  vgApelaciones.Col = 5
  vgApelaciones.Text = rs!Registra_Usuario
      
      
  vgApelaciones.Col = 6
  vgApelaciones.Text = rs!Presenta_Identificacion & ""
      
  vgApelaciones.Col = 7
  vgApelaciones.Text = rs!PRESENTA_NOMBRE & ""
      
      
  vgApelaciones.RowHeight(vgApelaciones.Row) = vgApelaciones.MaxTextRowHeight(vgApelaciones.Row)
  
 rs.MoveNext
Loop
rs.Close


'Resolucion
strSQL = "Select Cm.Cedula,Cm.Nombre,isnull(Ec.Asigna_Usuario,'No!') as 'ASIGNADO'" _
       & " from FSL_EXPEDIENTES Ex" _
       & "  inner join FSL_COMITES_MIEMBROS Cm on Ex.COD_COMITE = Cm.COD_COMITE" _
       & "   left join FSL_EXPEDIENTE_COMITE Ec on Ex.COD_EXPEDIENTE = Ec.COD_EXPEDIENTE and Ex.COD_COMITE = Ec.COD_COMITE" _
       & "         and Cm.Cedula = Ex.Cedula" _
       & " where Ex.cod_Expediente = " & txtExpediente.Text & " and Cm.Activo = 1"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 With lswComite.ListItems
    .Clear
    Do While Not rs.EOF
     Set itmX = .Add(, , rs!Nombre)
         itmX.Tag = rs!Cedula
         If rs!Asignado <> "No!" Then
            itmX.Checked = True
            itmX.SubItems(1) = "Validado"
         Else
            itmX.SubItems(1) = "Pendiente"
         End If
     rs.MoveNext
    Loop
 End With
 rs.Close



Me.MousePointer = vbDefault

End Sub

Private Sub lswComite_DblClick()
Dim strSQL As String, rs As New ADODB.Recordset

If lswComite.ListItems.Count = 0 Then Exit Sub

With lswComite.SelectedItem
  If .SubItems(1) = "Pendiente" Then
     fraValidaMiembro.Visible = True
     lblMiembro.Caption = .Text
     lblMiembro.Tag = .Tag

     strSQL = "select usuario_Vinculado from FSL_Comites_Miembros" _
           & " where cedula = '" & .Tag & "' and cod_comite = '" & mComite & "'"
     rs.Open strSQL, glogon.Conection, adOpenStatic
         txtMiembroUsuario.Text = rs!Usuario_Vinculado
     rs.Close
     
     txtMiembroClave.SetFocus

  End If
  
End With
End Sub

Private Sub tlbResolucion_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim y As Integer, vNumResolutores As Integer


If txtEstado.Tag = "X" Then
   MsgBox "Este Expediente se encuentra aplicado no se puede cambiar la resolución!", vbExclamation
   Exit Sub
End If

If mLinea = 0 Then
   MsgBox "No existe ninguna apelación pendiente de resolución", vbExclamation
   Exit Sub
End If


If Len(txtResolucionNotas.Text) < 10 Then
   MsgBox "Indique una nota válida para la resolución!", vbExclamation
   Exit Sub
End If

strSQL = "select NUMERO_RESOLUTORES from FSL_Comites" _
        & " where cod_Comite = '" & mComite & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
    vNumResolutores = rs!NUMERO_RESOLUTORES
rs.Close

y = 0
With lswComite.ListItems
   For i = 1 To .Count
      If .Item(i).Checked And Mid(.Item(i).SubItems(1), 1, 1) = "V" Then
            y = y + 1
      End If
   Next i
End With

If y < vNumResolutores Then
   MsgBox "Debe de indicar al menos (" & vNumResolutores & ") miembros del comité VALIDADOS! que den la resolución!", vbExclamation
   Exit Sub
End If

On Error GoTo vError
Me.MousePointer = vbHourglass


strSQL = "update FSL_EXPEDIENTES set RESOLUCION_ESTADO = '" & Mid(cboResolucion.Text, 1, 1) _
       & "', ESTADO = '" & Mid(cboResolucion.Text, 1, 1) & "' where COD_EXPEDIENTE = " & txtExpediente.Text
glogon.Conection.Execute strSQL

strSQL = "update FSL_EXPEDIENTES set RESOLUCION_ESTADO = '" & Mid(cboResolucion.Text, 1, 1) _
       & "', ESTADO = '" & Mid(cboResolucion.Text, 1, 1) & "' where COD_EXPEDIENTE = " & txtExpediente.Text
glogon.Conection.Execute strSQL


strSQL = "update FSL_EXPEDIENTES_APELACIONES set RESOLUCION_NOTAS = '" & txtResolucionNotas.Text _
       & "',RESOLUCION = '" & Mid(cboResolucion.Text, 1, 1) _
       & "',RESOLUCION_FECHA = getdate(), RESOLUCION_USUARIO = '" & glogon.Usuario _
       & "' where COD_EXPEDIENTE = " & txtExpediente.Text & " and Linea = " & mLinea
glogon.Conection.Execute strSQL


strSQL = "delete FSL_EXPEDIENTES_APELACIONES_COMITE WHERE COD_EXPEDIENTE = " & txtExpediente.Text
glogon.Conection.Execute strSQL


With lswComite.ListItems
   For i = 1 To .Count
      If .Item(i).Checked Then
            strSQL = "INSERT FSL_EXPEDIENTES_APELACIONES_COMITE(LINEA,COD_EXPEDIENTE,COD_COMITE,CEDULA,ASIGNA_FECHA,ASIGNA_USUARIO)" _
                   & " values(" & mLinea & "," & txtExpediente & ",'" & mComite & "','" & _
                   .Item(i).Tag & "',getdate(),'" & glogon.Usuario & "')"
            glogon.Conection.Execute strSQL
      End If
   Next i
End With

Me.MousePointer = vbDefault

MsgBox "Expediente actualizado satisfactoriamente...", vbInformation

Call sbInicializa

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub

Private Sub tlbValidaMiembro_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Valida"
     Call sbMiembroValida
  Case "Cerrar"
     fraValidaMiembro.Visible = False
     txtMiembroClave.Text = ""
End Select

End Sub

Private Sub sbMiembroValida()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

'Verifica Usuario / Cifrado Actual
strSQL = "exec spSEGLogon '" & txtMiembroUsuario.Text & "','" & SIFGlobal.fxSIFSeguridadCifrado(txtMiembroClave.Text) & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs!Existe = 0 Then
   MsgBox "No fue posible validar al usuario.: Verifique su contraseña!", vbExclamation
   txtMiembroClave.SetFocus
Else
 
 With lswComite.ListItems
   For i = 1 To .Count
     If .Item(i).Tag = lblMiembro.Tag Then
        .Item(i).SubItems(1) = "Validado"
        fraValidaMiembro.Visible = False
     End If
   Next i
 End With
End If
rs.Close

txtMiembroClave.Text = ""

End Sub

Private Sub txtMiembroClave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbMiembroValida
End Sub
