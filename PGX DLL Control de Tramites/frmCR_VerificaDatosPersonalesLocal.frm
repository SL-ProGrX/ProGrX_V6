VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCR_VerificaDatosPersonalesLocal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualización de datos personales"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8040
   HelpContextID   =   3029
   Icon            =   "frmCR_VerificaDatosPersonalesLocal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   975
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   3495
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6165
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
      TabCaption(0)   =   "Localización"
      TabPicture(0)   =   "frmCR_VerificaDatosPersonalesLocal.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label19(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label18"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label(25)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtpFecha"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboEstadoCivil"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDireccion"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboProvincia"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtEmail"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAptoPostal"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboSexo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtNotificaciones"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboDistrito"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cboCanton"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Conyuge"
      TabPicture(1)   =   "frmCR_VerificaDatosPersonalesLocal.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label16(7)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label16(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line9(6)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label(15)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label(16)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(6)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label(17)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label(18)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label(19)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Line9(7)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtConyugeNombre"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtConyugeCedula"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtConyugeTelTrabajoExt"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtConyugeTelTrabajo"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtConyugeTelCelular"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtAlbaceaCedula"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtAlbaceaNombre"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Nombramientos"
      TabPicture(2)   =   "frmCR_VerificaDatosPersonalesLocal.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label16(5)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label16(4)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(4)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(5)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Line9(5)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Line9(4)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label19(1)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "dtpNombramiento"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lswNombramiento"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtAniosSerivicio"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "optNombramiento(0)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "optNombramiento(1)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      Begin VB.ComboBox cboCanton 
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
         ItemData        =   "frmCR_VerificaDatosPersonalesLocal.frx":035E
         Left            =   1200
         List            =   "frmCR_VerificaDatosPersonalesLocal.frx":0360
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cboDistrito 
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
         ItemData        =   "frmCR_VerificaDatosPersonalesLocal.frx":0362
         Left            =   1200
         List            =   "frmCR_VerificaDatosPersonalesLocal.frx":0364
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtAlbaceaNombre 
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
         Left            =   -72960
         MaxLength       =   30
         TabIndex        =   53
         Top             =   3000
         Width           =   5175
      End
      Begin VB.TextBox txtAlbaceaCedula 
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
         Left            =   -74760
         MaxLength       =   15
         TabIndex        =   52
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txtConyugeTelCelular 
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
         Left            =   -71280
         MaxLength       =   15
         TabIndex        =   43
         Top             =   2160
         Width           =   1545
      End
      Begin VB.TextBox txtConyugeTelTrabajo 
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
         Left            =   -71280
         MaxLength       =   15
         TabIndex        =   42
         Top             =   1800
         Width           =   1545
      End
      Begin VB.TextBox txtConyugeTelTrabajoExt 
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
         Left            =   -68640
         MaxLength       =   6
         TabIndex        =   41
         Top             =   1800
         Width           =   945
      End
      Begin VB.TextBox txtConyugeCedula 
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
         Left            =   -74640
         MaxLength       =   15
         TabIndex        =   40
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtConyugeNombre 
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
         Left            =   -72840
         MaxLength       =   30
         TabIndex        =   39
         Top             =   1080
         Width           =   5175
      End
      Begin VB.OptionButton optNombramiento 
         Caption         =   "Interino"
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
         Left            =   -73560
         TabIndex        =   36
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton optNombramiento 
         Caption         =   "Propiedad"
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
         Left            =   -73560
         TabIndex        =   30
         Top             =   1080
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtAniosSerivicio 
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
         Height          =   315
         Left            =   -72960
         MaxLength       =   45
         TabIndex        =   29
         ToolTipText     =   "Campo para la Cédula de Identidad"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtNotificaciones 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   2640
         Width           =   6375
      End
      Begin VB.ComboBox cboSexo 
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
         ItemData        =   "frmCR_VerificaDatosPersonalesLocal.frx":0366
         Left            =   3960
         List            =   "frmCR_VerificaDatosPersonalesLocal.frx":0370
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Sexo del Socio"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtAptoPostal 
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
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   16
         ToolTipText     =   "Número de Apartado postal del socio"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   3960
         MaxLength       =   45
         TabIndex        =   15
         ToolTipText     =   "Dirección del correo electrónico"
         Top             =   2280
         Width           =   3615
      End
      Begin VB.ComboBox cboProvincia 
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
         ItemData        =   "frmCR_VerificaDatosPersonalesLocal.frx":0389
         Left            =   1200
         List            =   "frmCR_VerificaDatosPersonalesLocal.frx":03A2
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Provincia"
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtDireccion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   3600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   480
         Width           =   3975
      End
      Begin VB.ComboBox cboEstadoCivil 
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
         ItemData        =   "frmCR_VerificaDatosPersonalesLocal.frx":03EB
         Left            =   1200
         List            =   "frmCR_VerificaDatosPersonalesLocal.frx":0401
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Estado Civil de la persona"
         Top             =   1920
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   6360
         TabIndex        =   18
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   175505411
         CurrentDate     =   36832
      End
      Begin MSComctlLib.ListView lswNombramiento 
         Height          =   2415
         Left            =   -71880
         TabIndex        =   31
         Top             =   840
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Estado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "A Partir"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpNombramiento 
         Height          =   315
         Left            =   -73560
         TabIndex        =   32
         ToolTipText     =   "Fecha de Ingreso al sistema"
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   175505411
         CurrentDate     =   38899
         MaxDate         =   55153
         MinDate         =   14611
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   7560
         Y1              =   1720
         Y2              =   1720
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   -74880
         X2              =   -72480
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label 
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
         Height          =   315
         Index           =   0
         Left            =   -74760
         TabIndex        =   55
         Top             =   2760
         Width           =   1005
      End
      Begin VB.Label Label 
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
         Height          =   315
         Index           =   1
         Left            =   -72960
         TabIndex        =   54
         Top             =   2760
         Width           =   1005
      End
      Begin VB.Label Label 
         Caption         =   "Celular"
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
         Index           =   19
         Left            =   -72840
         TabIndex        =   50
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Label Label 
         Caption         =   "Extensión"
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
         Index           =   18
         Left            =   -69480
         TabIndex        =   49
         Top             =   1830
         Width           =   795
      End
      Begin VB.Label Label 
         Caption         =   "Trabajo"
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
         Index           =   17
         Left            =   -72840
         TabIndex        =   48
         Top             =   1830
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfonos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   -74640
         TabIndex        =   47
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label 
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
         Height          =   315
         Index           =   16
         Left            =   -72840
         TabIndex        =   46
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label 
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
         Height          =   315
         Index           =   15
         Left            =   -74640
         TabIndex        =   45
         Top             =   840
         Width           =   1005
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   -74880
         X2              =   -72480
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Años de Servicio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   -74640
         TabIndex        =   38
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   -74640
         X2              =   -72240
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   -71880
         X2              =   -69480
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   -74520
         TabIndex        =   34
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "A partir del"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   -74520
         TabIndex        =   33
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Notificaciones:"
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
         Index           =   25
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Sexo"
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
         Left            =   3240
         TabIndex        =   26
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "Distrito"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Apto Postal"
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
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "E-Mail"
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
         Left            =   3240
         TabIndex        =   23
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Fec.Nac."
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
         Left            =   5520
         TabIndex        =   22
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Provincia"
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
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Cantón"
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
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Estado Civil"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Situación Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   -74640
         TabIndex        =   35
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Historial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   -71880
         TabIndex        =   37
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos del Cónyuge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   44
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Albacea"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   56
         Top             =   2400
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   51
      Top             =   6315
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9596
            MinWidth        =   9596
            Object.ToolTipText     =   "Ultima Modificación"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblUnidadEtiqueta 
      Caption         =   "Sección"
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
      TabIndex        =   60
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblUnidadEtiqueta 
      Caption         =   "Departamento"
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
      Left            =   240
      TabIndex        =   59
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblInstCod 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Institución"
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
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8040
      X2              =   0
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblSeccion 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblSeccionDesc 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label lblDepartamento 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblDepartamentoDesc 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label lblCedula 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Cédula"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblInstDesc 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2640
      TabIndex        =   10
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmCR_VerificaDatosPersonalesLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vPaso As Boolean, vFechaActual As Date
Dim vCantonMascara As String, vDistritoMascara  As String

Private Sub sbConsulta()
Dim rs As New ADODB.Recordset, strSQL As String

On Error Resume Next

vFechaActual = fxFechaServidor

'Limpia Datos
lblNombre.Caption = ""
lblCedula.Caption = ""
txtDireccion = ""
cboEstadoCivil.Text = "Soltero"
cboSexo.Text = "Masculino"
txtEmail = ""
txtAptoPostal = ""
dtpFecha.Value = vFechaActual

ssTab.Tab = 0


If GLOBALES.SysASEVersion Then
   'ASE
    strSQL = "Select S.*,I.descripcion as DescInst,D.descripcion as DescDept" _
           & ",X.Ut_descripcion as DescSec" _
           & ",dbo.fxAFIAnioServicio(cedula,'" & Format(vFechaActual, "yyyy/mm/dd") & "') as AnioServicio" _
           & ",rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
           & " From socios S inner join Instituciones I on S.cod_institucion = I.cod_institucion" _
           & " left join uprogramatica D on S.UP = D.Codigo" _
           & " left join UTRABAJO X on S.UT = X.ut_codigo" _
           & " left join Provincias Prov on S.Provincia = Prov.Provincia" _
           & " left join Cantones Cant on S.Provincia = Cant.Provincia and S.Canton = Cant.Canton" _
           & " left join Distritos Dist on S.Provincia = Dist.Provincia and S.Canton = Dist.Canton and S.distrito = Dist.distrito" _
           & " where cedula='" & GLOBALES.gCedulaActual & "'"

Else
    'SIF
    strSQL = "Select S.*,I.descripcion as DescInst,D.descripcion as DescDept" _
           & ",X.descripcion as DescSec" _
           & ",dbo.fxAFIAnioServicio(cedula,'" & Format(vFechaActual, "yyyy/mm/dd") & "') as AnioServicio" _
           & ",rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
           & " From socios S inner join Instituciones I on S.cod_institucion = I.cod_institucion" _
           & " LEFT join AFDepartamentos D on S.cod_institucion = D.cod_institucion and S.cod_departamento = D.cod_departamento" _
           & " left join AFSecciones X on S.cod_institucion = X.cod_institucion" _
           & " left join Provincias Prov on S.Provincia = Prov.Provincia" _
           & " left join Cantones Cant on S.Provincia = Cant.Provincia and S.Canton = Cant.Canton" _
           & " left join Distritos Dist on S.Provincia = Dist.Provincia and convert(int,S.Canton) = convert(int,Dist.Canton) and S.distrito = Dist.distrito" _
           & " where cedula='" & GLOBALES.gCedulaActual & "'"


End If


rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF And Not rs.BOF Then
    lblNombre.Caption = rs!Nombre
    lblCedula.Caption = GLOBALES.gCedulaActual
    
    Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
    Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
    Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
        
    cboEstadoCivil.Text = fxEstadoCivil(rs!estadoCivil)
    txtDireccion = Trim(rs!Direccion & "")
    dtpFecha.Value = IIf(IsNull(rs!fecha_nac), Date, rs!fecha_nac)
    cboSexo.Text = IIf(IsNull(rs!sexo), "Masculino", IIf((rs!sexo = "F"), "Femenino", "Masculino"))
    txtAptoPostal = Trim(rs!apto & "")
    txtEmail = Trim(rs!AF_Email & "")
     
    
    StatusBarX.Panels(1).Text = "Ultima Actualización : "
    If IsNull(rs!ActualizaFecha) Then
      StatusBarX.Panels(1).Text = StatusBarX.Panels(1).Text & "No hay / Usuario : No hay"
    Else
      StatusBarX.Panels(1).Text = StatusBarX.Panels(1).Text & Format(rs!ActualizaFecha, "dd/mm/yyyy") & " / Usuario : " & rs!ActualizaUser
    End If
    
    If IsNull(rs!estadoLaboral) Then
      optNombramiento.Item(0).Value = True
    Else
     If rs!estadoLaboral = 1 Then
        optNombramiento.Item(0).Value = True
     Else
        optNombramiento.Item(1).Value = True
     End If
    End If
    
    If GLOBALES.SysASEVersion Then
        lblSeccion.Caption = rs!UT & ""
        lblDepartamento.Caption = rs!up & ""
    Else
        lblSeccion.Caption = rs!cod_seccion & ""
        lblDepartamento.Caption = rs!cod_departamento & ""
    End If
    
    lblSeccionDesc.Caption = rs!DescSec & ""
    lblDepartamentoDesc.Caption = rs!descDept & ""
    
    lblInstCod.Caption = rs!cod_institucion & ""
    lblInstDesc.Caption = rs!DescInst & ""
    
   dtpNombramiento.Value = IIf(IsNull(rs!nombramiento_fecha), rs!FechaIngreso, rs!nombramiento_fecha)
   lswNombramiento.ListItems.Clear
   
   txtAniosSerivicio.Text = Trim(rs!AnioServicio)
   
   txtConyugeCedula.Text = Trim(rs!conyuge_cedula & "")
   txtConyugeNombre.Text = Trim(rs!conyuge_nombre & "")
   txtConyugeTelCelular.Text = Trim(rs!conyuge_TelCell & "")
   txtConyugeTelTrabajo.Text = Trim(rs!conyuge_TelTra & "")
   txtConyugeTelTrabajoExt.Text = Trim(rs!conyuge_TelTraExt & "")
   
   txtAlbaceaCedula.Text = Trim(rs!albacea_Cedula & "")
   txtAlbaceaNombre.Text = Trim(rs!albacea_nombre & "")
   
   txtNotificaciones.Text = Trim(rs!Notificaciones & "")
    
     
End If
rs.Close

End Sub


Private Sub cboCanton_Click()
Dim vSQL As String

If vPaso Then Exit Sub

vSQL = " where Provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) _
     & " and Canton = '" & Format(cboCanton.ItemData(cboCanton.ListIndex), vCantonMascara) & "' order by descripcion"

Call sbCargaCbo(cboDistrito, "distritos", vSQL)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistrito.AddItem " "
cboDistrito.Text = " "

End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub

Private Sub cboProvincia_Click()
Dim vSQL As String

If vPaso Then Exit Sub

vSQL = " where provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) & " order by descripcion"

vPaso = True
 Call sbCargaCbo(cboCanton, "cantones", vSQL)
vPaso = False
Call cboCanton_Click

End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub cmdGrabar_Click()
Dim strSQL As String, bol As Boolean

On Error GoTo vError

If cboCanton.Text = "" Then
  MsgBox "Falta Información del Canton...", vbInformation
  Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "update socios set provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) & ", Canton = " & cboCanton.ItemData(cboCanton.ListIndex) _
       & ",distrito = '" & Format(cboDistrito.ItemData(cboDistrito.ListIndex), vDistritoMascara) _
       & "',estadocivil='" & fxEstadoCivil(cboEstadoCivil.Text) & "',direccion='" & txtDireccion.Text & "',fecha_nac='" & Format(dtpFecha, "yyyy/mm/dd") _
       & "',apto = '" & Trim(txtAptoPostal) & "',af_email = '" & Trim(txtEmail) & "',sexo = '" & IIf((cboSexo.Text = "Femenino"), "F", "M") _
       & "',EstadoLaboral = " & IIf((optNombramiento.Item(0).Value = True), 1, 0) & ",ActualizaFecha = GetDate(), ActualizaUser = '" & glogon.Usuario _
       & "',Nombramiento_Fecha = '" & Format(dtpNombramiento.Value, "yyyy/mm/dd") & "',Conyuge_Cedula = '" & txtConyugeCedula.Text _
       & "', Conyuge_Nombre = '" & txtConyugeNombre.Text & "',Conyuge_TelCell = '" & txtConyugeTelCelular.Text _
       & "',Conyuge_TelTra = '" & txtConyugeTelTrabajo.Text & "',Conyuge_TelTraExt = '" & txtConyugeTelTrabajoExt.Text _
       & "',Notificaciones = '" & txtNotificaciones.Text & "',Albacea_cedula = '" & txtAlbaceaCedula.Text & "',Albacea_nombre = '" _
       & txtAlbaceaNombre.Text & "' Where Cedula='" & GLOBALES.gCedulaActual & "'"
glogon.Conection.Execute strSQL


'Revisa Nombramiento / Variaciones para Registrarlas en el Histórico
strSQL = "exec spAFINombramiento '" & Trim(GLOBALES.gCedulaActual) & "'," & IIf(optNombramiento(0).Value = True, 1, 0) _
       & ",'" & Format(dtpNombramiento.Value, "yyyy/mm/dd") & "','" & glogon.Usuario & "'"
glogon.Conection.Execute strSQL

Call Bitacora("Modifica", "Informacion de la Persona con cedula=" & Trim(GLOBALES.gCedulaActual))

Me.MousePointer = vbDefault

MsgBox "Información Actualizada Satisfactoriamente...", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
End Sub



Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAptoPostal.SetFocus
End Sub


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
'CAMBIA AL MODULO DE AFILIACION
'PARA REGISTRAR EN BITACORAS, EL MOVIMIENTO EN EL MODULO CORRECTO
vModulo = 1

'Mascara del canton
vCantonMascara = "0"
strSQL = "select MAX(LEN(canton)) as Caracteres from CANTONES"
rs.Open strSQL, glogon.Conection
If Not rs.EOF And Not rs.BOF Then
 vCantonMascara = SIFGlobal.fxSIFRelleno(vCantonMascara, "D", "0", rs!Caracteres)
End If
rs.Close

'Mascara del distrito
vDistritoMascara = "0"
strSQL = "select MAX(LEN(distrito)) as Caracteres from Distritos"
rs.Open strSQL, glogon.Conection
If Not rs.EOF And Not rs.BOF Then
 vDistritoMascara = SIFGlobal.fxSIFRelleno(vDistritoMascara, "D", "0", rs!Caracteres)
End If
rs.Close


If GLOBALES.SysASEVersion Then
  lblUnidadEtiqueta(0).Caption = "U. Programatica"
  lblUnidadEtiqueta(1).Caption = "U. Trabajo"
Else
  lblUnidadEtiqueta(0).Caption = "Departamento"
  lblUnidadEtiqueta(1).Caption = "Sección"
End If


vPaso = True
Call sbCargaCbo(cboProvincia, "provincias")
vPaso = False

If GLOBALES.gCedulaActual <> "" Then
 Call sbConsulta
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Establece el Modulo de Credito para el sistema de seguridad y Bitacoras
vModulo = 3
End Sub




Private Sub SSTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListItem

If ssTab.Tab = 2 Then
   strSQL = "select * from afi_nombramientos where cedula = '" & GLOBALES.gCedulaActual & "' order by Fecha_Nombramiento desc"
   rs.Open strSQL, glogon.Conection, adOpenStatic
   lswNombramiento.ListItems.Clear
   Do While Not rs.EOF
     Set itmX = lswNombramiento.ListItems.Add(, , IIf((rs!Estado = "P"), "Propiedad", "Interino"))
         itmX.SubItems(1) = Format(rs!fecha_nombramiento, "dd/mm/yyyy")
         itmX.SubItems(2) = rs!Registro_Fecha
         itmX.SubItems(3) = rs!registro_usuario
     rs.MoveNext
   Loop
   rs.Close
End If
End Sub

Private Sub txtAlbaceaCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAlbaceaNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtAlbaceaCedula = gBusquedas.Resultado
  txtAlbaceaNombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtAlbaceaNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeCedula.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtAlbaceaCedula = gBusquedas.Resultado
  txtAlbaceaNombre = gBusquedas.Resultado2
End If
End Sub

Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub


Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstadoCivil.SetFocus
End Sub

Private Sub txtDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub txtConyugeCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtConyugeCedula = gBusquedas.Resultado
  txtConyugeNombre = gBusquedas.Resultado2
End If

End Sub



Private Sub txtConyugeNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeTelTrabajo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtConyugeCedula = gBusquedas.Resultado
  txtConyugeNombre = gBusquedas.Resultado2
End If

End Sub


Private Sub txtConyugeTelTrabajo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeTelTrabajoExt.SetFocus
End Sub

Private Sub txtConyugeTelTrabajoExt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConyugeTelCelular.SetFocus
End Sub

