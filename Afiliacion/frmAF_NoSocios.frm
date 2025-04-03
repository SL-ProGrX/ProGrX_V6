VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmAF_NoSocios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "No Socios"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "frmAF_NoSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   8355
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
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
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmAF_NoSocios.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ssTabSubX"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtpNacimiento"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dtpFechaIngreso"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtBoleta"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtEstado"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtNombre"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtApellido2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtApellido1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboSexo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboEstado"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      Begin VB.ComboBox cboEstado 
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
         ItemData        =   "frmAF_NoSocios.frx":686E
         Left            =   6000
         List            =   "frmAF_NoSocios.frx":6884
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1560
         Width           =   1335
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
         ItemData        =   "frmAF_NoSocios.frx":68BF
         Left            =   3360
         List            =   "frmAF_NoSocios.frx":68C9
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtApellido1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   360
         MaxLength       =   15
         TabIndex        =   12
         ToolTipText     =   "Campo para la Cédula de Identidad"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtApellido2 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   11
         ToolTipText     =   "Campo para la Cédula de Identidad"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   10
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtEstado 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   9
         ToolTipText     =   "Campo para la Cédula de Identidad"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtBoleta 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "1"
         Top             =   1560
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFechaIngreso 
         Height          =   315
         Left            =   3360
         TabIndex        =   13
         ToolTipText     =   "Fecha de Ingreso al sistema"
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   140115971
         CurrentDate     =   38899
         MaxDate         =   55153
         MinDate         =   14611
      End
      Begin MSComCtl2.DTPicker dtpNacimiento 
         Height          =   315
         Left            =   6000
         TabIndex        =   14
         ToolTipText     =   "Fecha de Ingreso al sistema"
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   140115971
         CurrentDate     =   36059
      End
      Begin TabDlg.SSTab ssTabSubX 
         Height          =   3375
         Left            =   360
         TabIndex        =   26
         Top             =   2040
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5953
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
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
         TabPicture(0)   =   "frmAF_NoSocios.frx":68E2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label(25)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label10(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label11"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtNotificaciones"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtEmail"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtApartado"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Frame1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Institucion y Centros de Trabajo"
         TabPicture(1)   =   "frmAF_NoSocios.frx":68FE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label8(1)"
         Tab(1).Control(1)=   "Label8(2)"
         Tab(1).Control(2)=   "Label8(3)"
         Tab(1).Control(3)=   "cboInstitucion"
         Tab(1).Control(4)=   "txtDeptDesc"
         Tab(1).Control(5)=   "txtDeptCodigo"
         Tab(1).Control(6)=   "txtSecDesc"
         Tab(1).Control(7)=   "txtSecCodigo"
         Tab(1).ControlCount=   8
         Begin VB.TextBox txtSecCodigo 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   -73560
            MaxLength       =   20
            TabIndex        =   42
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1740
            Width           =   615
         End
         Begin VB.TextBox txtSecDesc 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   -72960
            MaxLength       =   20
            TabIndex        =   41
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1740
            Width           =   4575
         End
         Begin VB.TextBox txtDeptCodigo 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   -73560
            MaxLength       =   20
            TabIndex        =   40
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1380
            Width           =   615
         End
         Begin VB.TextBox txtDeptDesc 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   -72960
            MaxLength       =   20
            TabIndex        =   39
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1380
            Width           =   4575
         End
         Begin VB.ComboBox cboInstitucion 
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
            Left            =   -73560
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   780
            Width           =   5175
         End
         Begin VB.Frame Frame1 
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1605
            Left            =   120
            TabIndex        =   30
            Top             =   420
            Width           =   7095
            Begin VB.TextBox txtDireccion 
               Height          =   1060
               Left            =   3000
               MaxLength       =   100
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   34
               ToolTipText     =   "Dirección exacta Aqui"
               Top             =   360
               Width           =   3975
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
               ItemData        =   "frmAF_NoSocios.frx":691A
               Left            =   840
               List            =   "frmAF_NoSocios.frx":6933
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   360
               Width           =   2055
            End
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
               ItemData        =   "frmAF_NoSocios.frx":697C
               Left            =   840
               List            =   "frmAF_NoSocios.frx":697E
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   720
               Width           =   2055
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
               ItemData        =   "frmAF_NoSocios.frx":6980
               Left            =   840
               List            =   "frmAF_NoSocios.frx":6982
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   1080
               Width           =   2055
            End
            Begin VB.Label Label9 
               Caption         =   "Distrito"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label8 
               Caption         =   "Canton"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label7 
               Caption         =   "Provincia"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.TextBox txtApartado 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   29
            ToolTipText     =   "Campo para la Cédula de Identidad"
            Top             =   2460
            Width           =   5895
         End
         Begin VB.TextBox txtEmail 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1320
            MaxLength       =   45
            TabIndex        =   28
            ToolTipText     =   "Campo para la Cédula de Identidad"
            Top             =   2100
            Width           =   5895
         End
         Begin VB.TextBox txtNotificaciones 
            Height          =   495
            Left            =   1320
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   2820
            Width           =   5895
         End
         Begin VB.Label Label8 
            Caption         =   "Sección"
            Height          =   255
            Index           =   3
            Left            =   -74400
            TabIndex        =   48
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Departam"
            Height          =   255
            Index           =   2
            Left            =   -74400
            TabIndex        =   47
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Institución"
            Height          =   255
            Index           =   1
            Left            =   -74400
            TabIndex        =   46
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Apto. Postal"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2460
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Email"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   2100
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Notificaciones:"
            Height          =   315
            Index           =   25
            Left            =   120
            TabIndex        =   43
            Top             =   2820
            Width           =   1005
         End
      End
      Begin VB.Label Label14 
         Caption         =   "Sexo"
         Height          =   255
         Left            =   2640
         TabIndex        =   25
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Estado Civil"
         Height          =   255
         Left            =   4920
         TabIndex        =   24
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Apellido 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Apellido 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Estado"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Ingreso"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nacimiento"
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "# Boleta"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.TextBox txtCedula 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   4
      ToolTipText     =   "Campo para la Cédula de Identidad"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtCedAlternativa 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4920
      MaxLength       =   15
      TabIndex        =   1
      ToolTipText     =   "Campo para la Cédula de Identidad"
      Top             =   600
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6765
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3245
            MinWidth        =   3245
            Object.ToolTipText     =   "Usuario Ingresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3599
            MinWidth        =   3599
            Object.ToolTipText     =   "Fecha Ingreso"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3245
            MinWidth        =   3245
            Object.ToolTipText     =   "Usuario Modifica"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3599
            MinWidth        =   3599
            Object.ToolTipText     =   "Fecha Modificacion"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar coolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   8355
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   4020
      NewRow1         =   0   'False
      Child2          =   "tlbIngreso"
      MinHeight2      =   330
      Width2          =   2820
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   3825
         _ExtentX        =   6747
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
               Key             =   "Reportes"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "consultar"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblCedula 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cédula Alternativa"
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
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmAF_NoSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEditar As Boolean, vFechaActual As Date
Dim vCedula As String, vSeek As Integer
Dim strSQL As String, vPaso As Boolean

Private Sub sbClearControles()
Dim vControl As Control

For Each vControl In Me
  If TypeOf vControl Is TextBox Then
     vControl.Text = ""
  End If
Next

   StatusBarx.Panels.Item(1) = ""
   StatusBarx.Panels.Item(2) = ""
   StatusBarx.Panels.Item(3) = ""
   StatusBarx.Panels.Item(4) = ""
   

End Sub

Private Sub sbCurrentRecord()
Dim rs As New ADODB.Recordset
Dim vApellido1 As String, vApellido2 As String
Dim vNombre1 As String, vNombre2 As String
Dim i As Integer, vEspacio As Integer

On Error Resume Next

strSQL = "Select S.*,I.descripcion as DescInst,D.descripcion as DescDept" _
       & ",X.descripcion as DescSec" _
       & ",rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & " From socios S inner join Instituciones I on S.cod_institucion = I.cod_institucion" _
       & " inner join AFDepartamentos D on S.cod_institucion = D.cod_institucion and S.cod_departamento = D.cod_departamento" _
       & " inner join AFSecciones X on S.cod_institucion = X.cod_institucion" _
       & " left join Provincias Prov on S.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on S.Provincia = Cant.Provincia and S.Canton = Cant.Canton" _
       & " left join Distritos Dist on S.Provincia = Dist.Provincia and S.Canton = Dist.Canton and S.distrito = Dist.distrito" _
       & " where cedula='" & Trim(txtCedula) & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   vEditar = True
   Call sbToolBar(Me.tlb, "activo")
   Call RefrescaTags(Me)
   Call sbLockControles("L")
   
   vCedula = Trim(txtCedula)
          
   vEspacio = 1
   For i = 1 To Len(Trim(rs!Nombre))
     If Mid(Trim(rs!Nombre), i, 1) <> " " Then
        Select Case vEspacio
         Case 1
          vApellido1 = vApellido1 & Mid(Trim(rs!Nombre), i, 1)
         Case 2
          vApellido2 = vApellido2 & Mid(Trim(rs!Nombre), i, 1)
         Case 3
          vNombre1 = vNombre1 & Mid(Trim(rs!Nombre), i, 1)
         Case Is >= 4
          vNombre2 = vNombre2 & Mid(Trim(rs!Nombre), i, 1)
        End Select
     Else
        vEspacio = vEspacio + 1
     End If
   Next i
     
   txtCedAlternativa = rs!cedular & ""
   
   txtApellido1 = vApellido1
   txtApellido2 = vApellido2
   txtNombre = vNombre1 & " " & vNombre2
     
   txtBoleta = rs!id_Boleta_AF & ""
     
   Select Case rs!EstadoActual
    Case "S"
      txtEstado = "SOCIO"
      tlb.Buttons.Item(2).Enabled = True
    Case "N"
      txtEstado = "NO SOCIO"
      tlb.Buttons.Item(2).Enabled = True
    Case "A"
      txtEstado = "Ren.Asociacion"
      tlb.Buttons.Item(2).Enabled = False
    Case "P"
      txtEstado = "Ren.Patrono"
      tlb.Buttons.Item(2).Enabled = False
   End Select
     
   dtpFechaIngreso = rs!FechaIngreso
   dtpNacimiento = rs!fecha_nac
   cboSexo = IIf(rs!sexo = "M", "Masculino", "Femenino")
     
   cboEstado = fxEstadoCivil(rs!estadoCivil)
     
     
   Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
   Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
   Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
     
   cboDistrito.ToolTipText = Trim(rs!distrito) & ""
   
   
   txtDireccion = Trim(rs!Direccion) & ""
   txtEMail = Trim(rs!AF_Email) & ""
   txtApartado = Trim(rs!apto) & ""
   
   txtNotificaciones.Text = Trim(rs!Notificaciones & "")
   
   cboInstitucion.Text = Trim(rs!DescInst)
   
   txtDeptCodigo = rs!cod_departamento & ""
   txtDeptDesc = Trim(rs!descDept & "")
   
   txtSecCodigo = rs!cod_seccion & ""
   txtSecDesc = Trim(rs!DescSec & "")
   
 
   txtCedula.SetFocus
   
   StatusBarx.Panels.Item(1) = rs!reg_user & ""
   StatusBarx.Panels.Item(2) = rs!reg_fecha & ""
   StatusBarx.Panels.Item(3) = rs!ActualizaUser & ""
   StatusBarx.Panels.Item(4) = rs!ActualizaFecha & ""
  
   
  Else
   
   If vEditar = True Or txtApellido1.Enabled = False Then
        vEditar = False
        Call sbToolBar(Me.tlb, "nuevo")
        Call RefrescaTags(Me)
        Call sbClearControles
        Call sbLockControles("L")
        txtCedula.SetFocus
   End If

End If
rs.Close


End Sub

Private Sub sbInsertAhorro()
Dim rs As New ADODB.Recordset

strSQL = "select coalesce(count(*),0) as Existe from ahorro_consolidado" _
       & " where cedula = '" & Trim(txtCedula) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    strSQL = "Insert into Ahorro_Consolidado(Cedula,Aporte,Ahorro,Extra,Capitaliza," _
           & "FecAporte,FecAhorro,FecExtra,FecCapitaliza,AportAnt,AhorroAnt) Values(" _
           & "'" & Trim(txtCedula) & "',0,0,0,0,dbo.MyGetdate(),dbo.MyGetdate(),dbo.MyGetdate(),dbo.MyGetdate(),0,0)"
    Call ConectionExecute(strSQL)
End If
rs.Close

End Sub

Sub sbLockControles(vModo As String)
Dim vControl As Control

For Each vControl In Me
  If (TypeOf vControl Is TextBox And vControl.Name <> "txtCedula" And vControl.Name <> _
     "txtEstado") Or TypeOf vControl Is DTPicker Or TypeOf vControl Is ComboBox Then
        If vModo = "L" Then
           If vControl.Name = "txtNombre" Then
            vControl.Locked = True
           Else
            vControl.Enabled = False
           End If
        Else
           If vControl.Name = "txtNombre" Then
            vControl.Locked = False
           Else
            vControl.Enabled = True
           End If
        End If
  End If
Next

dtpFechaIngreso.Enabled = False

End Sub


Private Sub sbDeleteRecord()
Dim i As Integer
On Error GoTo vError

If Trim(txtCedula) <> vCedula Then
   MsgBox "Ha modificado la cédula", vbExclamation
   Exit Sub
End If

i = MsgBox("Esta Seguro Que Desea Borrar Este Socio?", vbYesNo)
If i = vbYes Then
  vEditar = False
  strSQL = "delete Socios where Cedula='" & Trim(txtCedula) & "' and estadoactual = 'N'"
  Call ConectionExecute(strSQL)
  Call Bitacora("Borra", "No Socio - Cedula: " & Trim(txtCedula))
  
  Call sbClearControles
  Call sbToolBar(Me.tlb, "nuevo")
  Call RefrescaTags(Me)
  txtCedula.SetFocus
End If

Exit Sub

vError:
 MsgBox err.Description, vbCritical

End Sub

Private Function fxValida() As Boolean
Dim rs As New ADODB.Recordset, i As Integer
Dim vMensaje As String

vMensaje = ""

If txtEstado <> "NO SOCIO" Then vMensaje = vMensaje & " -  La persona a modificar no está disponible como > NO SOCIO < [verifique]" & vbCrLf
If Trim(txtCedula) = "" Then vMensaje = vMensaje & " - Falta el Numero de Cedula" & vbCrLf
If Trim(txtApellido1) = "" Then vMensaje = vMensaje & " - Falta el Apellido 1" & vbCrLf
If Trim(txtApellido2) = "" Then vMensaje = vMensaje & " - Falta el Apellido 2" & vbCrLf
If Trim(txtNombre) = "" Then vMensaje = vMensaje & " - Falta el Nombre" & vbCrLf
If Trim(cboSexo) = "" Then vMensaje = vMensaje & " - No se especificó el Sexo" & vbCrLf
If Trim(cboEstado) = "" Then vMensaje = vMensaje & " - No se especificó el Estado Civil" & vbCrLf
If Trim(cboProvincia.Text) = "" Then vMensaje = vMensaje & " - No se especificó la Provincia" & vbCrLf
If Trim(cboCanton.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Cantón" & vbCrLf
If Trim(txtDireccion) = "" Then vMensaje = vMensaje & " - No se especificó la Dirección" & vbCrLf
If Trim(txtBoleta) = "" Then vMensaje = vMensaje & " - Indique el Número de Boleta" & vbCrLf

If Not IsNumeric(txtBoleta) Then vMensaje = vMensaje & " - Número de Boleta no es válido" & vbCrLf

If Not vEditar Then

'- Verifica que no existe otra persona con el mismo nombre -> Solo Adventencia
    'Filtra nombre
    strSQL = ""
    For i = 1 To Len(txtNombre)
      If Mid(txtNombre, i, 1) <> " " Then
         strSQL = strSQL & Mid(txtNombre, i, 1)
      Else
         Exit For
      End If
    Next i
  
  strSQL = "select coalesce(count(*),0) as Existe from socios where nombre like '" & Trim(txtApellido1.Text) _
         & " " & Trim(txtApellido2.Text) & " " & strSQL & "%'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe > 0 Then
   i = MsgBox("Existe otra persona con el mismo nombre, está seguro de que sea continuar con el registro?", vbYesNo)
   If i = vbNo Then vMensaje = vMensaje & " - Existe otra persona con el mismo nombre, verifique..." & vbCrLf
  End If
  rs.Close

End If

If Len(vMensaje) = 0 Then
  fxValida = True
Else
  fxValida = False
  MsgBox vMensaje, vbExclamation
End If

End Function

Private Function fxVerificaCambioInst(xCedula As String, vInst As Integer) As Boolean
Dim vSQL As String, rsX As New ADODB.Recordset

fxVerificaCambioInst = True

vSQL = "select A.aporte,S.cod_institucion" _
     & " from ahorro_consolidado A inner join Socios S on A.cedula = S.cedula" _
     & " where S.cedula = '" & xCedula & "'"
rsX.Open vSQL, glogon.Conection, adOpenStatic
If Not rsX.EOF And Not rsX.BOF Then
   If rsX!cod_institucion <> vInst And rsX!Aporte > 0 Then
      fxVerificaCambioInst = False
   End If
End If

rsX.Close

End Function


Private Sub sbSaveRecord()
Dim vEstadoCivil As String, vActiva As Boolean
Dim i As Integer, vValor(3) As Integer


On Error GoTo vError

strSQL = ""
vActiva = False

'Valores por Defecto, analizar si se pasan a tabla de parametros
vValor(0) = 1 'Promotor
vValor(1) = 1 'Profesion
vValor(2) = 1 'Sector
vValor(3) = 1 'Nombramiento


If Not fxValida Then
  Exit Sub
End If

If vEditar = True Then
  
 If Trim(txtCedula) <> vCedula Then
   MsgBox "Ha modificado la cédula", vbExclamation
   Exit Sub
 End If
 'Verifica que pueda cambiar de institucion
 If fxVerificaCambioInst(txtCedula, cboInstitucion.ItemData(cboInstitucion.ListIndex)) = False Then
   MsgBox "No se puede cambiar la institución a esta persona porque ya tiene aportes registrados: debe liquidar primero", vbExclamation
   Exit Sub
 End If
 
End If

vEstadoCivil = fxEstadoCivil(cboEstado.Text)

strSQL = ""
For i = 1 To Len(txtApellido1)
  If Mid(txtApellido1, i, 1) <> " " Then
     strSQL = strSQL & Mid(txtApellido1, i, 1)
  End If
Next i
txtApellido1 = strSQL

strSQL = ""
For i = 1 To Len(txtApellido2)
  If Mid(txtApellido2, i, 1) <> " " Then
     strSQL = strSQL & Mid(txtApellido2, i, 1)
  End If
Next i
txtApellido2 = strSQL

If Not vEditar Then
   vActiva = True
   strSQL = "Insert Socios(Cedula,Nombre,EstadoActual,FechaIngreso,Fecha_Nac," _
          & "Sexo,EstadoCivil,Provincia,Canton,Distrito,Direccion,Af_Email,Apto," _
          & "Cod_sector,cod_profesion,id_promotor,EstadoLaboral,Ultimo_Estado,Cod_Institucion" _
          & ",cod_departamento,cod_seccion,id_boleta_af,cedulaR,af_npagos,hijos,prideduc,reg_user,reg_fecha,Notificaciones)" _
          & " Values('" & Trim(txtCedula) & "','" & Trim(txtApellido1) & " " & Trim(txtApellido2) & " " & Trim(txtNombre) & "'," _
          & "'N','" & Format(dtpFechaIngreso, "yyyy/mm/dd") & "','" & Format(dtpNacimiento, "yyyy/mm/dd") & "','" _
          & IIf(Trim(cboSexo) = "Masculino", "M", "F") & "','" & vEstadoCivil & "'," & cboProvincia.ItemData(cboProvincia.ListIndex) & "," & cboCanton.ItemData(cboCanton.ListIndex) & ",'" _
          & Format(cboDistrito.ItemData(cboDistrito.ListIndex), "###000") & "','" & Trim(txtDireccion) & "','" & Trim(txtEMail) & "','" & Trim(txtApartado) & "'," & vValor(2) & "," _
          & vValor(1) & "," & vValor(0) & "," & vValor(3) & ",'N'," & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ",'" & txtDeptCodigo & "','" _
          & txtSecCodigo & "'," & txtBoleta & ",'" & txtCedAlternativa & "',2,0," & fxgPrimerDeduccionIng(cboInstitucion.ItemData(cboInstitucion.ListIndex)) _
          & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'" & txtNotificaciones.Text & "')"
   Call ConectionExecute(strSQL)
   
    Call sbInsertAhorro
    Call Bitacora("Registra", "No Socio - Cedula: " & Trim(txtCedula))

Else
   vActiva = False
   strSQL = "Update Socios set Nombre = '" & Trim(txtApellido1) & " " & Trim(txtApellido2) & " " & Trim(txtNombre) _
          & "',FechaIngreso = '" & Format(dtpFechaIngreso, "yyyy/mm/dd") & "',Fecha_Nac='" & Format(dtpNacimiento, "yyyy/mm/dd") _
          & "',Sexo = '" & Mid(cboSexo.Text, 1, 1) & "',EstadoCivil = '" & vEstadoCivil & "',Provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) _
          & ",Canton = " & cboCanton.ItemData(cboCanton.ListIndex) & ",Distrito = '" & Format(cboDistrito.ItemData(cboDistrito.ListIndex), "###000") & "',Direccion='" & Trim(txtDireccion) _
          & "',Af_Email = '" & Trim(txtEMail) & "', hijos = 0,Apto = '" & Trim(txtApartado) & "',af_npagos = 2" _
          & ",cod_sector = " & vValor(2) & ",Id_promotor = " & vValor(0) _
          & ",cod_profesion = " & vValor(1) & ", EstadoLaboral = " & vValor(3) _
          & ",cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ",cod_departamento = '" & txtDeptCodigo _
          & "',cod_seccion = '" & txtSecCodigo & "',id_boleta_af = " & txtBoleta & ",cedular = '" & txtCedAlternativa _
          & "',Notificaciones = '" & txtNotificaciones.Text _
          & "' Where Cedula='" & vCedula & "'"
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Modifica", "No Socio - Cedula: " & vCedula)

End If


vCedula = Trim(txtCedula)
vEditar = True

Call sbToolBar(Me.tlb, "activo")
Call RefrescaTags(Me)
Call sbLockControles("L")

MsgBox "Información guardada satisfactoriamente...", vbInformation
txtCedula.SetFocus


Exit Sub

vError:
 MsgBox err.Description, vbCritical
 
End Sub


Private Sub cboCanton_Click()
Dim vSQL As String

If vPaso Then Exit Sub

vSQL = " where Provincia = " & cboProvincia.ItemData(cboProvincia.ListIndex) _
     & " and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) & "' order by descripcion"

Call sbCargaCbo(cboDistrito, "distritos", vSQL)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistrito.AddItem " "
cboDistrito.Text = " "

End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub

Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus
End Sub


Private Sub cboInstitucion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptCodigo.SetFocus
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

Private Sub cboSexo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub

Private Sub dtpFechaIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpNacimiento.SetFocus
End Sub


Private Sub dtpNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtBoleta.SetFocus
End Sub


Private Sub Form_Load()
Dim rs As New ADODB.Recordset

vModulo = 1
Call Formularios(Me)

SSTab1.Tab = 0

vEditar = False
vFechaActual = Format(fxFechaServidor, "dd/mm/yyyy")

dtpFechaIngreso = vFechaActual

Call sbCargaCbo(cboInstitucion, "instituciones")

vPaso = True
Call sbCargaCbo(cboProvincia, "provincias")
vPaso = False

Call sbToolBarIconos(Me.tlb, False)
Call sbToolBar(Me.tlb, "nuevo")

Call sbLockControles("L")
Call RefrescaTags(Me)


End Sub


Private Sub sbLimpiaDatos()

'Inicializa
ssTabSubX.Tab = 0

txtEstado.Text = "NO SOCIO"
dtpFechaIngreso.Value = vFechaActual
dtpNacimiento.Value = vFechaActual

txtBoleta = 0

cboSexo.Text = "Masculino"
cboEstado.Text = "Soltero"

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "nuevo"
    vEditar = False
    Call sbToolBar(Me.tlb, "edicion")
    Call sbClearControles
    Call sbLockControles("U")
    Call sbLimpiaDatos
    txtCedula.SetFocus
    
  Case "editar"
    If Trim(txtCedula) <> vCedula Then
     MsgBox "Ha modificado la cédula", vbExclamation
     Exit Sub
    End If
    
    vEditar = True
    vCedula = Trim(txtCedula)
    Call sbToolBar(Me.tlb, "edicion")
    Call sbLockControles("U")
    txtApellido1.SetFocus
        
  Case "borrar"
    Call sbDeleteRecord
        
  Case "guardar"
    Call sbSaveRecord
    
  Case "deshacer"
    vEditar = False
    Call sbToolBar(Me.tlb, "nuevo")
    Call RefrescaTags(Me)
    Call sbClearControles
    Call sbLockControles("L")
    txtCedula.SetFocus
    
  Case "consultar"
    Select Case vSeek
      Case 1, 2
       gBusquedas.Resultado = Trim(txtCedula)
       txtCedula = ""
       vCedula = ""
       gBusquedas.Convertir = "N"
       
       If vSeek = 1 Then
        gBusquedas.Columna = "Cedula"
        gBusquedas.Orden = "Cedula"
       Else
        gBusquedas.Columna = "Nombre"
        gBusquedas.Orden = "Nombre"
       End If
       
       gBusquedas.Consulta = "Select Cedula,Nombre From Socios"
   
       frmBusquedas.Show vbModal
   
       txtCedula = Trim(gBusquedas.Resultado)
       txtCedula_LostFocus
       
      Case 3
       If cboProvincia.Text = "" Then Exit Sub
       gBusquedas.Resultado = ""
       gBusquedas.Resultado2 = ""
   
       gBusquedas.Convertir = "N"
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "Descripcion"
       gBusquedas.Consulta = "Select Canton,Descripcion From Cantones"
       gBusquedas.Filtro = "And Provincia =" & cboProvincia.ItemData(cboProvincia.ListIndex)
   
       frmBusquedas.Show vbModal
   
'       txtCodigoCanton = Trim(gBusquedas.Resultado)
'       txtCanton = Trim(gBusquedas.Resultado2)
    End Select
    
End Select

End Sub

Private Sub txtApartado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotificaciones.SetFocus
End Sub


Private Sub txtApellido1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApellido2.SetFocus
End Sub

Private Sub txtApellido1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtApellido2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtApellido2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtBoleta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboSexo.SetFocus
End Sub

Private Sub txtCedula_GotFocus()
vSeek = 1
SSTab1.Tab = 0
ssTabSubX.Tab = 0

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtCedula)
   txtCedula = ""
   vCedula = ""
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "Select Cedula,Nombre From Socios"
   
   frmBusquedas.Show vbModal
   
   txtCedula = Trim(gBusquedas.Resultado)
   txtCedula_LostFocus
End If
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case 48 To 57, 8
 Case vbKeyReturn
  txtCedula_LostFocus
 Case Else
  KeyAscii = 0
End Select

End Sub


Private Sub txtCedula_LostFocus()

If Trim(txtCedula) = "" Then
  If vEditar = True Then
     vEditar = False
     Call sbToolBar(Me.tlb, "nuevo")
     Call RefrescaTags(Me)
     Call sbClearControles
     Call sbLockControles("L")
  End If
Else
  If vEditar = False Or (vEditar = True And vCedula <> Trim(txtCedula)) Then
     Call sbCurrentRecord
  End If
End If

End Sub


Private Sub txtDeptCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_departamento"
  gBusquedas.Orden = "cod_departamento"
  gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
  gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDeptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecCodigo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
  gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then txtEMail.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApartado.SetFocus
End Sub

Private Sub txtNombre_GotFocus()
vSeek = 2
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpNacimiento.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtCedula)
   txtCedula = ""
   vCedula = ""
      
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "Select Cedula,Nombre From Socios"
   
   frmBusquedas.Show vbModal
   
   txtCedula = Trim(gBusquedas.Resultado)
   txtCedula_LostFocus
End If

End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtNotificaciones_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    ssTabSubX.Tab = 1
    cboInstitucion.SetFocus
End If
End Sub




Private Sub txtSecCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecDesc.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cod_seccion"
  gBusquedas.Orden = "cod_seccion"
  gBusquedas.Consulta = "select cod_seccion,descripcion from AFSecciones"
  gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
            & " and cod_departamento = '" & txtDeptCodigo & "'"
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtSecDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_seccion,descripcion from AFSecciones"
  gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
            & " and cod_departamento = '" & txtDeptCodigo & "'"
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If

End Sub


