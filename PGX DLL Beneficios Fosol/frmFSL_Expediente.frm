VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Expediente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expediente FOSOL"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   20
      Top             =   1200
      Width           =   6255
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
      TabIndex        =   19
      Top             =   1200
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
      TabIndex        =   1
      ToolTipText     =   "Número de Tramite"
      Top             =   600
      Width           =   1695
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
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6135
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
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
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmFSL_Expediente.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Line3(1)"
      Tab(0).Control(1)=   "Label7(15)"
      Tab(0).Control(2)=   "Label7(14)"
      Tab(0).Control(3)=   "Label7(13)"
      Tab(0).Control(4)=   "Label3(2)"
      Tab(0).Control(5)=   "Label2(0)"
      Tab(0).Control(6)=   "Label3(0)"
      Tab(0).Control(7)=   "Label3(6)"
      Tab(0).Control(8)=   "Label3(7)"
      Tab(0).Control(9)=   "Line3(3)"
      Tab(0).Control(10)=   "Label2(2)"
      Tab(0).Control(11)=   "Label3(8)"
      Tab(0).Control(12)=   "Label7(0)"
      Tab(0).Control(13)=   "Label3(9)"
      Tab(0).Control(14)=   "Label3(10)"
      Tab(0).Control(15)=   "Label7(1)"
      Tab(0).Control(16)=   "Label3(15)"
      Tab(0).Control(17)=   "dtpEnfermedad"
      Tab(0).Control(18)=   "cboEnfermedad"
      Tab(0).Control(19)=   "cboCausa"
      Tab(0).Control(20)=   "cboTipo"
      Tab(0).Control(21)=   "txtEnfermedadNotas"
      Tab(0).Control(22)=   "txtPresentaCedula"
      Tab(0).Control(23)=   "txtPresentaNombre"
      Tab(0).Control(24)=   "txtPresentaNotas"
      Tab(0).Control(25)=   "txtNotas"
      Tab(0).Control(26)=   "cboRefTipoDoc"
      Tab(0).Control(27)=   "txtRefNumero"
      Tab(0).Control(28)=   "dtpRefFecha"
      Tab(0).Control(29)=   "cboComite"
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Requisitos"
      TabPicture(1)   =   "frmFSL_Expediente.frx":0142
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(14)"
      Tab(1).Control(1)=   "lswRequisitos"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Operaciones"
      TabPicture(2)   =   "frmFSL_Expediente.frx":01FB
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Line3(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label7(12)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label7(9)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label7(4)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label7(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblTD_Label_02"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lblTD_Label_01"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "vgCreditos"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtTotalAplicado"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtLiquidacionMonto"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtTotalDisponible"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtTD_Destino"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtTD_Texto_02"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtTD_Texto_01"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Resolución"
      TabPicture(3)   =   "frmFSL_Expediente.frx":0308
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label3(11)"
      Tab(3).Control(1)=   "Label3(12)"
      Tab(3).Control(2)=   "Label3(13)"
      Tab(3).Control(3)=   "Line3(4)"
      Tab(3).Control(4)=   "Label3(3)"
      Tab(3).Control(5)=   "Label3(4)"
      Tab(3).Control(6)=   "Label3(5)"
      Tab(3).Control(7)=   "imgRequisitos"
      Tab(3).Control(8)=   "imgTiempoPresentacion"
      Tab(3).Control(9)=   "imgExpedientesActivos"
      Tab(3).Control(10)=   "tlbResolucion"
      Tab(3).Control(11)=   "lswComite"
      Tab(3).Control(12)=   "txtResolucionNotas"
      Tab(3).Control(13)=   "cboResolucion"
      Tab(3).Control(14)=   "fraValidaMiembro"
      Tab(3).ControlCount=   15
      TabCaption(4)   =   "Gestiones"
      TabPicture(4)   =   "frmFSL_Expediente.frx":0409
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "vgGestiones"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Apelaciones"
      TabPicture(5)   =   "frmFSL_Expediente.frx":0505
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "vgApelaciones"
      Tab(5).ControlCount=   1
      Begin VB.Frame fraValidaMiembro 
         Caption         =   "Validación del Miembro de Comité"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -73440
         TabIndex        =   63
         Top             =   1680
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
            TabIndex        =   68
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
            TabIndex        =   67
            Top             =   1080
            Width           =   2175
         End
         Begin MSComctlLib.Toolbar tlbValidaMiembro 
            Height          =   360
            Left            =   3000
            TabIndex        =   66
            Top             =   1680
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   635
            ButtonWidth     =   1826
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
            MousePointer    =   1
         End
         Begin VB.Label lblMiembro 
            Alignment       =   1  'Right Justify
            Caption         =   "Miembro.: "
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   69
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
               Name            =   "Arial Narrow"
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
            TabIndex        =   65
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Usuario .: "
            BeginProperty Font 
               Name            =   "Arial Narrow"
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
            TabIndex        =   64
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.TextBox txtTD_Texto_01 
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
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   5160
         Width           =   1935
      End
      Begin VB.TextBox txtTD_Texto_02 
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
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   5520
         Width           =   1935
      End
      Begin VB.TextBox txtTD_Destino 
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
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   4800
         Width           =   1935
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
         Left            =   -70320
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   5400
         Width           =   2895
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
         Height          =   2775
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   1200
         Width           =   4335
      End
      Begin VB.ComboBox cboComite 
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
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   2640
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker dtpRefFecha 
         Height          =   315
         Left            =   -73320
         TabIndex        =   42
         Top             =   3360
         Width           =   1935
         _ExtentX        =   3413
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
         CalendarBackColor=   12648447
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   162004995
         CurrentDate     =   41732
      End
      Begin VB.TextBox txtRefNumero 
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
         Left            =   -73320
         TabIndex        =   40
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ComboBox cboRefTipoDoc 
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
         Left            =   -73320
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtTotalDisponible 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox txtLiquidacionMonto 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   5520
         Width           =   1935
      End
      Begin VB.TextBox txtTotalAplicado 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   5160
         Width           =   1935
      End
      Begin VB.TextBox txtNotas 
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
         Height          =   1455
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   4560
         Width           =   4695
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
         Height          =   615
         Left            =   -73800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1320
         Width           =   8295
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
         Left            =   -70800
         TabIndex        =   24
         Top             =   960
         Width           =   5055
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
         Left            =   -73800
         TabIndex        =   23
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtEnfermedadNotas 
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
         Height          =   1455
         Left            =   -70080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   4560
         Width           =   4575
      End
      Begin VB.ComboBox cboTipo 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3000
         Width           =   4335
      End
      Begin VB.ComboBox cboCausa 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3360
         Width           =   4335
      End
      Begin VB.ComboBox cboEnfermedad 
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
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3840
         Width           =   4335
      End
      Begin MSComctlLib.ListView lswComite 
         Height          =   3735
         Left            =   -70320
         TabIndex        =   46
         Top             =   1200
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
         Left            =   -67080
         TabIndex        =   49
         Top             =   5400
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   635
         ButtonWidth     =   2090
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
      Begin MSComctlLib.ListView lswRequisitos 
         Height          =   5535
         Left            =   -72480
         TabIndex        =   52
         Top             =   480
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   9763
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   1
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
            Text            =   "Requisito"
            Object.Width           =   8890
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Opcional ?"
            Object.Width           =   2540
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vgCreditos 
         Height          =   3975
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   9135
         _Version        =   524288
         _ExtentX        =   16113
         _ExtentY        =   7011
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
         MaxCols         =   13
         SpreadDesigner  =   "frmFSL_Expediente.frx":0633
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vgGestiones 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   55
         Top             =   480
         Width           =   9255
         _Version        =   524288
         _ExtentX        =   16325
         _ExtentY        =   9551
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
         MaxCols         =   4
         SpreadDesigner  =   "frmFSL_Expediente.frx":104A
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vgApelaciones 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   56
         Top             =   480
         Width           =   9255
         _Version        =   524288
         _ExtentX        =   16325
         _ExtentY        =   9551
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
         SpreadDesigner  =   "frmFSL_Expediente.frx":1696
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComCtl2.DTPicker dtpEnfermedad 
         Height          =   315
         Left            =   -73320
         TabIndex        =   74
         Top             =   3840
         Width           =   1935
         _ExtentX        =   3413
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
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   16777215
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   161939459
         CurrentDate     =   41732
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Diagnóstico Enfermedad"
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
         Index           =   15
         Left            =   -74760
         TabIndex        =   73
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Image imgExpedientesActivos 
         Height          =   255
         Left            =   -71280
         Top             =   4800
         Width           =   255
      End
      Begin VB.Image imgTiempoPresentacion 
         Height          =   255
         Left            =   -71280
         Top             =   4440
         Width           =   255
      End
      Begin VB.Image imgRequisitos 
         Height          =   255
         Left            =   -71280
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cumple No. Expedientes Activos.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   72
         Top             =   4800
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cumple Tiempo de Presentación.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   71
         Top             =   4440
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cumple Requisitos .:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   70
         Top             =   4080
         Width           =   3375
      End
      Begin VB.Label lblTD_Label_01 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
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
         Height          =   255
         Left            =   4680
         TabIndex        =   62
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label lblTD_Label_02 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
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
         Height          =   255
         Left            =   4800
         TabIndex        =   61
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   60
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Cumplimiento con Requisitos.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   53
         Top             =   480
         Width           =   2415
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   -74760
         X2              =   -65520
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   " Resolución.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   -72120
         TabIndex        =   50
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Miembros del Comité.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -70320
         TabIndex        =   48
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Notas de la Resolución.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   47
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Comité...:"
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
         Left            =   -70920
         TabIndex        =   43
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Ref. No.:"
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
         Index           =   10
         Left            =   -74760
         TabIndex        =   41
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Ref. Fecha.:"
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
         Index           =   9
         Left            =   -74760
         TabIndex        =   39
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Referencia...:"
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
         Left            =   -74760
         TabIndex        =   37
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Notas de la enfermedad/causa.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -70080
         TabIndex        =   36
         Top             =   4320
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   "Disponible"
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
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   35
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "(Liquidación)"
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
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   34
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Total Aplicado"
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
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   33
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Datos del Expediente"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   28
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   -74880
         X2              =   -65640
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label3 
         Caption         =   "Datos Adicionales"
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
         Left            =   -74760
         TabIndex        =   27
         Top             =   1320
         Width           =   975
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
         Left            =   -71760
         TabIndex        =   25
         Top             =   960
         Width           =   855
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
         Left            =   -74760
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Datos del Solicitante.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   21
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Notas del Expediente.:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   10
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Plan...:"
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
         Index           =   13
         Left            =   -70920
         TabIndex        =   9
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Causa..:"
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
         Index           =   14
         Left            =   -70800
         TabIndex        =   8
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Enfermedad ...:"
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
         Index           =   15
         Left            =   -71160
         TabIndex        =   7
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -74760
         X2              =   -65520
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   120
         X2              =   9360
         Y1              =   4680
         Y2              =   4680
      End
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   9885
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   2955
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   2520
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   3150
         TabIndex        =   13
         Top             =   30
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   582
         ButtonWidth     =   2937
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Gestiones"
               Key             =   "Gestiones"
               Object.ToolTipText     =   "Registro de Gestiones Realizadas"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Apelación"
               Key             =   "Apelacion"
               Object.ToolTipText     =   "Apelación de la Resolución"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar (FOSOL)"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplicación de la Resolución"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   12
         Top             =   30
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
               Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
               Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
               Object.ToolTipText     =   "Boleta de Registro"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
               Object.ToolTipText     =   "Ayuda General"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cierra esta ventana"
               Object.Tag             =   "1"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":1DBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":1ED9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":1FFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":20FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":21F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":2326
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":2450
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":256E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":2694
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":27BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1572865
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   8760
      Top             =   600
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
            Picture         =   "frmFSL_Expediente.frx":28B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":9119
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":9212
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   7875
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   9960
      Y1              =   1080
      Y2              =   1080
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
      TabIndex        =   17
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Expediente"
      BeginProperty Font 
         Name            =   "Arial Narrow"
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
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image imgEstado 
      Height          =   255
      Left            =   3720
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial Narrow"
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
      TabIndex        =   15
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmFSL_Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje        As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita          As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso           As Boolean 'Control de Activacion de Controles en proceso de carga
Dim vScroll         As Boolean

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

fxValida = True
vMensaje = ""

If Not vEdita Then
    strSQL = "select dbo.fxFSL_ExpedienteValidaRegistro('" & txtCedula.Text & "','" & SIFGlobal.fxSIFCodText(cboTipo.Text) _
           & "','" & SIFGlobal.fxSIFCodText(cboCausa.Text) & "',0) as 'Cumple'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If rs!Cumple = 0 Then
        vMensaje = vMensaje & vbCrLf & "- El caso ya fue presentado anteriormente...verifique!"
    End If
    
End If
    
If Len(txtNotas.Text) <= 10 Then vMensaje = vMensaje & vbCrLf & "- Indique una Nota válida!"


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub cboCausa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEnfermedad.SetFocus

End Sub

Private Sub cboComite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus

End Sub

Private Sub cboEnfermedad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEnfermedadNotas.SetFocus

End Sub

Private Sub cboRefTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRefNumero.SetFocus

End Sub

Private Sub cboTipo_Click()
Dim strSQL As String

cboCausa.Clear

If vPaso Then Exit Sub
If cboTipo.ListCount = 0 Then Exit Sub

strSQL = "select rtrim(COD_CAUSA) + ' - ' + descripcion as 'ItmX'" _
       & " from FSL_PLANES_CAUSAS" _
       & " where COD_PLAN = '" & SIFGlobal.fxSIFCodText(cboTipo.Text) _
       & "' order by COD_CAUSA"
Call sbLlenaCbo(cboCausa, strSQL, False, False)

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCausa.SetFocus
End Sub



Private Sub dtpRefFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboComite.SetFocus

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
ssTab.Tab = 0
If txtExpediente.Text = "" Then txtExpediente.Text = "0"
If FlatScrollBar.Tag = "" Then FlatScrollBar.Tag = 0

strSQL = "select Top 1 cod_Expediente from FSL_Expedientes"

If FlatScrollBar.Value > CLng(FlatScrollBar.Tag) Then
   strSQL = strSQL & " where cod_Expediente > " & txtExpediente & " order by cod_Expediente asc"
Else
   strSQL = strSQL & " where cod_Expediente < " & txtExpediente & " order by cod_Expediente desc"
End If

FlatScrollBar.Tag = FlatScrollBar.Value

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
  txtExpediente.Text = rs!COD_EXPEDIENTE
  Call sbConsulta
End If
rs.Close

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub Form_Load()
Dim strSQL As String

 vModulo = 22
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "activo")

 Call Formularios(Me)
 Call RefrescaTags(Me)
 

Call sbPantallaInicial
Call sbLimpiaPantalla


End Sub

Private Sub sbPantallaInicial()
Dim strSQL As String

Me.MousePointer = vbHourglass
vPaso = True

strSQL = "select Rtrim(COD_PLAN) + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  FSL_PLANES where activo = 1"
Call sbLlenaCbo(cboTipo, strSQL, False)

strSQL = "select Rtrim(COD_COMITE) + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  FSL_COMITES where activo = 1"
Call sbLlenaCbo(cboComite, strSQL, False)

strSQL = "select Rtrim(COD_ENFERMEDAD) + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  FSL_TIPOS_ENFERMEDADES where activa = 1"
Call sbLlenaCbo(cboEnfermedad, strSQL, False)

cboRefTipoDoc.Clear
cboRefTipoDoc.AddItem "Accion de personal"
cboRefTipoDoc.AddItem "Acta defunción"
cboRefTipoDoc.Text = "Accion de personal"

cboResolucion.Clear
cboResolucion.AddItem "Aprobado"
cboResolucion.AddItem "Rechazado"
cboResolucion.Text = "Aprobado"


vPaso = False
Call cboTipo_Click

Me.MousePointer = vbDefault

End Sub

Private Sub sbLimpiaPantalla()
Dim strSQL As String

Me.MousePointer = vbHourglass
vPaso = True
 tlbAux.Buttons.Item(1).Enabled = False 'Gestiones
 tlbAux.Buttons.Item(2).Enabled = False 'Apelacion
 tlbAux.Buttons.Item(4).Enabled = False 'Aplicar

 ssTab.TabEnabled(1) = False
 ssTab.TabEnabled(2) = False
 ssTab.TabEnabled(3) = False
 ssTab.TabEnabled(4) = False
 ssTab.TabEnabled(5) = False

 ssTab.Tab = 0

Set imgEstado.Picture = ImageList1.ListImages.Item(6).Picture
Set imgRequisitos.Picture = ImageList1.ListImages.Item(6).Picture
Set imgExpedientesActivos.Picture = ImageList1.ListImages.Item(6).Picture
Set imgTiempoPresentacion.Picture = ImageList1.ListImages.Item(6).Picture

txtExpediente.Text = ""
txtEstado.Text = "Pendiente"

txtCedula.Text = ""
txtNombre.Text = ""

txtPresentaCedula.Text = ""
txtPresentaNombre.Text = ""
txtPresentaNotas.Text = ""

txtNotas.Text = ""
txtEnfermedadNotas.Text = ""

 dtpRefFecha.Value = fxFechaServidor
 dtpEnfermedad.Value = dtpRefFecha.Value
 txtRefNumero.Text = ""
 
 StatusBarX.Panels(1).Text = ""
 StatusBarX.Panels(2).Text = ""
 StatusBarX.Panels(3).Text = ""
 
vPaso = False

Me.MousePointer = vbDefault

End Sub



Private Sub sbConsulta()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

strSQL = "select Ex.*, Soc.NOMBRE" _
        & ", Pl.COD_PLAN + ' - ' + rtrim(Pl.DESCRIPCION) as 'Plan'" _
        & ", Pc.COD_CAUSA + ' - ' + rtrim(Pc.DESCRIPCION) as 'Causa'" _
        & ", Te.COD_ENFERMEDAD + ' - ' + rtrim(Te.DESCRIPCION) as 'Enfermedad'" _
        & ", Co.COD_COMITE + ' - ' + rtrim(Co.DESCRIPCION) as 'Comite'" _
        & " from FSL_EXPEDIENTES Ex" _
        & " inner join SOCIOS Soc on Ex.Cedula = Soc.Cedula" _
        & " inner join FSL_PLANES Pl on Ex.COD_PLAN = Pl.COD_PLAN" _
        & " inner join FSL_PLANES_CAUSAS Pc on Ex.COD_PLAN = Pc.COD_PLAN and Ex.COD_CAUSA = Pc.COD_CAUSA" _
        & " inner join FSL_TIPOS_ENFERMEDADES Te on Ex.COD_ENFERMEDAD = Te.COD_ENFERMEDAD" _
        & " inner join FSL_COMITES Co on Ex.COD_COMITE = Co.COD_COMITE" _
        & " Where Ex.COD_EXPEDIENTE = " & txtExpediente.Text

rs.Open strSQL, glogon.Conection, adOpenStatic

Call sbLimpiaPantalla

If Not rs.EOF And Not rs.BOF Then

  Call sbToolBar(tlb, "activo")
  vEdita = True

  txtExpediente.Text = rs!COD_EXPEDIENTE
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre

  
  txtPresentaCedula.Text = rs!Presenta_Cedula & ""
  txtPresentaNombre.Text = rs!PRESENTA_NOMBRE & ""
  txtPresentaNotas.Text = rs!Presenta_Notas & ""
  
  txtRefNumero.Text = rs!Referencia_Numero
  dtpRefFecha.Value = rs!Fecha_Establece_Causa
  
  dtpEnfermedad.Value = rs!Enfermedad_Fecha
  
  
  txtNotas.Text = rs!notas & ""
  txtEnfermedadNotas.Text = rs!Enfermedad_Notas & ""
       
       
  vPaso = True
        Call sbCboAsignaDato(cboTipo, rs!Plan, True)
        Call sbCboAsignaDato(cboCausa, rs!Causa, True)
        Call sbCboAsignaDato(cboComite, rs!Comite, True)
        Call sbCboAsignaDato(cboEnfermedad, rs!Enfermedad, True)
        Call sbCboAsignaDato(cboRefTipoDoc, rs!Referencia_Documento, True)
  vPaso = False
  
  ssTab.TabEnabled(1) = True
  ssTab.TabEnabled(2) = True
  ssTab.TabEnabled(3) = True
  ssTab.TabEnabled(4) = True
  ssTab.TabEnabled(5) = True

  tlbAux.Buttons.Item(1).Enabled = False 'Gestiones
  tlbAux.Buttons.Item(2).Enabled = False 'Apelacion
  tlbAux.Buttons.Item(4).Enabled = False 'Aplicar
  
  Select Case rs!Estado
   Case "P" 'Pendiente
        txtEstado.Text = "PENDIENTE"
        txtEstado.Tag = "P"
        Set imgEstado.Picture = ImageList1.ListImages.Item(8).Picture
        
        tlbAux.Buttons.Item(1).Enabled = True 'Gestiones
     
    Case "A" 'Aprobado
        txtEstado.Text = "APROBADO"
        txtEstado.Tag = "A"
        Set imgEstado.Picture = ImageList1.ListImages.Item(7).Picture
        
        tlbAux.Buttons.Item(1).Enabled = True 'Gestiones
        tlbAux.Buttons.Item(2).Enabled = False 'Apelacion
        tlbAux.Buttons.Item(4).Enabled = True 'Aplicar
   
    Case "R" 'Rechazado
        txtEstado.Text = "RECHAZADO"
        txtEstado.Tag = "R"
        Set imgEstado.Picture = ImageList1.ListImages.Item(9).Picture
        
        tlbAux.Buttons.Item(1).Enabled = True 'Gestiones
        tlbAux.Buttons.Item(2).Enabled = True 'Apelacion
        tlbAux.Buttons.Item(4).Enabled = False 'Aplicar
   
    Case "X" 'Aplicado
        txtEstado.Text = "APLICADO"
        txtEstado.Tag = "X"
        Set imgEstado.Picture = ImageList1.ListImages.Item(10).Picture
        
        tlbAux.Buttons.Item(1).Enabled = False 'Gestiones
        tlbAux.Buttons.Item(2).Enabled = False 'Apelacion
        tlbAux.Buttons.Item(4).Enabled = False 'Aplicar
 
  End Select


 'Totales (Operaciones)
 If rs!TIPO_DESEMBOLSO = "F" Then
    lblTD_Label_01.Caption = "Plan"
    lblTD_Label_02.Caption = "Contrato"
    txtTD_Destino.Text = "-- FONDOS --"
    txtTD_Texto_01.Text = rs!FND_PLAN & ""
    txtTD_Texto_02.Text = rs!FND_CONTRATO & ""
 Else
    lblTD_Label_01.Caption = "No. Solicitud"
    lblTD_Label_02.Caption = "No. Remesa"
    txtTD_Destino.Text = "-- TESORERIA --"
    txtTD_Texto_01.Text = rs!Tesoreria_Solicitud & ""
    txtTD_Texto_02.Text = rs!TESORERIA_REMESA & ""
 End If
  
 txtTotalAplicado.Text = Format(rs!Total_Aplicado, "Standard")
 txtTotalDisponible.Text = Format(rs!Total_Disponible, "Standard")
 txtLiquidacionMonto.Text = Format(rs!TOTAL_SOBRANTE, "Standard")
  
 txtTotalAplicado.ToolTipText = rs!Apl_Tipo_Doc & " .: " & rs!Apl_Num_Doc
  
 'Fin de Totales
  txtResolucionNotas.Text = rs!Resolucion_Notas & ""

  Select Case rs!Resolucion_Estado
    Case "A" 'Aprobado
        cboResolucion.Text = "Aprobado"
    Case "R" 'Rechazado
        cboResolucion.Text = "Rechazado"
  End Select
    
    StatusBarX.Panels(1).Text = "Resolución.:" & rs!Resolucion_Fecha & " ¦ " & rs!Resolucion_Usuario
    StatusBarX.Panels(2).Text = "Rf.:" & rs!Registro_Fecha & ""
    StatusBarX.Panels(3).Text = "Ru.: " & rs!Registro_Usuario & ""

Else
    MsgBox "No existe el expediente, verifique!", vbCritical
End If
rs.Close

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

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
           & " where cedula = '" & .Tag & "' and cod_comite = '" & SIFGlobal.fxSIFCodText(cboComite.Text) & "'"
     rs.Open strSQL, glogon.Conection, adOpenStatic
         txtMiembroUsuario.Text = rs!Usuario_Vinculado
     rs.Close
     
     txtMiembroClave.SetFocus

  End If
  
End With

End Sub

Private Sub lswRequisitos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

On Error GoTo vError

If txtEstado.Tag <> "P" Then
    MsgBox "El expediente no está Pendiente! No se pueden modificar los requisitos!", vbExclamation
    Exit Sub
End If

strSQL = "update FSL_EXPEDIENTES_REQUISITOS set Estado = " & IIf(Item.Checked, 1, 0) _
       & ", registro_fecha = getdate(), registro_usuario = '" & glogon.Usuario & "'" _
       & " where cod_expediente = " & txtExpediente.Text & " and Cod_Requisito = '" & Item.Tag & "'"

glogon.Conection.Execute strSQL
Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub



Private Sub sbReporte()

 With frmContenedor.Crt
  .Reset
  .WindowShowGroupTree = True
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "Reportes FOSOL"
 
  .Connect = glogon.ConectRPT
   
  .ReportFileName = SIFGlobal.fxSIFPathReportes("FSL_ExpedienteBoleta.rpt")
  
  .Formulas(0) = "fxCodigoBarras = '*" & txtExpediente.Text & "*'"
  .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
  .Formulas(2) = "fxUsuario='USUARIO: " & glogon.Usuario & "'"
  .Formulas(3) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
                   
  .SelectionFormula = "{vFSL_CasosLista.COD_EXPEDIENTE} =" & txtExpediente.Text

  

  .PrintReport
  
  
End With

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCedula.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtPresentaCedula.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
'      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If txtExpediente.Text = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "Ex.Cedula"
       gBusquedas.Orden = "Ex.Cedula"
       gBusquedas.Consulta = "select Ex.cod_expediente,Soc.nombre from FSL_EXPEDIENTES Ex inner join SOCIOS Soc on EX.cedula = Soc.Cedula"
       frmBusquedas.Show vbModal
       txtExpediente.SetFocus
       txtExpediente = gBusquedas.Resultado

    Case "REPORTES"
       Call sbReporte
       
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub tlbResolucion_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim y As Integer, vNumResolutores As Integer

On Error GoTo vError

If txtEstado.Tag <> "P" Then
   MsgBox "Este Expediente No se encuentra pendiente, no puede registrarse una resolución!", vbExclamation
   Exit Sub
End If

If Mid(cboResolucion.Text, 1, 1) = "A" Then
    If imgRequisitos.Tag = 0 Then
       MsgBox "Este Expediente no cumple con los requisitos para su Aprobación!", vbExclamation
       Exit Sub
    End If
    
    If imgTiempoPresentacion.Tag = 0 Then
       MsgBox "Este Expediente no cumple el Tiempo de Presentación para su Aprobación!", vbExclamation
       Exit Sub
    End If
    
    If imgExpedientesActivos.Tag = 0 Then
       MsgBox "Este Expediente No puede ser Aprobado por que ya fueron presentadas otras solicitudes de la misma persona (Revise DUPLICADOS!)", vbExclamation
       Exit Sub
    End If
End If


If Len(txtResolucionNotas.Text) < 10 Then
   MsgBox "Indique una nota válida para la resolución!", vbExclamation
   Exit Sub
End If

strSQL = "select NUMERO_RESOLUTORES from FSL_Comites" _
        & " where cod_Comite = '" & SIFGlobal.fxSIFCodText(cboComite.Text) & "'"
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



Me.MousePointer = vbHourglass


strSQL = "update FSL_EXPEDIENTES set RESOLUCION_NOTAS = '" & txtResolucionNotas.Text _
       & "',RESOLUCION_ESTADO = '" & Mid(cboResolucion.Text, 1, 1) & "', RESOLUCION_FECHA = getdate()" _
       & " ,RESOLUCION_USUARIO = '" & glogon.Usuario & "',ESTADO = '" & Mid(cboResolucion.Text, 1, 1) _
       & "' where COD_EXPEDIENTE = " & txtExpediente.Text
glogon.Conection.Execute strSQL


strSQL = "delete FSL_EXPEDIENTE_COMITE WHERE COD_EXPEDIENTE = " & txtExpediente.Text
glogon.Conection.Execute strSQL


With lswComite.ListItems
   For i = 1 To .Count
      If .Item(i).Checked Then
            strSQL = "INSERT FSL_EXPEDIENTE_COMITE(COD_EXPEDIENTE,COD_COMITE,CEDULA,ASIGNA_FECHA,ASIGNA_USUARIO,RESOLUCION_ESTADO)" _
                   & " values(" & txtExpediente & ",'" & SIFGlobal.fxSIFCodText(cboComite.Text) & "','" & _
                   .Item(i).Tag & "',getdate(),'" & glogon.Usuario & "','" & Mid(cboResolucion.Text, 1, 1) & "')"
            glogon.Conection.Execute strSQL
      End If
   Next i
End With

Me.MousePointer = vbDefault

MsgBox "Expediente actualizado satisfactoriamente...", vbInformation

Call sbConsulta

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

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select CEDULA,NOMBRE from SOCIOS"
    gBusquedas.Orden = "CEDULA"
    gBusquedas.Columna = "CEDULA"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCedula_LostFocus()
txtNombre.Text = fxNombre(txtCedula)

End Sub

Private Sub txtExpediente_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then Call sbConsulta

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select COD_EXPEDIENTE,CEDULA from FSL_EXPEDIENTES"
    gBusquedas.Orden = "COD_EXPEDIENTE"
    gBusquedas.Columna = "COD_EXPEDIENTE"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtExpediente.Text = gBusquedas.Resultado
    txtCedula.Text = gBusquedas.Resultado2
    If Trim(txtExpediente.Text) <> "" Then Call sbConsulta
End If

End Sub

Private Sub txtExpediente_LostFocus()
If IsNumeric(txtExpediente.Text) Then
  Call sbConsulta
End If
End Sub


Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If Not IsNumeric(txtExpediente.Text) Then
    ssTab.Tab = 0
End If

Select Case ssTab.Tab
  Case 0 'General
  Case 1 'Requisitos
       vPaso = True
       strSQL = "Select Ex.COD_REQUISITO,Rq.DESCRIPCION, EX.Estado, Ex.Opcional " _
              & " from FSL_EXPEDIENTES_REQUISITOS Ex " _
              & "  inner join FSL_REQUISITOS Rq on Ex.cod_requisito = Rq.cod_requisito" _
              & " where Ex.cod_Expediente = " & txtExpediente.Text
        rs.Open strSQL, glogon.Conection, adOpenStatic
        With lswRequisitos.ListItems
           .Clear
           Do While Not rs.EOF
            Set itmX = .Add(, , rs!Descripcion)
                itmX.Tag = rs!Cod_Requisito
                If rs!Opcional = 1 Then
                    itmX.SubItems(1) = "Sí"
                Else
                    itmX.SubItems(1) = "No"
                End If
                itmX.Checked = rs!Estado
            rs.MoveNext
           Loop
        End With
        rs.Close
        vPaso = False
        
  Case 2 'Operaciones
        strSQL = "select E.ID_SOLICITUD, E.REFERENCIA, Gar.DESCRIPCION, R.PRIDEDUC, R.MONTOAPR " _
               & "   , E.SALDO_CORTE , E.MONTO_BASE, E.PORC_RELACION , E.TIPO_TABLA , E.PORCENTAJE" _
               & "   , E.MONTO_RECONOCIMIENTO, E.TIEMPO_TRANS, case when E.Tipo_Base = 'S' then 'Saldo' else 'Mnt.Form.' end as 'BASE'" _
               & " from FSL_EXPEDIENTES_DETALLE  E inner join REG_CREDITOS R on E.ID_SOLICITUD = R.ID_SOLICITUD" _
               & " inner join CRD_GARANTIA_TIPOS Gar on R.GARANTIA = Gar.GARANTIA" _
               & " Where E.COD_EXPEDIENTE = " & txtExpediente.Text & " Order by isnull(E.referencia,E.id_Solicitud) desc"
        Call sbCargaGrid(vgCreditos, 13, strSQL, True)
  
  
  Case 3 'Resolucion
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
  
        'Refresca Validaciones

        strSQL = "select dbo.fxFSL_ExpedienteValidaRequisitos(Ex.Cod_Expediente) as 'CumpleRequisitos'" _
                & ", dbo.fxFSL_ExpedienteValidaTiempoPresentacion(Ex.Cod_Expediente) as 'CumpleTiempo'" _
                & ", dbo.fxFSL_ExpedienteValidaRegistro(Ex.Cedula, Ex.Cod_Plan, Ex.Cod_Causa,Ex.Cod_Expediente) as 'CumpleRegistro'" _
                & " from FSL_EXPEDIENTES Ex" _
                & " Where Ex.COD_EXPEDIENTE = " & txtExpediente.Text
        rs.Open strSQL, glogon.Conection, adOpenStatic
            
            imgRequisitos.Tag = rs!CumpleRequisitos
            imgTiempoPresentacion.Tag = rs!CumpleTiempo
            imgExpedientesActivos.Tag = rs!CumpleRegistro
            
            If rs!CumpleRequisitos = 1 Then
               Set imgRequisitos.Picture = ImageList1.ListImages.Item(7).Picture
            Else
               Set imgRequisitos.Picture = ImageList1.ListImages.Item(9).Picture
            End If
            
            
            If rs!CumpleTiempo = 1 Then
               Set imgTiempoPresentacion.Picture = ImageList1.ListImages.Item(7).Picture
            Else
               Set imgTiempoPresentacion.Picture = ImageList1.ListImages.Item(9).Picture
            End If
            
            If rs!CumpleRegistro = 1 Then
               Set imgExpedientesActivos.Picture = ImageList1.ListImages.Item(7).Picture
            Else
               Set imgExpedientesActivos.Picture = ImageList1.ListImages.Item(9).Picture
            End If
   
        rs.Close
  
  Case 4 'Gestiones
        strSQL = "select Tg.Descripcion, Eg.*" _
               & " from FSL_EXPEDIENTE_GESTIONES Eg inner join FSL_TIPOS_GESTIONES Tg on Eg.COD_GESTION = Tg.COD_GESTION" _
               & " Where Eg.cod_Expediente = " & txtExpediente.Text & " order by registro_fecha desc"
        rs.Open strSQL, glogon.Conection, adOpenStatic
        vgGestiones.MaxRows = 0
        Do While Not rs.EOF
          vgGestiones.MaxRows = vgGestiones.MaxRows + 1
          vgGestiones.Row = vgGestiones.MaxRows
          
          vgGestiones.Col = 1
          vgGestiones.Text = rs!Descripcion
          vgGestiones.TextTip = TextTipFixed
          vgGestiones.TextTipDelay = 1000
        
          vgGestiones.CellNote = "Fecha : " & rs!Registro_Fecha & vbCrLf & "Usuario : " & rs!Registro_Usuario
          vgGestiones.CellTag = CStr(rs!Linea)
            
          vgGestiones.Col = 2
          vgGestiones.Text = rs!notas
              
          vgGestiones.Col = 3
          vgGestiones.Text = rs!Registro_Fecha
              
          vgGestiones.Col = 4
          vgGestiones.Text = rs!Registro_Usuario
         
          vgGestiones.RowHeight(vgGestiones.Row) = vgGestiones.MaxTextRowHeight(vgGestiones.Row)
          
         rs.MoveNext
        Loop
        rs.Close
  
  
  Case 5 'Apelaciones

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

End Select


End Sub

Private Function fxExpedienteConsecutivo(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(max(cod_Expediente),0) as 'Ultimo'" _
       & " from FSL_Expedientes where Cedula = '" & pCedula & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
    fxExpedienteConsecutivo = rs!ultimo
rs.Close

End Function

Private Function fxPlanTipoDesembolso(pPlan As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select TIPO_DESEMBOLSO" _
       & " from FSL_PLANES where cod_plan = '" & pPlan & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
    fxPlanTipoDesembolso = rs!TIPO_DESEMBOLSO
rs.Close

End Function


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoDesembolso As String

On Error GoTo vError


If Mid(txtEstado.Text, 1, 1) <> "P" Then
    MsgBox "No se puede modificar este trámite porque no se encuentra pendiente", vbExclamation
    Exit Sub
End If

vTipoDesembolso = fxPlanTipoDesembolso(SIFGlobal.fxSIFCodText(cboTipo.Text))


If Not vEdita Then
   strSQL = "insert FSL_EXPEDIENTES(COD_EXPEDIENTE,CEDULA, COD_PLAN, COD_CAUSA,COD_COMITE,COD_ENFERMEDAD,ESTADO,RESOLUCION_ESTADO" _
          & ",PRESENTA_CEDULA,PRESENTA_NOMBRE,PRESENTA_NOTAS,REFERENCIA_DOCUMENTO,REFERENCIA_NUMERO" _
          & ",ENFERMEDAD_FECHA,ENFERMEDAD_USUARIO,ENFERMEDAD_NOTAS,FECHA_ESTABLECE_CAUSA,NOTAS" _
          & ",TOTAL_DISPONIBLE,TOTAL_APLICADO,TOTAL_SOBRANTE,REGISTRO_FECHA,REGISTRO_USUARIO" _
          & ",TIPO_DESEMBOLSO)" _
          & " VALUES(dbo.fxFSL_ExpedienteConsecutivo(),'" & txtCedula.Text & "','" & SIFGlobal.fxSIFCodText(cboTipo.Text) _
          & "','" & SIFGlobal.fxSIFCodText(cboCausa.Text) & "','" & SIFGlobal.fxSIFCodText(cboComite.Text) _
          & "','" & SIFGlobal.fxSIFCodText(cboEnfermedad.Text) & "','P','P','" & txtPresentaCedula.Text _
          & "','" & txtPresentaNombre.Text & "','" & txtPresentaNotas.Text & "','" & cboRefTipoDoc.Text _
          & "','" & txtRefNumero.Text & "','" & Format(dtpEnfermedad.Value, "yyyy/mm/dd") _
          & "','" & glogon.Usuario & "','" & txtEnfermedadNotas.Text & "','" & Format(dtpRefFecha.Value, "yyyy/mm/dd") _
          & "','" & txtNotas.Text & "',0,0,0,getdate(),'" & glogon.Usuario & "','" & vTipoDesembolso & "')"
   
   glogon.Conection.Execute strSQL

   txtExpediente.Text = fxExpedienteConsecutivo(txtCedula.Text)


Else

  strSQL = "update FSL_EXPEDIENTES set COD_PLAN = '" & SIFGlobal.fxSIFCodText(cboTipo.Text) _
         & "',COD_CAUSA = '" & SIFGlobal.fxSIFCodText(cboCausa.Text) & "' ,COD_COMITE ='" & SIFGlobal.fxSIFCodText(cboComite.Text) _
         & "',COD_ENFERMEDAD = '" & SIFGlobal.fxSIFCodText(cboEnfermedad.Text) _
         & "',notas = '" & txtNotas.Text & "',PRESENTA_CEDULA = '" & txtPresentaCedula.Text & "', PRESENTA_NOMBRE = '" _
         & txtPresentaNombre.Text & "', REFERENCIA_DOCUMENTO = '" & cboRefTipoDoc.Text & "', REFERENCIA_NUMERO = '" _
         & txtRefNumero.Text & "', PRESENTA_NOTAS = '" & txtPresentaNotas.Text & "', FECHA_ESTABLECE_CAUSA = '" _
         & Format(dtpRefFecha.Value, "yyyy/mm/dd") & "', ENFERMEDAD_FECHA = '" & Format(dtpEnfermedad.Value, "yyyy/mm/dd") _
         & "',ENFERMEDAD_NOTAS = '" & txtEnfermedadNotas.Text & "', MODIFICA_USUARIO = '" & glogon.Usuario _
         & "', MODIFICA_FECHA = GETDATE(), TIPO_DESEMBOLSO = '" & vTipoDesembolso _
         & "' where COD_EXPEDIENTE = " & txtExpediente.Text
  glogon.Conection.Execute strSQL

End If

'Actualiza Requisitos
strSQL = "exec spFSL_ExpedienteRequisitos " & txtExpediente.Text & ",'" & glogon.Usuario & "'"
glogon.Conection.Execute strSQL


'Actualiza Calculos de Creditos (FOSOL)
strSQL = "exec spFSL_ExpedienteOperaciones " & txtExpediente.Text & ",'" & glogon.Usuario & "'"
glogon.Conection.Execute strSQL


Call sbToolBar(tlb, "activo")
Call sbConsulta

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub


Private Sub sbAplicarFosol()
Dim strSQL As String, i As Integer
Dim rs As New ADODB.Recordset, pTipoDoc As String, pNumDoc As String

On Error GoTo vError

i = MsgBox("Esta seguro que aplicar los calculos del FOSOL?", vbYesNo)
If i = vbNo Then Exit Sub


Me.MousePointer = vbHourglass

strSQL = "exec spFSL_AplicacionFosol " & txtExpediente.Text & ",'" & glogon.Usuario & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
  pTipoDoc = rs!Tipo_Documento
  pNumDoc = rs!Numero_Documento
rs.Close

Me.MousePointer = vbDefault
MsgBox "Aplicación realizada satisfactoriamente..!", vbInformation

Call sbImprimeRecibo(pNumDoc, pTipoDoc)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub


Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError

GLOBALES.gTag = txtExpediente.Text

Select Case Button.Key
 Case "Gestiones"
    Call sbSIFForms("frmFSL_ExpedienteGestiones", 1, , , False, Me)
 
 Case "Apelacion"
    Call sbSIFForms("frmFSL_ExpedienteApelaciones", 1, , , False, Me)

 Case "Aplicar"
    Call sbAplicarFosol
End Select

Call sbConsulta

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Public Sub sbConsultaExterna(xTramTemp As String)
 txtExpediente.Text = xTramTemp
 If Trim(txtExpediente) <> "" Then
    Call txtExpediente_KeyDown(vbKeyReturn, 0)
 End If

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

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPresentaCedula.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select CEDULA,NOMBRE from SOCIOS"
    gBusquedas.Orden = "NOMBRE"
    gBusquedas.Columna = "NOMBRE"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtPresentaCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPresentaNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select CEDULA,NOMBRE from SOCIOS"
    gBusquedas.Orden = "CEDULA"
    gBusquedas.Columna = "CEDULA"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtPresentaCedula.Text = gBusquedas.Resultado
    txtPresentaNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtPresentaCedula_LostFocus()
Dim pNombre As String

pNombre = fxNombre(txtPresentaCedula)

If Trim(pNombre) <> "" Then
   txtPresentaNombre.Text = pNombre
End If

End Sub

Private Sub txtPresentaNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPresentaNotas.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select CEDULA,NOMBRE from SOCIOS"
    gBusquedas.Orden = "NOMBRE"
    gBusquedas.Columna = "NOMBRE"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtPresentaCedula.Text = gBusquedas.Resultado
    txtPresentaNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtPresentaNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then cboRefTipoDoc.SetFocus

End Sub

Private Sub txtRefNumero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpRefFecha.SetFocus

End Sub

