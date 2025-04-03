VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmRadar_Colocadores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Radar: Colocadores"
   ClientHeight    =   6216
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6216
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      MaxLength       =   38
      TabIndex        =   1
      Top             =   480
      Width           =   6252
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "e"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   5172
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   9732
      _ExtentX        =   17166
      _ExtentY        =   9123
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmRadar_Colocadores.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(13)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(15)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(17)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboTipo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtIdentificacion"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkActivo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtEMail"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtAptoPostal"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtTelefono2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCelular"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtTelefono"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtEMail2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtWebSite"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtTelFax"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "GroupBox1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "GroupBox2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Usuarios"
      TabPicture(1)   =   "frmRadar_Colocadores.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lsw"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Asociar"
      TabPicture(2)   =   "frmRadar_Colocadores.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GroupBox3"
      Tab(2).Control(1)=   "txtLineaCod"
      Tab(2).Control(2)=   "txtLineaDesc"
      Tab(2).Control(3)=   "lswDestinos"
      Tab(2).Control(4)=   "tlbDestino"
      Tab(2).Control(5)=   "fsbCredito"
      Tab(2).ControlCount=   6
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   2292
         Left            =   -74760
         TabIndex        =   45
         Top             =   480
         Width           =   1812
         _Version        =   1245185
         _ExtentX        =   3196
         _ExtentY        =   4043
         _StockProps     =   79
         Caption         =   "Asociar con:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton optLinea 
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   46
            Top             =   480
            Width           =   2172
            _Version        =   1245185
            _ExtentX        =   3831
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Convenio"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optLinea 
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   47
            Top             =   840
            Width           =   2772
            _Version        =   1245185
            _ExtentX        =   4890
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Destino"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   2
         End
      End
      Begin VB.TextBox txtLineaCod 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72840
         TabIndex        =   41
         Top             =   600
         Width           =   852
      End
      Begin VB.TextBox txtLineaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   600
         Width           =   5172
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1092
         Left            =   240
         TabIndex        =   35
         Top             =   2400
         Width           =   9252
         _Version        =   1245185
         _ExtentX        =   16319
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Redes Sociales"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin VB.TextBox txtRs_02 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   37
            Top             =   720
            Width           =   6132
         End
         Begin VB.TextBox txtRs_01 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            TabIndex        =   36
            Top             =   360
            Width           =   6132
         End
         Begin VB.Label Label1 
            Caption         =   "Red Social Secundaria"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   18
            Left            =   1200
            TabIndex        =   39
            Top             =   720
            Width           =   2532
         End
         Begin VB.Label Label1 
            Caption         =   "Red Social Principal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   16
            Left            =   1200
            TabIndex        =   38
            Top             =   360
            Width           =   2532
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1452
         Left            =   240
         TabIndex        =   27
         Top             =   3480
         Width           =   9252
         _Version        =   1245185
         _ExtentX        =   16319
         _ExtentY        =   2561
         _StockProps     =   79
         Caption         =   "Dirección"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin VB.TextBox txtDireccion 
            Appearance      =   0  'Flat
            DataField       =   "e"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   3240
            MultiLine       =   -1  'True
            TabIndex        =   28
            ToolTipText     =   "Dirección Exacta"
            Top             =   420
            Width           =   6015
         End
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   312
            Left            =   1200
            TabIndex        =   29
            Top             =   420
            Width           =   1932
            _Version        =   1245185
            _ExtentX        =   3408
            _ExtentY        =   550
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
         End
         Begin XtremeSuiteControls.ComboBox cboCanton 
            Height          =   312
            Left            =   1200
            TabIndex        =   30
            Top             =   780
            Width           =   1932
            _Version        =   1245185
            _ExtentX        =   3408
            _ExtentY        =   550
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
         End
         Begin XtremeSuiteControls.ComboBox cboDistrito 
            Height          =   312
            Left            =   1200
            TabIndex        =   31
            Top             =   1140
            Width           =   1932
            _Version        =   1245185
            _ExtentX        =   3408
            _ExtentY        =   550
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
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   852
         End
         Begin VB.Label Label1 
            Caption         =   "Cantón"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   852
         End
         Begin VB.Label Label1 
            Caption         =   "Distrito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   6
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   852
         End
      End
      Begin VB.TextBox txtTelFax 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   1620
         Width           =   2055
      End
      Begin VB.TextBox txtWebSite 
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   11
         Top             =   900
         Width           =   4935
      End
      Begin VB.TextBox txtEMail2 
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   10
         Top             =   1620
         Width           =   4935
      End
      Begin VB.TextBox txtTelefono 
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   900
         Width           =   2055
      End
      Begin VB.TextBox txtCelular 
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   1980
         Width           =   2055
      End
      Begin VB.TextBox txtTelefono2 
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1260
         Width           =   2055
      End
      Begin VB.TextBox txtAptoPostal 
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   6
         Top             =   1980
         Width           =   4935
      End
      Begin VB.TextBox txtEMail 
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   5
         Top             =   1260
         Width           =   4935
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Activo ?"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   8280
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtIdentificacion 
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   4332
         Left            =   -74760
         TabIndex        =   13
         Top             =   540
         Width           =   9132
         _ExtentX        =   16108
         _ExtentY        =   7641
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lswDestinos 
         Height          =   3852
         Left            =   -72840
         TabIndex        =   42
         Top             =   1080
         Width           =   7452
         _ExtentX        =   13145
         _ExtentY        =   6795
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Asociar ?"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   3775
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbDestino 
         Height          =   288
         Left            =   -66120
         TabIndex        =   43
         Top             =   600
         Width           =   732
         _ExtentX        =   1291
         _ExtentY        =   508
         ButtonWidth     =   487
         ButtonHeight    =   466
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Add"
               Object.ToolTipText     =   "Agregar o Modificar"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cancel"
               Object.ToolTipText     =   "Eliminar"
               ImageIndex      =   9
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComCtl2.FlatScrollBar fsbCredito 
         Height          =   252
         Left            =   -66720
         TabIndex        =   44
         Top             =   600
         Width           =   492
         _ExtentX        =   868
         _ExtentY        =   445
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   4560
         TabIndex        =   48
         Top             =   480
         Width           =   3132
         _Version        =   1245185
         _ExtentX        =   5525
         _ExtentY        =   550
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
      End
      Begin VB.Label Label1 
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   360
         TabIndex        =   23
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Web Site"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3600
         TabIndex        =   22
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail (2)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   3600
         TabIndex        =   21
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono (1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   20
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Celular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   19
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono (2)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Apto. Postal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   17
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail (1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3600
         TabIndex        =   16
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   3600
         TabIndex        =   15
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   9120
      TabIndex        =   24
      Top             =   480
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   264
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   466
      ButtonWidth     =   487
      ButtonHeight    =   466
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
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   240
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":0054
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":0150
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":0255
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":0365
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":0477
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":059B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":07B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":0FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":1630
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRadar_Colocadores.frx":204E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Colocador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   14
      Left            =   360
      TabIndex        =   26
      Top             =   480
      Width           =   1572
   End
End
Attribute VB_Name = "frmRadar_Colocadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vTipoLinea As String
Dim vCantonMascara As String, vDistritoMascara As String, vFechaActual As Date
Dim vScroll As Boolean, vPaso As Boolean


Private Sub cboCanton_Click()
Dim strSQL As String

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
            & "' and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistrito, strSQL, False, True)

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

Private Sub cboProvincia_Click()
Dim strSQL As String

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 COLOCADOR_ID from RADAR_COLOCADORES"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COLOCADOR_ID > '" & txtCodigo.Text & "' order by COLOCADOR_ID asc"
    Else
       strSQL = strSQL & " where COLOCADOR_ID < '" & txtCodigo.Text & "' order by COLOCADOR_ID desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!COLOCADOR_ID
      Call txtCodigo_LostFocus
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 37
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

 vModulo = 37

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
ssTab.Tab = 0

vEdita = False
vFechaActual = Format(fxFechaServidor, "dd/mm/yyyy")

'Carga Clasificaciones de Clientes
strSQL = "select rtrim(COLOCADOR_TIPO) as 'IdX', rtrim(Descripcion) as 'ItmX' from RADAR_COLOCADORES_TIPOS where activo = 1"
Call sbCbo_Llena_New(cboTipo, strSQL, False, True)

vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False
 
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


txtIdentificacion.Text = ""
txtNombre = ""
txtTelefono.Text = ""
txtTelefono2.Text = ""
txtTelFax.Text = ""
txtCelular.Text = ""
txtWebSite.Text = ""
txtEMail.Text = ""
txtEmail2.Text = ""
txtAptoPostal.Text = ""

txtRs_01.Text = ""
txtRs_02.Text = ""

txtDireccion = ""

chkActivo.Value = vbChecked


ssTab.Tab = 0
ssTab.TabEnabled(1) = False
ssTab.TabEnabled(2) = False

End Sub


Private Sub sbCargaListaAsociar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

lswDestinos.ListItems.Clear


If vTipoLinea = "C" Then
   strSQL = "Select CVL.CODIGO, C.DESCRIPCION, CVL.REGISTRO_FECHA, CVL.REGISTRO_USUARIO" _
          & " from RADAR_COLOCADORES_LINKS CVL inner join CRD_CONVENIOS C on CVL.CODIGO = C.COD_CONVENIO" _
          & " Where CVL.TIPO = 'C' AND COLOCADOR_ID ='" & vCodigo & "'"

Else
   strSQL = "Select CVD.CODIGO,C.DESCRIPCION, CVD.REGISTRO_FECHA, CVD.REGISTRO_USUARIO" _
          & " from RADAR_COLOCADORES_LINKS CVD inner join Catalogo_destinos C on CVD.CODIGO = C.COD_DESTINO" _
          & " Where CVD.TIPO = 'D' AND CVD.COLOCADOR_ID ='" & vCodigo & "'"
End If


Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   Set itmX = lswDestinos.ListItems.Add(, , rs!codigo)
    itmX.SubItems(1) = Trim(rs!Descripcion)
    itmX.SubItems(2) = "...."
    itmX.SubItems(3) = Format(rs!registro_Fecha, "dd/mm/yyyy")
    itmX.SubItems(4) = rs!registro_usuario
    
   rs.MoveNext
Loop
     
rs.Close

End Sub



Private Sub fsbCredito_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
      
      If optLinea(0) Then 'se mueve en tabla convenios lineas
        strSQL = "select Top 1 COD_CONVENIO AS 'CODIGO', Descripcion from CRD_CONVENIOS"
        If Len(txtLineaCod.Text) > 0 Then
            If fsbCredito.Value = 1 Then
               strSQL = strSQL & " where COD_CONVENIO > '" & txtLineaCod.Text & "' and ACTIVO = 1  order by COD_CONVENIO asc"
            Else
               strSQL = strSQL & " where COD_CONVENIO < '" & txtLineaCod.Text & "' and ACTIVO = 1  order by COD_CONVENIO desc"
            End If
            
        End If
                
        Call OpenRecordSet(rs, strSQL)
      
      Else 'Se mueve en la tabla de convenios_destinos
      
        strSQL = "select Top 1 COD_DESTINO as 'Codigo',DESCRIPCION from Catalogo_Destinos"
        If Len(txtLineaCod.Text) > 0 Then
            If fsbCredito.Value = 1 Then
               strSQL = strSQL & " where COD_DESTINO > '" & txtLineaCod.Text & "' order by COD_DESTINO asc"
            Else
               strSQL = strSQL & " where COD_DESTINO < '" & txtLineaCod.Text & "' order by COD_DESTINO desc"
            End If
        End If
        Call OpenRecordSet(rs, strSQL)
        
      End If 'Fin de optLinea(0)
      

    If Not rs.EOF And Not rs.BOF Then
      txtLineaCod.Text = rs!codigo
      txtLineaDesc.Text = rs!Descripcion
    End If
    rs.Close
End If

vScroll = False
fsbCredito.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswDestinos_Click()

If lswDestinos.ListItems.Count <= 0 Then Exit Sub

txtLineaCod.Text = Trim(lswDestinos.SelectedItem.Text)
txtLineaDesc.Text = Trim(lswDestinos.SelectedItem.SubItems(1))

End Sub

Private Sub optLinea_Click(Index As Integer)
  Select Case Index
    Case 0
       vTipoLinea = "C"
    Case 1
       vTipoLinea = "D"
  End Select
  
  txtLineaCod.Text = ""
  txtLineaDesc.Text = ""
  
  Call sbCargaListaAsociar
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, curMonto As Currency, curSaldo As Currency

If vCodigo = "" Then
  ssTab.Tab = 0
  Exit Sub
End If

Me.MousePointer = vbHourglass


vPaso = True
Select Case ssTab.Tab
   Case 1 'Operaciones
      
'       vPaso = True
'       lswOperaciones.ListItems.Clear
'       curMonto = 0
'       curSaldo = 0
'
'       strSQL = "exec spCxCPersonasCuentas '" & txtCodigo.Text & "','A'"
'       Call OpenRecordSet(rs, strSQL)
'       Do While Not rs.EOF
'         Set itmX = lswOperaciones.ListItems.Add(, , rs!Operacion)
'             itmX.SubItems(1) = rs!Num_Documento
'             itmX.SubItems(2) = Format(rs!Activa_Fecha, "dd/mm/yyyy")
'             itmX.SubItems(3) = Format(rs!Fecha_Vencimiento, "dd/mm/yyyy")
'             itmX.SubItems(4) = Format(rs!Fecha_Pago, "dd/mm/yyyy")
'             itmX.SubItems(5) = Format(rs!Monto, "Standard")
'             itmX.SubItems(6) = Format(rs!Saldo, "Standard")
'             itmX.SubItems(7) = rs!Estado
'             itmX.SubItems(8) = rs!ConceptoDesc
'             itmX.SubItems(9) = rs!OficinaDesc
'             itmX.SubItems(10) = rs!Nombre_Pagador
'
'             curMonto = curMonto + rs!Monto
'             curSaldo = curSaldo + rs!Saldo
'
'          rs.MoveNext
'       Loop
'       rs.Close
'         Set itmX = lswOperaciones.ListItems.Add(, , "---")
'             itmX.SubItems(5) = "-----------"
'             itmX.SubItems(6) = "-----------"
'         Set itmX = lswOperaciones.ListItems.Add(, , lswOperaciones.ListItems.Count - 1)
'             itmX.SubItems(5) = Format(curMonto, "Standard")
'             itmX.SubItems(6) = Format(curSaldo, "Standard")
'
       
       
       vPaso = False
   
End Select

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select COLOCADOR_ID,nombre from RADAR_COLOCADORES"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String


On Error GoTo vError

If Not fxSIFValidaCadena(pCodigo) Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "select P.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & ",rtrim(B.Descripcion) as 'Tipo'" _
       & " from RADAR_COLOCADORES P inner Join RADAR_COLOCADORES_TIPOS B on P.colocador_tipo = B.colocador_tipo" _
       & " left join Provincias Prov on P.Cod_Provincia = Prov.Provincia" _
       & " left join Cantones Cant on P.Cod_Provincia = Cant.Provincia and P.Cod_Canton = Cant.Canton" _
       & " left join Distritos Dist on P.Cod_Provincia = Dist.Provincia and P.Cod_Canton = Dist.Canton and P.Cod_distrito = Dist.distrito" _
       & " where P.COLOCADOR_ID = '" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCodigo = rs!COLOCADOR_ID
  txtCodigo.Text = rs!COLOCADOR_ID
  
  txtIdentificacion.Text = rs!Identificacion & ""
  txtNombre = rs!Nombre & ""
  
  txtTelefono.Text = rs!Telefono_01 & ""
  txtTelefono2.Text = rs!Telefono_02 & ""
  txtTelFax.Text = rs!TEL_FAX & ""
  txtCelular.Text = rs!Tel_Movil & ""

  txtWebSite.Text = rs!Sitio_Web & ""
  txtEMail.Text = rs!Email_01 & ""
  txtEmail2.Text = rs!Email_02 & ""
  txtAptoPostal.Text = rs!APTO_POSTAL & ""

  txtRs_01.Text = rs!SM_01 & ""
  txtRs_02.Text = rs!SM_02 & ""
  

  Call sbCboAsignaDato(cboProvincia, rs!ProvDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
  Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
  Call sbCboAsignaDato(cboDistrito, rs!DistDesc & "")
  
  Call sbCboAsignaDato(cboTipo, rs!Tipo, True, rs!Colocador_Tipo)
     
  cboDistrito.ToolTipText = Trim(rs!cod_distrito) & ""
  txtDireccion.Text = rs!Direccion


  ssTab.Tab = 0
  ssTab.TabEnabled(1) = True
  ssTab.TabEnabled(2) = True

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre no es válido ..."

strSQL = "select count(*) as 'Existe' from RADAR_COLOCADORES" _
        & " where identificacion = '" & txtIdentificacion.Text & "' and COLOCADOR_ID <> '" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
    vMensaje = vMensaje & vbCrLf & " - El número de identificacion ya esta siendo utilizado por otro Colocador (Revise!) ..."
End If
rs.Close
 

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipo As String

On Error GoTo vError


vTipo = cboTipo.ItemData(cboTipo.ListIndex)

'Accion:
strSQL = "select count(*) as 'Existe' from RADAR_COLOCADORES where colocador_id = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 1 Then
  vEdita = True
Else
  vEdita = False
End If
rs.Close
   

If vEdita Then
  strSQL = "update RADAR_COLOCADORES set nombre = '" & Trim(txtNombre.Text) & "',Telefono_01 = '" & txtTelefono.Text & "',Telefono_02 = '" & txtTelefono2.Text _
         & "',Tel_Fax = '" & txtTelFax.Text & "',Tel_Movil ='" & txtCelular.Text & "',Sitio_Web = '" & txtWebSite.Text & "',apto_postal = '" & txtAptoPostal _
         & "',email_01 = '" & txtEMail & "', email_02 = '" & txtEmail2.Text & "',direccion = '" & txtDireccion _
         & "',cod_distrito = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "',cod_canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
         & "',cod_provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
         & "',Identificacion = '" & txtIdentificacion.Text & "',Activo = " & chkActivo.Value & ", colocador_Tipo = '" & vTipo _
         & "',SM_01 = '" & txtRs_01.Text & "',SM_02 = '" & txtRs_02.Text _
         & "' where COLOCADOR_ID = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Colocadores:" & vCodigo)

Else

   vCodigo = txtCodigo.Text
   
   strSQL = "insert into RADAR_COLOCADORES(COLOCADOR_ID,identificacion, nombre,Telefono_01,Telefono_02,Tel_Movil, Tel_fax,Activo,registro_fecha,registro_usuario" _
          & ",apto_postal,email_01,email_02,Sitio_Web,direccion,cod_distrito,cod_provincia,cod_canton,colocador_tipo, SM_01, SM_02)" _
          & " values('" & vCodigo & "','" & txtIdentificacion.Text & "','" & txtNombre.Text & "','" & txtTelefono.Text & "','" & txtTelefono2.Text _
          & "','" & txtCelular.Text & "','" & txtTelFax.Text _
          & "'," & chkActivo.Value & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtAptoPostal.Text & "','" & txtEMail.Text & "','" & txtEmail2.Text & "','" & txtWebSite.Text _
          & "','" & txtDireccion.Text & "','" & cboDistrito.ItemData(cboDistrito.ListIndex) & "','" _
          & cboProvincia.ItemData(cboProvincia.ListIndex) & "','" & cboCanton.ItemData(cboCanton.ListIndex) _
          & "','" & vTipo & "','" & txtRs_01.Text & "','" & txtRs_02.Text & "')"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Colocadores:" & vCodigo)

End If

ssTab.TabEnabled(1) = True

Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete RADAR_COLOCADORES where COLOCADOR_ID = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Colocadores:" & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValidaCodigos(vCodigo As String, vTabla As String, vCampo As String, vFiltro As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as existe from " & vTabla & " where " & vCampo & " =  '" & vCodigo _
    & "' " & vFiltro
Call OpenRecordSet(rs, strSQL)

If rs!Existe > 0 Then
  fxValidaCodigos = True
Else
  fxValidaCodigos = False
End If

rs.Close

End Function

Private Sub tlbDestino_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo.Text = "" Or txtLineaCod.Text = "" Then Exit Sub

Select Case Button.Key
  
  Case "Add"
     
      
        If Not fxValidaCodigos(txtLineaCod.Text, "RADAR_COLOCADORES_LINKS", "CODIGO", " and Tipo = '" & vTipoLinea & "'") Then
            strSQL = "insert RADAR_COLOCADORES_LINKS(codigo,colocador_id,Tipo,registro_fecha,registro_usuario)" _
                   & " values('" & txtLineaCod.Text & "','" & txtCodigo.Text & "','" & vTipoLinea & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
            Call ConectionExecute(strSQL)
        Else
            MsgBox "Este Código: " & txtLineaCod.Text & " ya se encuentra asigando a otro Colocador", vbInformation
        End If
     
  
  Case "Cancel"
     
        strSQL = "Delete RADAR_COLOCADORES_LINKS where codigo = '" & txtLineaCod.Text & "' and COLOCADOR_ID = '" _
               & txtCodigo.Text & "' and Tipo = '" & vTipoLinea & "'"
        Call ConectionExecute(strSQL)
     
End Select

Call sbCargaListaAsociar

txtLineaCod.Text = ""
txtLineaDesc.Text = ""

Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COLOCADOR_ID"
  gBusquedas.Orden = "COLOCADOR_ID"
  gBusquedas.Consulta = "select COLOCADOR_ID,NOMBRE from RADAR_COLOCADORES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()
  Call sbConsulta(txtCodigo.Text)
'  txtNombre.SetFocus
End Sub



Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContacto.SetFocus
End Sub

Private Sub txtEMail2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAptoPostal.SetFocus
End Sub

Private Function fxLineaDescrip(vCodLinea As String, vTipo As String)
Dim strSQL As String, rs As New ADODB.Recordset


If vTipo = "C" Then
        strSQL = "select C.cod_convenio as 'CODIGO',C.descripcion" _
               & " from CRD_Convenios C" _
               & "  left join RADAR_COLOCADORES_LINKS X On C.cod_convenio = X.codigo and X.Tipo = 'C'" _
               & " where C.cod_convenio = '" & vCodLinea & "' and C.ACTIVO = 1"
Else
        strSQL = "select C.cod_destino as 'CODIGO',C.descripcion" _
               & " from CATALOGO_DESTINOS C" _
               & "  left join RADAR_COLOCADORES_LINKS X On C.cod_destino = X.codigo and X.Tipo = 'C'" _
               & " where C.cod_destino = '" & vCodLinea & "'"
End If

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
    fxLineaDescrip = rs!Descripcion
Else
   fxLineaDescrip = ""
End If
rs.Close
End Function



Private Sub txtLineaCod_KeyDown(KeyCode As Integer, Shift As Integer)

If optLinea.Item(0).Value = True Then
    If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select COD_CONVENIO,DESCRIPCION from CRD_CONVENIOS"
        gBusquedas.Columna = "COD_CONVENIO"
        gBusquedas.Filtro = " and ACTIVO = 1"
        gBusquedas.Orden = "DESCRIPCION"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        frmBusquedas.Show vbModal
        
        txtLineaCod.Text = Trim(gBusquedas.Resultado)
        txtLineaDesc.Text = Trim(gBusquedas.Resultado2)
        

    End If
Else
    If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select rtrim(cod_destino) as codigo ,descripcion from CATALOGO_DESTINOS"
        gBusquedas.Columna = "cod_destino"
        gBusquedas.Filtro = ""
        gBusquedas.Orden = "prioridad"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        frmBusquedas.Show vbModal
        
        txtLineaCod.Text = Trim(gBusquedas.Resultado)
        txtLineaDesc.Text = Trim(gBusquedas.Resultado2)

    End If

End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtLineaDesc.SetFocus

End Sub

Private Sub txtLineaCod_LostFocus()
  If txtLineaCod.Text <> "" Then txtLineaDesc.Text = fxLineaDescrip(txtLineaCod.Text, vTipoLinea)
End Sub

Private Sub txtLineaDesc_KeyDown(KeyCode As Integer, Shift As Integer)

If optLinea.Item(0).Value = True Then
    If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select COD_CONVENIO,DESCRIPCION from CRD_CONVENIOS"
        gBusquedas.Columna = "DESCRIPCION"
        gBusquedas.Filtro = " and ACTIVO = 1"
        gBusquedas.Orden = "DESCRIPCION"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        frmBusquedas.Show vbModal
        
        txtLineaCod.Text = Trim(gBusquedas.Resultado)
        txtLineaDesc.Text = Trim(gBusquedas.Resultado2)
        

    End If
Else
    If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select rtrim(cod_destino) as codigo, descripcion from CATALOGO_DESTINOS"
        gBusquedas.Columna = "descripcion"
        gBusquedas.Filtro = ""
        gBusquedas.Orden = "prioridad"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        frmBusquedas.Show vbModal
        
        txtLineaCod.Text = Trim(gBusquedas.Resultado)
        txtLineaDesc.Text = Trim(gBusquedas.Resultado2)

    End If

End If

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select COLOCADOR_ID,NOMBRE from RADAR_COLOCADORES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtRs_01_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRs_02.SetFocus
End Sub

Private Sub txtRs_02_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus
End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelFax.SetFocus
End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRs_01.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail2.SetFocus
End Sub


Private Sub txtTelFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtWebSite.SetFocus
End Sub

Private Sub txtWebSite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEMail.SetFocus
End Sub
