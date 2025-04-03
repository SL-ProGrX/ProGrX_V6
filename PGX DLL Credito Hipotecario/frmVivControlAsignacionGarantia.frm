VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmVivControlAsignacionGarantia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de asignación de garantias"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   7.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10950
   Begin XtremeSuiteControls.RadioButton optTipoProfesional 
      Height          =   252
      Index           =   0
      Left            =   7680
      TabIndex        =   42
      Top             =   120
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Ingenieros"
      Transparent     =   -1  'True
      Appearance      =   2
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton optTipoProfesional 
      Height          =   252
      Index           =   1
      Left            =   9240
      TabIndex        =   43
      Top             =   120
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Abogados"
      Transparent     =   -1  'True
      Appearance      =   2
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   7320
      Width           =   5895
   End
   Begin VB.TextBox txtUltimaNota 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   6600
      Width           =   6135
   End
   Begin TabDlg.SSTab sstabGeneral 
      Height          =   5832
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10692
      _ExtentX        =   18865
      _ExtentY        =   10292
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Asignación de garantía"
      TabPicture(0)   =   "frmVivControlAsignacionGarantia.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Linea(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lvwGarantias"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvwProfesionales"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Entrega Garantías"
      TabPicture(1)   =   "frmVivControlAsignacionGarantia.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboEntregaProf"
      Tab(1).Control(1)=   "lvwEntregas"
      Tab(1).Control(2)=   "lblEtiqueta(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Recepción Documentos"
      TabPicture(2)   =   "frmVivControlAsignacionGarantia.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboRecibeProf"
      Tab(2).Control(1)=   "lvwRecibidas"
      Tab(2).Control(2)=   "lblEtiqueta(2)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Registro Información Avaluó"
      TabPicture(3)   =   "frmVivControlAsignacionGarantia.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cboRegistroProf"
      Tab(3).Control(1)=   "lvwRegistro"
      Tab(3).Control(2)=   "lblAyuda"
      Tab(3).Control(3)=   "lblEtiqueta(0)"
      Tab(3).ControlCount=   4
      Begin VB.ComboBox cboRegistroProf 
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
         ItemData        =   "frmVivControlAsignacionGarantia.frx":0070
         Left            =   -73440
         List            =   "frmVivControlAsignacionGarantia.frx":0072
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   4092
      End
      Begin VB.ComboBox cboRecibeProf 
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
         ItemData        =   "frmVivControlAsignacionGarantia.frx":0074
         Left            =   -73320
         List            =   "frmVivControlAsignacionGarantia.frx":0076
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   4932
      End
      Begin VB.ComboBox cboEntregaProf 
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
         ItemData        =   "frmVivControlAsignacionGarantia.frx":0078
         Left            =   -73320
         List            =   "frmVivControlAsignacionGarantia.frx":007A
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   4932
      End
      Begin VB.Frame Frame5 
         Caption         =   "Información personal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   5295
         Left            =   -75000
         TabIndex        =   1
         Top             =   480
         Width           =   9000
         Begin VB.CommandButton cmdGuadarDuennos 
            Caption         =   "&Guardar"
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
            Left            =   7560
            TabIndex        =   11
            Top             =   4800
            Width           =   1215
         End
         Begin VB.TextBox txtIdentificacion 
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
            Left            =   840
            MaxLength       =   15
            TabIndex        =   10
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   840
            MaxLength       =   30
            TabIndex        =   9
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox Text12 
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
            Left            =   5340
            MaxLength       =   4
            TabIndex        =   8
            ToolTipText     =   "Código de la provincia de residencia"
            Top             =   360
            Width           =   540
         End
         Begin VB.TextBox Text11 
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
            Left            =   5340
            MaxLength       =   4
            TabIndex        =   7
            ToolTipText     =   "Código del cantón de residencia"
            Top             =   720
            Width           =   540
         End
         Begin VB.TextBox Text10 
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
            Left            =   5340
            MaxLength       =   4
            TabIndex        =   6
            ToolTipText     =   "Código del distrito de residencia"
            Top             =   1080
            Width           =   540
         End
         Begin VB.TextBox Text9 
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
            Left            =   5895
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Descripción de la provincia de residencia"
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox Text8 
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
            Left            =   5895
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Descripción del cantón de residencia"
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox Text13 
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
            Left            =   5895
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Descripción del distrito de residencia"
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5340
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1440
            Width           =   3495
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2535
            Left            =   240
            TabIndex        =   12
            Top             =   2160
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   4471
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label 
            Caption         =   "Cédula:"
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
            Index           =   7
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label 
            Caption         =   "Nombre:"
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
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label 
            Caption         =   "Dirección por señas:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   0
            Left            =   4440
            TabIndex        =   17
            Top             =   1440
            Width           =   885
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Distrito"
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
            Index           =   4
            Left            =   4440
            TabIndex        =   16
            Top             =   1155
            Width           =   615
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Cantón"
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
            Index           =   5
            Left            =   4440
            TabIndex        =   15
            Top             =   795
            Width           =   615
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Provincia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   4440
            TabIndex        =   14
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Dirección residencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   13
            Top             =   120
            Width           =   2775
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   3240
            Picture         =   "frmVivControlAsignacionGarantia.frx":007C
            Top             =   360
            Width           =   480
         End
      End
      Begin MSComctlLib.ListView lvwProfesionales 
         Height          =   2532
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   10428
         _ExtentX        =   18389
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwGarantias 
         Height          =   2172
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   10368
         _ExtentX        =   18283
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilsIconosLista"
         SmallIcons      =   "ilsIconosLista"
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwRecibidas 
         Height          =   4692
         Left            =   -74880
         TabIndex        =   26
         Top             =   960
         Width           =   10428
         _ExtentX        =   18389
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilsIconosLista"
         SmallIcons      =   "ilsIconosLista"
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwEntregas 
         Height          =   4692
         Left            =   -74880
         TabIndex        =   28
         Top             =   960
         Width           =   10428
         _ExtentX        =   18389
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilsIconosLista"
         SmallIcons      =   "ilsIconosLista"
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwRegistro 
         Height          =   4692
         Left            =   -74880
         TabIndex        =   30
         Top             =   960
         Width           =   10428
         _ExtentX        =   18389
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilsIconosLista"
         SmallIcons      =   "ilsIconosLista"
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
         NumItems        =   0
      End
      Begin VB.Label lblAyuda 
         Caption         =   "Haga doble clic sobre el item para registrar la información del avaluo"
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
         Height          =   492
         Left            =   -69120
         TabIndex        =   32
         Top             =   480
         Width           =   4572
      End
      Begin VB.Label lblEtiqueta 
         Alignment       =   1  'Right Justify
         Caption         =   "Profesional"
         Height          =   168
         Index           =   0
         Left            =   -74760
         TabIndex        =   31
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label lblEtiqueta 
         Alignment       =   1  'Right Justify
         Caption         =   "Profesional"
         Height          =   168
         Index           =   2
         Left            =   -74760
         TabIndex        =   27
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label lblEtiqueta 
         Alignment       =   1  'Right Justify
         Caption         =   "Profesional"
         Height          =   168
         Index           =   1
         Left            =   -74760
         TabIndex        =   24
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Profesionales"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   228
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   2448
      End
      Begin VB.Line Linea 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   2160
         X2              =   6960
         Y1              =   2910
         Y2              =   2910
      End
   End
   Begin MSComctlLib.Toolbar tlbDetalle 
      Height          =   810
      Left            =   9840
      TabIndex        =   38
      Top             =   6720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1429
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      Style           =   1
      ImageList       =   "ilstblDetalle"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "notas"
            Object.ToolTipText     =   "Ver notas"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilstblDetalle 
      Left            =   10680
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivControlAsignacionGarantia.frx":0946
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsIconosLista 
      Left            =   10680
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivControlAsignacionGarantia.frx":3DD8
            Key             =   "Asingacion"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivControlAsignacionGarantia.frx":3EF6
            Key             =   "Verde"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivControlAsignacionGarantia.frx":4014
            Key             =   "Amarillo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVivControlAsignacionGarantia.frx":413A
            Key             =   "Rojo"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNumeroFinca 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   1680
      TabIndex        =   41
      Top             =   7188
      Width           =   1812
   End
   Begin VB.Label lblNumOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   1680
      TabIndex        =   40
      Top             =   6888
      Width           =   1812
   End
   Begin VB.Label lblEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   1680
      TabIndex        =   39
      Top             =   6600
      Width           =   1812
   End
   Begin VB.Label Label4 
      Caption         =   "No. de Finca"
      Height          =   252
      Left            =   480
      TabIndex        =   37
      Top             =   7188
      Width           =   1332
   End
   Begin VB.Label Label3 
      Caption         =   "No. Operación"
      Height          =   252
      Left            =   480
      TabIndex        =   34
      Top             =   6888
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Estado"
      Height          =   252
      Left            =   480
      TabIndex        =   33
      Top             =   6600
      Width           =   732
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   -120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11172
   End
End
Attribute VB_Name = "frmVivControlAsignacionGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_AsignaGarantia As Boolean

'Para Listas
Private vItem As ListItem
Public ItemSeleccionado As ListItem
Public ItemTemp As ListItem
Private vProfesional As String

Private vCargaCboEntregar As Boolean 'Control de evento clik del combo cboEntregaProf
Private vCargaCboRecibir As Boolean 'Control de evento clik del combo cboProfRecibir
Private vCargaRegistroProf As Boolean 'Control de evento clik del combo cboRegistroProf

Private vOperacionG As String

Private Type vParametros

  'Tiempo para validar Abogados
  gTMaxEntregaAbogado As Integer
  gTAlertaEntregaAbogado As Integer
  gTMaxFirmasAbogado As Integer
  gTAlertaFirmasAbogado As Integer
  gTMaxInscripcionAbogado As Integer
  gTAlertaInscripcionAbogado As Integer
  
  'Tiempo para validar Ingenieros
  gTMaxEntregaIngeniero As Integer
  gTAlertaEntregaIngeniero As Integer
  gTMaxRecepcionIngeniero As Integer
  gTAlertaRecepcionIngeniero As Integer
  gTMaxRegistroIngeniero As Integer
  gTAlertaRegistroIngeniero As Integer
End Type

Private vLocales As vParametros
Dim strSQL As String, rs As New ADODB.Recordset


'Tab Asignacion-----------------------------------
Private Sub sbListaProfesionalesxZona(ByVal pTipoProfesional As String, ByVal pIdZona As Integer, _
                                      ByVal pIdGarantia As Long)
Dim vKey As String

On Error GoTo vError
  
With lvwProfesionales
    .ColumnHeaders.Clear
    .ListItems.Clear

    .ColumnHeaders.Add , , "Identificación", 1400
    .ColumnHeaders.Add , , "Nombre", 3000
    .ColumnHeaders.Add , , "Prefesional", 1000
    .ColumnHeaders.Add , , "Empresa", 2500
    .ColumnHeaders.Add , , "Cant.Oper.", 1000, 1
    .ColumnHeaders.Add , , "Monto", 1500, 1
    .ColumnHeaders.Add , , "Condicion", 0
    
End With

        
strSQL = "exec spCRDVivTraerProfAsingaGarantia " & pIdZona & ",'" & pTipoProfesional & "'," & pIdGarantia
Call OpenRecordSet(rs, strSQL)
If Not glogon.error Then
    
    Do While Not rs.EOF
    
        vKey = "(VV)" & Trim(rs!idZona) _
               & "(Iz)" & Trim(rs!IdContacto) _
               & "(Ic)" & Trim(pIdGarantia) _
               & "(Ig)" & Trim(rs!Identificacion) _
               & "(Id)" & Trim(rs!IdEmpresa) _
               & "(Em)" & Trim(rs!TipoProfesional) & "(Tp)"
        
               
        Set vItem = lvwProfesionales.ListItems.Add(, vKey, rs!Identificacion)
            vItem.SubItems(1) = Trim(rs!Nombre)
            vItem.SubItems(2) = Trim(rs!Profesional)
            vItem.SubItems(3) = Trim(rs!NombreEmpresa)
            vItem.SubItems(4) = IIf((rs!CantOp = -1), 0, rs!CantOp)
            vItem.SubItems(5) = IIf(IsNull(rs!montoapr), 0, Format(rs!montoapr, "Standard"))  'Monto
            vItem.SubItems(6) = "N"
        
        If rs!ItemAsignado <> -1 Then '' -1 es cuando no tiene asignado ningún trámite
           vItem.Checked = True
           vItem.SubItems(6) = "M"
        End If

        If vItem.Checked Then
           vItem.ForeColor = vbBlue
           vItem.ListSubItems(1).ForeColor = vbBlue
        End If
         
    
       rs.MoveNext
    Loop
    rs.Close
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub sbListaGarantias()
Dim vKey As String

On Error GoTo vError

Call sblimpiarInfoObsrvaciones

With lvwGarantias
    .ColumnHeaders.Clear
    .ListItems.Clear
    
    .ColumnHeaders.Add , , "No. Operación", 1300 '0
    .ColumnHeaders.Add , , "Monto", 1500, 1 '0
    .ColumnHeaders.Add , , "Cédula", 1500 '1
    .ColumnHeaders.Add , , "Nombre", 3000 '2
    .ColumnHeaders.Add , , "Número Finca", 2000 '3
    .ColumnHeaders.Add , , "Plano catastro", 1500 '4
    .ColumnHeaders.Add , , "Tipo Derecho", 1000 '5
    .ColumnHeaders.Add , , "Grado Hiporteca", 2000 '6
    .ColumnHeaders.Add , , "Área Finca(m2)", 1500, 1 '7
    .ColumnHeaders.Add , , "Zona", 1500  '7
    .ColumnHeaders.Add , , "Usuario Registro", 1500 '8
    .ColumnHeaders.Add , , "Fecha Registro", 1500 '9
    .ColumnHeaders.Add , , "IdZona", 0 '9
    
End With

strSQL = "SELECT vGarantia.NumeroOperacion, S.CEDULA, S.NOMBRE, RCR.MONTOAPR, vGarantia.IdGarantia" _
       & ",vGarantia.UbicacionCanton, vGarantia.UbicacionDistrito, vGarantia.IdZona" _
       & ",vGarantia.NumeroFinca, vGarantia.TipoDerecho, vGarantia.NumPlanoCatastro" _
       & ",vGarantia.GradoHipoteca, vGarantia.AreaFinca, vGarantia.Estado" _
       & ",vGarantia.Direccion, vGarantia.AnotacionesFinca, vGarantia.Gravamenes" _
       & ",vGarantia.AnotacionesGravamen, vGarantia.ObservacionAvaluo, vGarantia.RegistroUsuario" _
       & ",vGarantia.RegistroFecha, vZonas.Descripcion AS DescZona, P.DESCRIPCION AS DescProvincia" _
       & ",C.DESCRIPCION AS DescCanton, D.DESCRIPCION AS DescDistrito" _
       & ",CASE VGarantia.GradoHipoteca       WHEN 'P' THEN 'Primer Grado'" _
       & "      WHEN 'S' THEN 'Segundo Grado' WHEN 'T' THEN 'Tercer Grado' ELSE '' END AS DescGradoHipoteca" _
       & " FROM  SOCIOS S " _
       & " INNER JOIN REG_CREDITOS RCR ON S.CEDULA = RCR.CEDULA" _
       & " INNER JOIN ViviendaGarantia vGarantia ON vGarantia.NumeroOperacion = RCR.ID_SOLICITUD" _
       & " INNER JOIN PROVINCIAS AS P ON vGarantia.UbicacionProvincia = P.PROVINCIA" _
       & " INNER JOIN CANTONES C ON C.PROVINCIA = vGarantia.UbicacionProvincia and C.CANTON = vGarantia.UbicacionCanton " _
       & "  LEFT OUTER JOIN DISTRITOS AS D ON vGarantia.UbicacionProvincia = D.PROVINCIA" _
       & "         AND vGarantia.UbicacionCanton = D.CANTON AND vGarantia.UbicacionDistrito = D.DISTRITO" _
       & " INNER JOIN ViviendaZonas vZonas ON vZonas.IdZona = vGarantia.IdZona" _

                                
If vProfesional = "I" Then
    strSQL = strSQL & " where  RCR.ESTADOSOL not in ('F','N') and VGarantia.idGarantia" _
                    & " not in(select VGT.idGarantia from ViviendaGarantiaTramite as VGT where VGT.tipo = '" & vProfesional & "')"
                    ' and ViviendaGarantia.estado = 'S'
Else
strSQL = strSQL & " INNER JOIN ViviendaGarantiaTramite AS VGT ON vGarantia.IdGarantia = VGT.IdGarantia" _
                & " WHERE (RCR.ESTADOSOL NOT IN ('F', 'N')) and VGT.Estado = 'R'" _
                & " and VGarantia.idGarantia not in (select GT.idGarantia from ViviendaGarantiaTramite as GT where GT.tipo = '" & vProfesional & "')"
End If


Call OpenRecordSet(rs, strSQL)
If Not glogon.error Then

Do While Not rs.EOF
    vKey = "(VV)" & Trim(rs!NumeroOperacion) _
           & "(Op)" & Trim(rs!IdGarantia) _
           & "(Ig)" & Trim(rs!NumeroFinca) & "(Nf)"
    
    Set vItem = lvwGarantias.ListItems.Add(, vKey, rs!NumeroOperacion, , gIconoLista)
        vItem.SubItems(1) = Format(rs!montoapr, "Standard")
        vItem.SubItems(2) = Trim(rs!cedula)
        vItem.SubItems(3) = Trim(rs!Nombre)
        vItem.SubItems(4) = Trim(rs!NumeroFinca)
        vItem.SubItems(5) = Trim(rs!NumPlanoCatastro)
        vItem.SubItems(6) = Trim(rs!TipoDerecho)
        vItem.SubItems(7) = Trim(rs!DescGradoHipoteca)
        vItem.SubItems(8) = Trim(rs!AreaFinca)
        vItem.SubItems(9) = Trim(rs!DescZona)
        vItem.SubItems(10) = Trim(rs!RegistroUsuario)
        vItem.SubItems(11) = Format(rs!RegistroFecha, "dd-mm-yyyy")
        vItem.SubItems(12) = rs!idZona
       rs.MoveNext
Loop
rs.Close

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub sbListaEntregaGarantias(ByVal pIdContacto As Long)
Dim vKey As String

On Error GoTo vError


With lvwEntregas

    .ColumnHeaders.Clear
    .ListItems.Clear
    
    .ColumnHeaders.Add , , "No. Operación", 1300 '0
    .ColumnHeaders.Add , , "Monto", 1500, 1 '0
    .ColumnHeaders.Add , , "Cédula", 1500 '1
    .ColumnHeaders.Add , , "Nombre", 3000 '2
    .ColumnHeaders.Add , , "Número Finca", 2000 '3
    .ColumnHeaders.Add , , "Plano catastro", 1500 '4
    .ColumnHeaders.Add , , "Tipo Derecho", 1000 '5
    .ColumnHeaders.Add , , "Grado Hiporteca", 2000 '6
    .ColumnHeaders.Add , , "Área Finca(m2)", 1500, 1 '7
    .ColumnHeaders.Add , , "Zona", 1500  '7
    .ColumnHeaders.Add , , "Usuario Asigna ", 1500 '8
    .ColumnHeaders.Add , , "Fecha Asinación", 1500 '9
    .ColumnHeaders.Add , , "IdZona", 0 '9
    .ColumnHeaders.Add , , "Condicion", 0 '10
    .ColumnHeaders.Add , , "Tiempo Profesional", 2000 '11
    .ColumnHeaders.Add , , "Tiempo Ejecutivo", 2000 '12
    
End With


strSQL = "SELECT ViviendaGarantiaTramite.IdContacto,ViviendaGarantia.NumeroOperacion, SOCIOS.CEDULA, SOCIOS.NOMBRE, REG_CREDITOS.MONTOAPR, ViviendaGarantiaTramite.IdGarantia, " & _
        " ViviendaGarantia.UbicacionCanton, ViviendaGarantia.UbicacionDistrito, ViviendaGarantia.IdZona, ViviendaGarantia.NumeroFinca, " & _
        " ViviendaGarantia.TipoDerecho, ViviendaGarantia.NumPlanoCatastro, ViviendaGarantia.GradoHipoteca, ViviendaGarantia.AreaFinca, " & _
        " ViviendaGarantia.Estado, ViviendaGarantia.RegistroUsuario, ViviendaGarantia.RegistroFecha, ViviendaZonas.Descripcion AS DescZona, " & _
        "Case ViviendaGarantia.GradoHipoteca " & _
        " WHEN 'P' THEN 'Primer Grado' " & _
        " WHEN 'S' THEN 'Segundo Grado' " & _
        " WHEN 'T' THEN 'Tercer Grado' ELSE '' END AS DescGradoHipoteca," & _
        " ViviendaGarantiaTramite.AsignacionUsuario, ViviendaGarantiaTramite.AsignacionFecha, " & _
        " ViviendaGarantiaTramite.EntregaUsuario, ViviendaGarantiaTramite.EntregaFecha, " & _
        " DATEDIFF(day, ViviendaGarantiaTramite.AsignacionFecha, dbo.MyGetdate()) as diasTransProfesional " & _
        " FROM ViviendaGarantiaTramite INNER JOIN " & _
        " ViviendaGarantia ON ViviendaGarantiaTramite.IdGarantia = ViviendaGarantia.IdGarantia INNER JOIN " & _
        " REG_CREDITOS ON ViviendaGarantia.NumeroOperacion = REG_CREDITOS.ID_SOLICITUD INNER JOIN " & _
        " SOCIOS ON REG_CREDITOS.CEDULA = SOCIOS.CEDULA INNER JOIN " & _
        " ViviendaZonas ON ViviendaGarantia.IdZona = ViviendaZonas.IdZona " & _
        " and ViviendaGarantiaTramite.AsignacionFecha is not null and ViviendaGarantiaTramite.AsignacionUsuario is not null " & _
        " and ViviendaGarantiaTramite.EntregaFecha is null " & _
        " where  ViviendaGarantiaTramite.IdContacto = " & pIdContacto & " And ViviendaGarantiaTramite.Tipo = '" & vProfesional & "'"

Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then

  Do While Not rs.EOF
    vKey = "(VV)" & Trim(rs!NumeroOperacion) _
           & "(Op)" & Trim(rs!IdGarantia) _
           & "(Ig)" & Trim(rs!NumeroFinca) & "(Nf)"
       
    gIconoLista = "Verde"
    
    If vProfesional = "I" Then 'Validación para Ingenieros
        If rs!diasTransProfesional = vLocales.gTAlertaEntregaIngeniero Then
            gIconoLista = "Amarillo"
        ElseIf rs!diasTransProfesional >= vLocales.gTMaxEntregaIngeniero Then
            gIconoLista = "Rojo"
        End If
    Else 'Validación para Abogados
        If rs!diasTransProfesional = vLocales.gTAlertaEntregaAbogado Then
            gIconoLista = "Amarillo"
        ElseIf rs!diasTransProfesional >= vLocales.gTMaxEntregaAbogado Then
            gIconoLista = "Rojo"
        End If
    End If
    
    Set vItem = lvwEntregas.ListItems.Add(, vKey, rs!NumeroOperacion, , gIconoLista)
        vItem.SubItems(1) = Format(rs!montoapr, "Standard")
        vItem.SubItems(2) = Trim(rs!cedula)
        vItem.SubItems(3) = Trim(rs!Nombre)
        vItem.SubItems(4) = Trim(rs!NumeroFinca)
        vItem.SubItems(5) = Trim(rs!NumPlanoCatastro)
        vItem.SubItems(6) = Trim(rs!TipoDerecho)
        vItem.SubItems(7) = Trim(rs!DescGradoHipoteca)
        vItem.SubItems(8) = Trim(rs!AreaFinca)
        vItem.SubItems(9) = Trim(rs!DescZona)
        vItem.SubItems(10) = Trim(rs!AsignacionUsuario)
        vItem.SubItems(11) = Format(rs!AsignacionFecha, "dd-mm-yyyy")
        vItem.SubItems(12) = rs!idZona
        vItem.SubItems(13) = "N"
        
        If Not IsNull(rs!EntregaFecha) Then
            vItem.Checked = True
            vItem.SubItems(13) = "M"
        End If
        
        vItem.SubItems(14) = IIf(IsNull(rs!diasTransProfesional), 0, rs!diasTransProfesional)
       
       rs.MoveNext
    Loop
    rs.Close

End If 'Rs

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub sbListaRecibeGarantias(ByVal pIdContacto As Long)
Dim vKey As String

On Error GoTo vError


With lvwRecibidas
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add , , "No. Operación", 1300 '0
    .ColumnHeaders.Add , , "Monto", 1500, 1 '1
    .ColumnHeaders.Add , , "Cédula", 1500 '2
    .ColumnHeaders.Add , , "Nombre", 3000 '3
    .ColumnHeaders.Add , , "Número Finca", 2000 '4
    .ColumnHeaders.Add , , "Plano catastro", 1500 '5
    .ColumnHeaders.Add , , "Tipo Derecho", 1000 '6
    .ColumnHeaders.Add , , "Grado Hiporteca", 2000 '7
    .ColumnHeaders.Add , , "Área Finca(m2)", 1500, 1 '8
    .ColumnHeaders.Add , , "Zona", 1500  '9
    .ColumnHeaders.Add , , "Entrega Usuario", 1500 '10
    .ColumnHeaders.Add , , "Entrega Fecha", 1500 '11
    .ColumnHeaders.Add , , "IdZona", 0 '12
    .ColumnHeaders.Add , , "Condicion", 0 '13
    .ColumnHeaders.Add , , "Tiempo Profesional", 2000 '14
    .ColumnHeaders.Add , , "Tiempo Ejecutivo", 2000 '15

End With

strSQL = "SELECT ViviendaGarantiaTramite.IdContacto,ViviendaGarantia.NumeroOperacion, SOCIOS.CEDULA, SOCIOS.NOMBRE, REG_CREDITOS.MONTOAPR, ViviendaGarantiaTramite.IdGarantia, " & _
                "ViviendaGarantia.UbicacionCanton, ViviendaGarantia.UbicacionDistrito, ViviendaGarantia.IdZona, ViviendaGarantia.NumeroFinca, " & _
                "ViviendaGarantia.TipoDerecho, ViviendaGarantia.NumPlanoCatastro, ViviendaGarantia.GradoHipoteca, ViviendaGarantia.AreaFinca, " & _
                "ViviendaGarantia.Estado, ViviendaGarantia.RegistroUsuario, ViviendaGarantia.RegistroFecha, ViviendaZonas.Descripcion AS DescZona, " & _
                "Case ViviendaGarantia.GradoHipoteca " & _
                "WHEN 'P' THEN 'Primer Grado' " & _
                "WHEN 'S' THEN 'Segundo Grado' " & _
                "WHEN 'T' THEN 'Tercer Grado' ELSE '' END AS DescGradoHipoteca, " & _
                "ViviendaGarantiaTramite.RecepcionUsuario, ViviendaGarantiaTramite.RecepcionFecha, " & _
                "ViviendaGarantiaTramite.EntregaUsuario, ViviendaGarantiaTramite.EntregaFecha, " & _
                "DATEDIFF(day, ViviendaGarantiaTramite.EntregaFecha, dbo.MyGetdate()) as diasTransProfesional " & _
                "FROM ViviendaGarantiaTramite INNER JOIN " & _
                "ViviendaGarantia ON ViviendaGarantiaTramite.IdGarantia = ViviendaGarantia.IdGarantia INNER JOIN " & _
                "REG_CREDITOS ON ViviendaGarantia.NumeroOperacion = REG_CREDITOS.ID_SOLICITUD INNER JOIN " & _
                "SOCIOS ON REG_CREDITOS.CEDULA = SOCIOS.CEDULA INNER JOIN " & _
                "ViviendaZonas ON ViviendaGarantia.IdZona = ViviendaZonas.IdZona " & _
                "and  ViviendaGarantiaTramite.EntregaFecha is not null and ViviendaGarantiaTramite.EntregaUsuario is not null "
                
If vProfesional = "A" Then
   strSQL = strSQL & "and ViviendaGarantiaTramite.FirmasFecha is null "
Else
    strSQL = strSQL & "and ViviendaGarantiaTramite.RecepcionFecha is null "
End If
              
strSQL = strSQL & " where ViviendaGarantiaTramite.IdContacto = " & pIdContacto _
       & " And ViviendaGarantiaTramite.Tipo = '" & vProfesional & "'"

Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then

Do While Not rs.EOF
    vKey = "(VV)" & Trim(rs!NumeroOperacion) _
           & "(Op)" & Trim(rs!IdGarantia) _
           & "(Ig)" & Trim(rs!NumeroFinca) & "(Nf)"
           
    gIconoLista = "Verde"
    
    If vProfesional = "I" Then 'Validación para Ingenieros
        If rs!diasTransProfesional = vLocales.gTAlertaRecepcionIngeniero Then
            gIconoLista = "Amarillo"
        ElseIf rs!diasTransProfesional >= vLocales.gTMaxRecepcionIngeniero Then
            gIconoLista = "Rojo"
        End If
    Else 'Validación para Abogados
        If rs!diasTransProfesional = vLocales.gTAlertaFirmasAbogado Then
            gIconoLista = "Amarillo"
        ElseIf rs!diasTransProfesional >= vLocales.gTMaxFirmasAbogado Then
            gIconoLista = "Rojo"
        End If
    End If
    
    Set vItem = lvwRecibidas.ListItems.Add(, vKey, rs!NumeroOperacion, , gIconoLista)
        vItem.SubItems(1) = Format(rs!montoapr, "Standard")
        vItem.SubItems(2) = Trim(rs!cedula)
        vItem.SubItems(3) = Trim(rs!Nombre)
        vItem.SubItems(4) = Trim(rs!NumeroFinca)
        vItem.SubItems(5) = Trim(rs!NumPlanoCatastro)
        vItem.SubItems(6) = Trim(rs!TipoDerecho)
        vItem.SubItems(7) = Trim(rs!DescGradoHipoteca)
        vItem.SubItems(8) = Trim(rs!AreaFinca)
        vItem.SubItems(9) = Trim(rs!DescZona)
        vItem.SubItems(10) = Trim(rs!EntregaUsuario)
        vItem.SubItems(11) = Format(rs!EntregaFecha, "dd-mm-yyyy")
        vItem.SubItems(12) = rs!idZona
        vItem.SubItems(13) = "N"
        
        If Not IsNull(rs!RecepcionFecha) Then
            vItem.Checked = True
            vItem.SubItems(13) = "M"
        End If
        vItem.SubItems(14) = IIf(IsNull(rs!diasTransProfesional), 0, rs!diasTransProfesional)
       
 rs.MoveNext
 Loop
 rs.Close
 
End If 'Rs

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub sbListaRegistroGarantias(ByVal pIdContacto As Long)
Dim vKey As String

On Error GoTo vError

With lvwRegistro
    .ColumnHeaders.Clear
    .ListItems.Clear
    
    .ColumnHeaders.Add , , "No. Operación", 1300 '0
    .ColumnHeaders.Add , , "Monto", 1500, 1 '0
    .ColumnHeaders.Add , , "Cédula", 1500 '1
    .ColumnHeaders.Add , , "Nombre", 3000 '2
    .ColumnHeaders.Add , , "Número Finca", 2000 '3
    .ColumnHeaders.Add , , "Plano catastro", 1500 '4
    .ColumnHeaders.Add , , "Tipo Derecho", 1000 '5
    .ColumnHeaders.Add , , "Grado Hiporteca", 2000 '6
    .ColumnHeaders.Add , , "Área Finca(m2)", 1500, 1 '7
    .ColumnHeaders.Add , , "Zona", 1500  '7
    .ColumnHeaders.Add , , "Usuario Asigna ", 1500 '8
    .ColumnHeaders.Add , , "Fecha Asinación", 1500 '9
    .ColumnHeaders.Add , , "IdZona", 0 '10
    .ColumnHeaders.Add , , "Condicion", 0 '11
    .ColumnHeaders.Add , , "Tiempo Profesional", 2000 '12
    .ColumnHeaders.Add , , "Tiempo Ejecutivo", 2000 '13
    
End With

'If ObjConsultar.fxTraerRegistroGarantias(pIdContacto, vProfesional) Then

strSQL = "SELECT ViviendaGarantiaTramite.IdContacto,ViviendaGarantia.NumeroOperacion, SOCIOS.CEDULA, SOCIOS.NOMBRE, REG_CREDITOS.MONTOAPR, ViviendaGarantiaTramite.IdGarantia, " & _
                "ViviendaGarantia.UbicacionCanton, ViviendaGarantia.UbicacionDistrito, ViviendaGarantia.IdZona, ViviendaGarantia.NumeroFinca, " & _
                "ViviendaGarantia.TipoDerecho, ViviendaGarantia.NumPlanoCatastro, ViviendaGarantia.GradoHipoteca, ViviendaGarantia.AreaFinca, " & _
                "ViviendaGarantia.Estado, ViviendaGarantia.RegistroUsuario, ViviendaGarantiaTramite.RegistroFecha, ViviendaZonas.Descripcion AS DescZona, " & _
                "Case ViviendaGarantia.GradoHipoteca " & _
                "WHEN 'P' THEN 'Primer Grado' " & _
                "WHEN 'S' THEN 'Segundo Grado' " & _
                "WHEN 'T' THEN 'Tercer Grado' ELSE '' END AS DescGradoHipoteca," & _
                "ViviendaGarantiaTramite.AsignacionUsuario, ViviendaGarantiaTramite.AsignacionFecha, " & _
                "ViviendaGarantiaTramite.EntregaUsuario, ViviendaGarantiaTramite.EntregaFecha, " & _
                "DATEDIFF(day, ViviendaGarantiaTramite.RecepcionFecha, dbo.MyGetdate()) as diasTransProfesional, " & _
                "DATEDIFF(day, ViviendaGarantiaTramite.firmasFecha, dbo.MyGetdate()) as diasTransAbogado, firmasFecha " & _
                "FROM ViviendaGarantiaTramite INNER JOIN " & _
                "ViviendaGarantia ON ViviendaGarantiaTramite.IdGarantia = ViviendaGarantia.IdGarantia INNER JOIN " & _
                "REG_CREDITOS ON ViviendaGarantia.NumeroOperacion = REG_CREDITOS.ID_SOLICITUD INNER JOIN " & _
                "SOCIOS ON REG_CREDITOS.CEDULA = SOCIOS.CEDULA INNER JOIN " & _
                "ViviendaZonas ON ViviendaGarantia.IdZona = ViviendaZonas.IdZona " & _
                "and ViviendaGarantiaTramite.AsignacionFecha is not null " & _
                "and ViviendaGarantiaTramite.EntregaFecha is not null " & _
                "and ViviendaGarantiaTramite.RegistroFecha is null "
             
strSQL = strSQL & "where  ViviendaGarantiaTramite.IdContacto = " & pIdContacto & " And ViviendaGarantiaTramite.Tipo = '" & vProfesional & "'"
               

Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then


Do While Not rs.EOF
    vKey = "(VV)" & Trim(rs!NumeroOperacion) _
           & "(Op)" & Trim(rs!IdGarantia) _
           & "(Ig)" & Trim(rs!NumeroFinca) & "(Nf)"
           
    gIconoLista = "Verde"
    
    If vProfesional = "I" Then 'Validación para Ingenieros
        If rs!diasTransProfesional = vLocales.gTAlertaRegistroIngeniero Then
            gIconoLista = "Amarillo"
        ElseIf rs!diasTransProfesional >= vLocales.gTMaxRegistroIngeniero Then
            gIconoLista = "Rojo"
        End If
    Else 'Validación para Abogados
        If rs!diasTransAbogado = vLocales.gTAlertaInscripcionAbogado Then
            gIconoLista = "Amarillo"
        ElseIf rs!diasTransAbogado >= vLocales.gTMaxInscripcionAbogado Then
            gIconoLista = "Rojo"
        End If
    End If

    Set vItem = lvwRegistro.ListItems.Add(, vKey, rs!NumeroOperacion, , gIconoLista)
        vItem.SubItems(1) = Format(rs!montoapr, "Standard")
        vItem.SubItems(2) = Trim(rs!cedula)
        vItem.SubItems(3) = Trim(rs!Nombre)
        vItem.SubItems(4) = Trim(rs!NumeroFinca)
        vItem.SubItems(5) = Trim(rs!NumPlanoCatastro)
        vItem.SubItems(6) = Trim(rs!TipoDerecho)
        vItem.SubItems(7) = Trim(rs!DescGradoHipoteca)
        vItem.SubItems(8) = Trim(rs!AreaFinca)
        vItem.SubItems(9) = Trim(rs!DescZona)
        vItem.SubItems(10) = Trim(rs!AsignacionUsuario)
        vItem.SubItems(11) = Format(rs!AsignacionFecha, "dd-mm-yyyy")
        vItem.SubItems(12) = rs!idZona
        vItem.SubItems(13) = IIf(IsNull(rs!diasTransProfesional), 0, rs!diasTransProfesional)
        If Not IsNull(rs!firmasFecha) Then
            vItem.SubItems(13) = IIf(IsNull(rs!diasTransAbogado), 0, rs!diasTransAbogado)
        End If
        
 rs.MoveNext
 Loop
 rs.Close
 
End If 'rs

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub cboEntregaProf_Click()

Call sblimpiarInfoObsrvaciones

If vCargaCboEntregar Then Exit Sub

Call sbListaEntregaGarantias(cboEntregaProf.ItemData(cboEntregaProf.ListIndex))

End Sub

Private Sub sblimpiarInfoObsrvaciones()
    txtUltimaNota.Text = Empty
    lblEstado.Caption = Empty
    txtUsuario.Text = Empty
    lblNumOperacion.Caption = Empty
    lblNumeroFinca.Caption = Empty
End Sub

Private Sub cboRecibeProf_Click()

Call sblimpiarInfoObsrvaciones

If vCargaCboRecibir Then Exit Sub

Call sbListaRecibeGarantias(cboRecibeProf.ItemData(cboRecibeProf.ListIndex))

End Sub

Private Sub cboRegistroProf_Click()

Call sblimpiarInfoObsrvaciones

If vCargaRegistroProf Then Exit Sub

Call sbListaRegistroGarantias(cboRegistroProf.ItemData(cboRegistroProf.ListIndex))

End Sub



Private Sub Form_Activate()

Call sbCargarTimeposProfesional(vProfesional)

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

'Select Case SSTabGeneral.Tab
'    Case 0
'        Call sstabGeneral_Click(0)
'    Case 1
'        Call sstabGeneral_Click(1)
'    Case 2
'        Call sstabGeneral_Click(2)
'    Case 3 ' Tab para el registro de garantias
'        Call sstabGeneral_Click(3)
'
'End Select

End Sub

Private Sub Form_Load()
vModulo = 3

'' Carga nombre de la ternimal
If Len(glogon.Maquina) = 0 Then
    Call sbMaquina
End If

gIconoLista = "Asingacion"

Call optTipoProfesional_Click(0)

sstabGeneral.Tab = 0
vProfesional = "I"

Call sstabGeneral_Click(0)


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub




Private Sub lvwEntregas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vIdGarantia As Long
Dim vIdcontacto As Long

On Error GoTo vError

vIdGarantia = fxDeCodePK(Item.Key, 5, "(Op)")
vIdGarantia = fxDeCodePK(Item.Key, gPosIni, "(Ig)")
vIdcontacto = cboEntregaProf.ItemData(cboEntregaProf.ListIndex)

If lvwEntregas.ListItems.Count > 0 Then
    If (Item.Checked) And (Item.ListSubItems(13).Text = "N") Then
        
        strSQL = "exec spCRDVivEntregaGarantia_M " & vIdGarantia & "," & vIdcontacto & ",'" & glogon.Usuario & "','S'"
        Call ConectionExecute(strSQL)
        
        If Not glogon.error Then
            Item.ListSubItems(13).Text = "M"
            Call Bitacora("APLICA", "Entrega Garantía Hipotecaria " & vIdGarantia & " Contacto: " & vIdcontacto)
            'Item.ForeColor = vbBlue
        End If
        
    ElseIf (Item.Checked = False) And (Item.ListSubItems(13).Text = "M") Then
        
        strSQL = "exec spCRDVivEntregaGarantia_M " & vIdGarantia & "," & vIdcontacto & ",'" & glogon.Usuario & "','N'"
        Call ConectionExecute(strSQL)
        
        If Not glogon.error Then
            Item.ListSubItems(13).Text = "N"
            Call Bitacora("BORRA", "Entrega Garantía Hipotecaria " & vIdGarantia & " Contacto: " & vIdcontacto)
        End If
    End If
End If
    
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub lvwEntregas_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vTemp As String
Set ItemSeleccionado = Item

vTemp = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
Call sbCargarInfoNota(fxDeCodePK(Item.Key, gPosIni, "(Ig)"), vProfesional)

End Sub

Private Sub sbCargarInfoNota(ByVal pIdGarantia As Long, ByVal pProfesional As String)

Call sblimpiarInfoObsrvaciones

strSQL = "SELECT VGT.IdNota, VGT.Nota," & _
    "    DescEstado = CASE VGT.Estado " & _
    "    WHEN 'R' THEN 'Garantía Registrada' " & _
    "    WHEN 'X' THEN 'Proceso de avaluo' " & _
    "    WHEN 'A' THEN 'Avaluo Registrado' " & _
    "    WHEN 'Y' THEN'Proceso de registro' " & _
    "    WHEN 'S' THEN 'Solicitada' ELSE '' END, " & _
    "    VGT.Usuario,CONVERT(nvarchar(30), VGT.Fecha, 103) as FechaRegistro, VGT.IdGarantia, VGT.IdContacto, VGT.Tipo, " & _
    "    VGT.Estado,  vcontactos.Nombre, vGarantia.NumeroOperacion, vGarantia.NumeroFinca " & _
    "FROM   ViviendaGarantiaTramiteNotas AS VGT INNER JOIN " & _
    "ViviendaContactos AS vcontactos ON VGT.IdContacto = vcontactos.IdContacto " & _
    "INNER JOIN ViviendaGarantia  as vGarantia ON VGT.IdGarantia = vGarantia.IdGarantia " & _
    "WHERE  (VGT.IdGarantia = " & pIdGarantia & ") AND (VGT.Tipo = '" & pProfesional & "') " & _
    "order by VGT.Fecha desc"
                
Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then
    If Not rs.EOF Then
        txtUltimaNota.Text = rs!Nota
        lblEstado.Caption = rs!DescEstado
        txtUsuario.Text = rs!Usuario & "                " & rs!FechaRegistro
        lblNumOperacion.Caption = rs!NumeroOperacion
        lblNumeroFinca = rs!NumeroFinca
    End If
End If

End Sub

Private Sub lvwGarantias_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vTemp As String, vIdZona As Long

Set ItemSeleccionado = Item

If Not ItemTemp Is Nothing Then
    If ItemTemp.Key <> ItemSeleccionado.Key Then
        Call sbListaGarantias
        lvwProfesionales.Sorted = False
        Call fxSelecionarItem(ItemSeleccionado)
'
        vIdZona = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
        Call sbListaProfesionalesxZona(vProfesional, ItemSeleccionado.ListSubItems(12), fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Ig)"))
        Set ItemTemp = Nothing
    Else
        lvwProfesionales.Sorted = False
        vIdZona = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
        Call sbListaProfesionalesxZona(vProfesional, Item.SubItems(12), fxDeCodePK(Item.Key, gPosIni, "(Ig)"))
    End If
Else
    lvwProfesionales.Sorted = False
    vIdZona = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
    
    Call sbListaProfesionalesxZona(vProfesional, Item.SubItems(12), fxDeCodePK(Item.Key, gPosIni, "(Ig)"))
End If
    
vTemp = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
vOperacionG = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")

Call sbCargarInfoNota(fxDeCodePK(Item.Key, gPosIni, "(Ig)"), vProfesional)
   
   
End Sub

'Private Function fxValidaMarcarSoloUno() As Boolean
'On Error GoTo vError
'Dim i As Long
'Dim vEncontrados As Integer
'Dim vReturn As Boolean
'
'vReturn = False
'Set ItemTemp = Nothing
'For i = lvwProfesionales.ListItems.Count To 1 Step -1
'    If lvwProfesionales.ListItems(i).Checked Then
'        vEncontrados = vEncontrados + 1
'        Set ItemTemp = ItemSeleccionado
'        If vEncontrados > 1 Then
'            lvwProfesionales.ListItems(i).Checked = False
'            vReturn = True
'            Exit For
'        End If
'    End If
'Next i
'
'
'fxValidaMarcarSoloUno = vReturn
'
'salir:
'    Exit Function
'vError:
'    fxValidaMarcarSoloUno = vReturn
'    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'    Resume salir
'
'End Function

Private Function fxValidaMarcarSoloUno() As Boolean
Dim i As Long, vEncontrados As Integer

On Error GoTo vError

    fxValidaMarcarSoloUno = True
    For i = lvwProfesionales.ListItems.Count To 1 Step -1
        If lvwProfesionales.ListItems(i).Checked Then
            vEncontrados = vEncontrados + 1
            If vEncontrados > 1 Then
                fxValidaMarcarSoloUno = False
                Exit Function
            End If
        End If
    Next i

    Exit Function
vError:
    fxValidaMarcarSoloUno = False
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Function fxSelecionarItem(ByVal pItem As MSComctlLib.ListItem) As Boolean
Dim i As Long, vReturn As Boolean

On Error GoTo vError

vReturn = False

For i = lvwGarantias.ListItems.Count To 1 Step -1
    If lvwGarantias.ListItems(i).Key = pItem.Key Then
        lvwGarantias.ListItems(i).Selected = False
        Exit For
    End If
Next i

fxSelecionarItem = vReturn

Exit Function

vError:
    fxSelecionarItem = vReturn
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Function


Private Sub lvwProfesionales_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)


On Error GoTo vError
    
    lvwProfesionales.SortKey = ColumnHeader.Index - 1
    
    If (lvwProfesionales.SortOrder = lvwAscending) Then
        lvwProfesionales.SortOrder = lvwDescending
    Else
        lvwProfesionales.SortOrder = lvwAscending
    End If
    
    lvwProfesionales.Sorted = True
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Function fxHonorarios_Registra(ByVal pIdGarantia As Long) As Boolean

On Error GoTo vError

fxHonorarios_Registra = False
                
strSQL = "SELECT RegistraCalHonorarios,RegistraCalHonorariosDt" _
       & " FROM  ViviendaGarantia" _
       & " where IdGarantia = " & pIdGarantia
Call OpenRecordSet(rs, strSQL)
                       
If Not glogon.error Then
        If rs!RegistraCalHonorarios = 1 And rs!RegistraCalHonorariosDT = 1 Then
            fxHonorarios_Registra = True
        End If
End If

Exit Function

vError:
MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Function


Private Sub lvwProfesionales_DblClick()
Dim vOperacion As String, vIdcontacto As Long
Dim vIdZona As Long, vIdGarantia As Long
Dim vFecha As Date

On Error GoTo vError
        
    If ItemSeleccionado Is Nothing Then Exit Sub
    If lvwProfesionales.ListItems.Count = 0 Then Exit Sub
    If vProfesional = "I" Then Exit Sub
    
    vIdGarantia = fxDeCodePK(ItemSeleccionado.Key, 5, "(Iz)")
    vIdcontacto = fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Ic)")
    vIdGarantia = fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Ig)")
    gOperacion = Trim(vOperacionG)
    vFecha = fxFechaServidor
    
    m_AsignaGarantia = False
    
    If fxHonorarios_Registra(vIdGarantia) Then
        'Manual - Detalle

            GLOBALES.gTag = vIdcontacto
            GLOBALES.gTag2 = vIdGarantia
            GLOBALES.gTag3 = vProfesional
            
            Call sbSIFForms("frmVivHonorariosDetalle", 1, , , False)
            
            If m_AsignaGarantia = True Then
                sstabGeneral.Tab = 1
            End If
    
    Else
       'Automatico
        If optTipoProfesional.Item(1).Value Then
            If (MsgBox("¿ Confirma que desea realiza el recibo de la garantía inscrita (Automatico).?", vbQuestion + vbYesNo) = vbYes) Then
               
                    If ItemSeleccionado.ListSubItems(6).Text = "N" Then
                        strSQL = "exec spCRDVivAsingaGarantia_A " & vIdGarantia & "," & vIdcontacto & ",'" & vProfesional _
                               & "','" & glogon.Usuario & "','" & Format(vFecha, "yyyy/mm/dd") & "'"
                            'pAsignacionFecha
                        Call ConectionExecute(strSQL)
                        If Not glogon.error Then
                            m_AsignaGarantia = True
                            Call Bitacora("APLICA", "Asignación Garantía Hipotecaria " & vIdGarantia & " Contacto: " & vIdcontacto)
                            
                           MsgBox "Información fue registrada corretamente.", vbInformation
                           sstabGeneral.Tab = 1
                        End If
                    
                    End If 'ItemSeleccionado
                
             End If 'box
             
        End If 'optTipoProfesional.Item(1).Value
    End If

    Exit Sub
    
vError:
        MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


'TODO
Private Sub lvwProfesionales_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vIdZona As Long
Dim vIdGarantia As Long
Dim vIdcontacto As Long

On Error GoTo vError

If Not fxValidaMarcarSoloUno Then
    Item.Checked = False
    MsgBox "Solo puede asignar la garantía a un profesional"
    Exit Sub
End If

vIdGarantia = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
vIdGarantia = fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Ig)")

If lvwProfesionales.ListItems.Count > 0 Then
    If (Item.Checked) And (Item.ListSubItems(6).Text = "N") Then
        vIdZona = fxDeCodePK(Item.Key, 5, "(Iz)")
        vIdcontacto = fxDeCodePK(Item.Key, gPosIni, "(Ic)")
        
        strSQL = "exec spCRDVivAsingaGarantia_A " & vIdGarantia & "," & vIdcontacto & ",'" & vProfesional _
               & "','" & glogon.Usuario & "','" & Format(fxFechaServidor, "yyyy/mm/dd") & "'"
        'pAsignacionFecha
        Call ConectionExecute(strSQL)
        If Not glogon.error Then
            
            Item.ListSubItems(6).Text = "M"
            vIdZona = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
            
            Call sbListaProfesionalesxZona(vProfesional, ItemSeleccionado.SubItems(12), fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Ig)"))
            
            Item.ForeColor = vbBlue
            
            Call Bitacora("APLICA", "Asignación Garantía Hipotecaria " & vIdGarantia & " Contacto: " & vIdcontacto)
            
        End If

    ElseIf (Item.Checked = False) And (Item.ListSubItems(6).Text = "M") Then
        vIdZona = fxDeCodePK(Item.Key, 5, "(Iz)")
        vIdcontacto = fxDeCodePK(Item.Key, gPosIni, "(Ic)")
        
        strSQL = "exec spCRDVivAsingaGarantia_B " & vIdGarantia & "," & vIdcontacto
        
        Call ConectionExecute(strSQL)
        If Not glogon.error Then
            Item.ListSubItems(6).Text = "N"
            vIdZona = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
            
            Call sbListaProfesionalesxZona(vProfesional, ItemSeleccionado.SubItems(12), fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Ig)"))
            Call Bitacora("BORRAR", "Asignación Garantía Hipotecaria " & vIdGarantia & " Contacto: " & vIdcontacto)
            
        End If
    End If
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub lvwProfesionales_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If lvwProfesionales.Checkboxes = False Then
        Set ItemSeleccionado = Item
    End If

End Sub

Private Sub lvwRecibidas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vIdZona As Long
Dim vIdGarantia As Long
Dim vIdcontacto As Long


On Error GoTo vError

Set ItemSeleccionado = Item

vIdGarantia = fxDeCodePK(Item.Key, 5, "(Op)")
vIdGarantia = fxDeCodePK(Item.Key, gPosIni, "(Ig)")

If (Item.Checked) And (Item.ListSubItems(13).Text = "N") Then

    vIdcontacto = cboRecibeProf.ItemData(cboRecibeProf.ListIndex)
    If optTipoProfesional.Item(0).Value Then 'Para ingenieros
        strSQL = "exec spCRDVivRecepcionGarantia_M " & vIdGarantia & "," & vIdcontacto & ",'" & glogon.Usuario & "','S'"
        Call ConectionExecute(strSQL)
        
        If Not glogon.error Then
            Item.ListSubItems(13).Text = "M"
            Call Bitacora("APLICA", "Recepción Garantía Hipotecaria " & vIdGarantia & " Contacto: " & vIdcontacto)
        End If
    
    Else 'Seleccionado abogados
        strSQL = "exec spCRDVivCtlAsignacionGarantia " & vIdGarantia & "," & vIdcontacto & ",'" & glogon.Usuario & "','S','A','F'"
        Call ConectionExecute(strSQL)
        
        If Not glogon.error Then
            Item.ListSubItems(13).Text = "M"
            Call Bitacora("APLICA", "Asignación Garantía Hipotecaria(F): " & vIdGarantia & " Contacto: " & vIdcontacto)
        End If
    
    End If
    

ElseIf (Item.Checked = False) And (Item.ListSubItems(13).Text = "M") Then
 
    vIdcontacto = cboRecibeProf.ItemData(cboRecibeProf.ListIndex)
    If optTipoProfesional.Item(0).Value Then 'Para ingenieros
        strSQL = "exec spCRDVivRecepcionGarantia_M " & vIdGarantia & "," & vIdcontacto & ",'" & glogon.Usuario & "','N'"
        Call ConectionExecute(strSQL)
        
        If Not glogon.error Then
            Item.ListSubItems(13).Text = "N"
            Call Bitacora("BORRA", "Recepción Garantía Hipotecaria " & vIdGarantia & " Contacto: " & vIdcontacto)
        End If
    
    Else 'Seleccionado abogados
        strSQL = "exec spCRDVivCtlAsignacionGarantia " & vIdGarantia & "," & vIdcontacto & ",'" & glogon.Usuario & "','N','A','F'"
        Call ConectionExecute(strSQL)
        
        If Not glogon.error Then
            Item.ListSubItems(13).Text = "N"
            Call Bitacora("BORRA", "Asignación Garantía Hipotecaria (F): " & vIdGarantia & " Contacto: " & vIdcontacto)
        End If
    End If
    
End If
    
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    
End Sub


Private Sub lvwRecibidas_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim vTemp As String

Set ItemSeleccionado = Item

vTemp = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
Call sbCargarInfoNota(fxDeCodePK(Item.Key, gPosIni, "(Ig)"), vProfesional)

End Sub

Private Sub lvwRegistro_DblClick()
Dim vOperacion As String
Dim vIdGarantia As Long
Dim vIdcontacto As Long


If ItemSeleccionado Is Nothing Then Exit Sub
If lvwRegistro.ListItems.Count = 0 Then Exit Sub
If vProfesional = "A" Then Exit Sub



vOperacion = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
vIdGarantia = fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Ig)")
vIdcontacto = cboRegistroProf.ItemData(cboRegistroProf.ListIndex)

'frmVivRegistroAvaluo.vNumOperacion = vOperacion
'frmVivRegistroAvaluo.vIdGarantia = vIdGarantia
'frmVivRegistroAvaluo.vIdcontacto = vIdcontacto

GLOBALES.gTag = vOperacion
GLOBALES.gTag2 = vIdGarantia
GLOBALES.gTag3 = vIdcontacto

Call sbSIFForms("frmVivRegistroAvaluo", 1, , , False)


DoEvents
Call sstabGeneral_Click(3)

End Sub

Private Sub lvwRegistro_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim vIdZona As Long
Dim vIdGarantia As Long
Dim vIdcontacto As Long

On Error GoTo vError

vIdGarantia = fxDeCodePK(Item.Key, 5, "(Op)")
vIdGarantia = fxDeCodePK(Item.Key, gPosIni, "(Ig)")

If Item.Checked Then

    vIdcontacto = cboRegistroProf.ItemData(cboRegistroProf.ListIndex)
    gOperacion = Item.Text
    

    If (MsgBox("¿ Confirma que desea realiza el recibo de la garantía inscrita.?", vbQuestion + vbYesNo) = vbYes) Then

' Este procedimiento lo realiza desde el sp de Asignacion en la BD.
'                    glogon.strSQL = "exec spCRDViviendaDesembolsoPendiente " & vIdcontacto & ",'A'," & vIdGarantia & ",'" & glogon.usuario & "'"
'                    glogon.Conection.Execute glogon.strSQL
                   
        strSQL = "exec spCRDVivCtlAsignacionGarantia " & vIdGarantia & "," & vIdcontacto & ",'" & glogon.Usuario & "','S','A','I'"
        Call ConectionExecute(strSQL)
        
        If Not glogon.error Then
            Call Bitacora("APLICA", "Asignación Garantía Hipotecaria (I): " & vIdGarantia & " Contacto: " & vIdcontacto)
            MsgBox "Información registrada satisfactoriamente!", vbInformation
        End If
    End If

    Call sstabGeneral_Click(3)

End If
    
Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub lvwRegistro_ItemClick(ByVal Item As MSComctlLib.ListItem)

Dim vTemp As String
Set ItemSeleccionado = Item

vTemp = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
Call sbCargarInfoNota(fxDeCodePK(Item.Key, gPosIni, "(Ig)"), vProfesional)
End Sub
Private Sub sbCambiarNombreTab()
If optTipoProfesional.Item(0).Value Then 'Ingenieros
    sstabGeneral.TabCaption(0) = "Asignación de garantía"
    sstabGeneral.TabCaption(1) = "Entrega Garantías"
    sstabGeneral.TabCaption(2) = "Recepción Documentos"
    sstabGeneral.TabCaption(3) = "Registro Información Avaluó"
Else 'Abogados
    sstabGeneral.TabCaption(0) = "Asignación de garantía"
    sstabGeneral.TabCaption(1) = "Entrega de garantía"
    sstabGeneral.TabCaption(2) = "Firma de Garantía"
    sstabGeneral.TabCaption(3) = "Recibo de garantías Inscritas"
    

End If

End Sub

Private Sub optTipoProfesional_Click(Index As Integer)

Call sbCambiarNombreTab

Select Case Index
    Case 0 'Ingeniero
            vProfesional = "I"
    Case 1 'Abogado
            vProfesional = "A"
End Select


Call sstabGeneral_Click(0)

lblAyuda.Visible = True
If optTipoProfesional.Item(1).Value Then
    lblAyuda.Visible = False
End If

Call sbCargarTimeposProfesional(vProfesional)


End Sub

Private Sub sstabGeneral_Click(PreviousTab As Integer)
Set ItemSeleccionado = Nothing
Call sblimpiarInfoObsrvaciones

Select Case sstabGeneral.Tab

    Case 0
        gIconoLista = "Asingacion"
        lvwProfesionales.ColumnHeaders.Clear
        lvwProfesionales.ListItems.Clear
        
        If optTipoProfesional.Item(0).Value Then 'Para ingenieros
            Me.lvwProfesionales.Checkboxes = True
        Else
            Me.lvwProfesionales.Checkboxes = False
        End If
        
        Call sbListaGarantias
        
    Case 1
        vCargaCboEntregar = True
        Call sbCargaLista(cboEntregaProf, "ENTREGA_PROF", "where ViviendaGarantiaTramite.tipo = '" & vProfesional & "'")
        vCargaCboEntregar = False
        If cboEntregaProf.ListCount > 0 Then
            Call sbListaEntregaGarantias(cboEntregaProf.ItemData(cboEntregaProf.ListIndex))
        Else
            lvwEntregas.ColumnHeaders.Clear
            lvwEntregas.ListItems.Clear
        End If
    Case 2
        vCargaCboRecibir = True
        If optTipoProfesional.Item(0).Value Then 'Para ingenieros
            Call sbCargaLista(cboRecibeProf, "RECEPCION_PROF", "where ViviendaGarantiaTramite.tipo = '" & vProfesional & "'")
        Else 'Para Abogados
            Call sbCargaLista(cboRecibeProf, "FIRMAS_PROF", "where ViviendaGarantiaTramite.tipo = '" & vProfesional & "'")
        End If
        vCargaCboRecibir = False
        If cboRecibeProf.ListCount > 0 Then
            Call sbListaRecibeGarantias(cboRecibeProf.ItemData(cboRecibeProf.ListIndex))
        Else
            lvwRecibidas.ColumnHeaders.Clear
            lvwRecibidas.ListItems.Clear
        End If

    Case 3 ' Tab para el registro de garantias
        vCargaRegistroProf = True
        If optTipoProfesional.Item(0).Value Then 'Para ingenieros
            Me.lvwRegistro.Checkboxes = False
            Call sbCargaLista(cboRegistroProf, "REGISTRO_PROF", "where ViviendaGarantiaTramite.tipo = '" & vProfesional & "'")
        Else 'Para Abogados
            Me.lvwRegistro.Checkboxes = True
            Call sbCargaLista(cboRegistroProf, "RECIBO_PROF", "where ViviendaGarantiaTramite.tipo = '" & vProfesional & "'")
        End If
        
        vCargaRegistroProf = False
        If cboRegistroProf.ListCount > 0 Then
            Call sbListaRegistroGarantias(cboRegistroProf.ItemData(cboRegistroProf.ListIndex))
        Else
            lvwRegistro.ColumnHeaders.Clear
            lvwRegistro.ListItems.Clear
        End If
        
        
End Select

End Sub

Private Function fxExisteUnItemMarcado(ByRef pIdContacto As Long) As Boolean
Dim i As Long, vEncontrados As Integer, vReturn As Boolean

On Error GoTo vError

vReturn = False
Set ItemTemp = Nothing

For i = lvwProfesionales.ListItems.Count To 1 Step -1
    If lvwProfesionales.ListItems(i).Checked Then
        pIdContacto = fxDeCodePK(lvwProfesionales.ListItems(i).Key, 5, "(Iz)")
        pIdContacto = fxDeCodePK(lvwProfesionales.ListItems(i).Key, gPosIni, "(Ic)")
        vReturn = True
        Exit For
    End If
Next i


fxExisteUnItemMarcado = vReturn

Exit Function

vError:
    fxExisteUnItemMarcado = vReturn
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Function

Private Sub tlbDetalle_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim vIdcontacto As Long

If Not ItemSeleccionado Is Nothing Then
    frmVivTramiteNotas.m_NumOperacion = fxDeCodePK(ItemSeleccionado.Key, 5, "(Op)")
    frmVivTramiteNotas.m_IdGarantia = fxDeCodePK(ItemSeleccionado.Key, gPosIni, "(Ig)")
    frmVivTramiteNotas.m_Profesional = "A"
    
    If optTipoProfesional.Item(0).Value Then
        frmVivTramiteNotas.m_Profesional = "I"
    End If
    
    Select Case sstabGeneral.Tab

    Case 0
        If Not fxExisteUnItemMarcado(vIdcontacto) Then
            MsgBox "Debe seleccionar un profesional antes de agregar una nota", vbExclamation
            If lvwProfesionales.Enabled Then lvwProfesionales.SetFocus
            Exit Sub
        Else
            frmVivTramiteNotas.m_IdContacto = vIdcontacto
        End If
    Case 1
        frmVivTramiteNotas.m_IdContacto = cboEntregaProf.ItemData(cboEntregaProf.ListIndex)
    Case 2
        frmVivTramiteNotas.m_IdContacto = cboRecibeProf.ItemData(cboRecibeProf.ListIndex)
    Case 3 ' Tab para el registro de garantias
        frmVivTramiteNotas.m_IdContacto = cboRegistroProf.ItemData(cboRegistroProf.ListIndex)
        
End Select
    frmVivTramiteNotas.Show vbModal, Me
'    Call sbSIFForms("frmVivTramiteNotas", 1, , , False)
    Call lvwEntregas_ItemClick(ItemSeleccionado)
End If

End Sub

Private Sub sbCargarTimeposProfesional(ByVal pProfesional As String)
On Error GoTo vError
'Carga los tiempo para seguimiento de garantía

'Tiempo para validar Abogados
vLocales.gTMaxEntregaAbogado = 0
vLocales.gTAlertaEntregaAbogado = 0
vLocales.gTMaxFirmasAbogado = 0
vLocales.gTAlertaFirmasAbogado = 0
vLocales.gTMaxInscripcionAbogado = 0
vLocales.gTAlertaInscripcionAbogado = 0

'Tiempo para validar Ingenieros
vLocales.gTMaxEntregaIngeniero = 0
vLocales.gTAlertaEntregaIngeniero = 0
vLocales.gTMaxRecepcionIngeniero = 0
vLocales.gTAlertaRecepcionIngeniero = 0
vLocales.gTMaxRegistroIngeniero = 0
vLocales.gTAlertaRegistroIngeniero = 0

strSQL = "select Profesional, Proceso, TiempoMaximo, TiempoAlerta from ViviendaTiemposSeguimiento " _
       & " where profesional = '" & pProfesional & "'"
Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then
    Do While Not rs.EOF
        If rs!Profesional = "A" Then 'Tiempo para validar Abogados
            Select Case Trim(rs!Proceso)
                Case "E"
                    vLocales.gTMaxEntregaAbogado = rs!TiempoMaximo
                    vLocales.gTAlertaEntregaAbogado = rs!TiempoAlerta
                Case "F"
                    vLocales.gTMaxFirmasAbogado = rs!TiempoMaximo
                    vLocales.gTAlertaFirmasAbogado = rs!TiempoAlerta
                Case "I"
                    vLocales.gTMaxInscripcionAbogado = rs!TiempoMaximo
                    vLocales.gTAlertaInscripcionAbogado = rs!TiempoAlerta
            End Select
        ElseIf rs!Profesional = "I" Then 'Tiempo para validar Ingenieros
            Select Case Trim(rs!Proceso)
                Case "E"
                    vLocales.gTMaxEntregaIngeniero = rs!TiempoMaximo
                    vLocales.gTAlertaEntregaIngeniero = rs!TiempoAlerta
                Case "R"
                    vLocales.gTMaxRecepcionIngeniero = rs!TiempoMaximo
                    vLocales.gTAlertaRecepcionIngeniero = rs!TiempoAlerta
                Case "X"
                    vLocales.gTMaxRegistroIngeniero = rs!TiempoMaximo
                    vLocales.gTAlertaRegistroIngeniero = rs!TiempoAlerta
            End Select
        End If
    rs.MoveNext
    Loop
    rs.Close
End If

Exit Sub

vError:
  
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub
