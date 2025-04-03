VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_SeguimientoRevisionesTag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisión de Afiliaciones"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAF_SeguimientoRevisionesTag.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   11490
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   10800
      Top             =   120
   End
   Begin VB.Frame FraControles 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   11175
      Begin TabDlg.SSTab SSTab1 
         Height          =   6615
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   11668
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
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
         TabCaption(0)   =   "Afiliaciones"
         TabPicture(0)   =   "frmAF_SeguimientoRevisionesTag.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "tlbRefresh"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "imgRefresh"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "vGrid"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Detalle"
         TabPicture(1)   =   "frmAF_SeguimientoRevisionesTag.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtEstadoPersona"
         Tab(1).Control(1)=   "txtSexo"
         Tab(1).Control(2)=   "txtEstado"
         Tab(1).Control(3)=   "txtSector"
         Tab(1).Control(4)=   "txtProfesion"
         Tab(1).Control(5)=   "txtHijos"
         Tab(1).Control(6)=   "txtBoleta"
         Tab(1).Control(7)=   "txtCodPromotor"
         Tab(1).Control(8)=   "txtNombrePromotor"
         Tab(1).Control(9)=   "ssTabSubX"
         Tab(1).Control(10)=   "dtpFechaIngreso"
         Tab(1).Control(11)=   "dtpNacimiento"
         Tab(1).Control(12)=   "Label5"
         Tab(1).Control(13)=   "Label6"
         Tab(1).Control(14)=   "Label1(0)"
         Tab(1).Control(15)=   "Label14"
         Tab(1).Control(16)=   "Label15(0)"
         Tab(1).Control(17)=   "Label10(1)"
         Tab(1).Control(18)=   "Label10(2)"
         Tab(1).Control(19)=   "Label12(0)"
         Tab(1).Control(20)=   "Label12(1)"
         Tab(1).Control(21)=   "Label13"
         Tab(1).ControlCount=   22
         TabCaption(2)   =   "Seguimiento"
         TabPicture(2)   =   "frmAF_SeguimientoRevisionesTag.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "vGridSeguimiento"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Revisión"
         TabPicture(3)   =   "frmAF_SeguimientoRevisionesTag.frx":68A6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cboEtiquetas"
         Tab(3).Control(1)=   "txtObservacion"
         Tab(3).Control(2)=   "tlbAplicar"
         Tab(3).Control(3)=   "lswErrores"
         Tab(3).Control(4)=   "Label27"
         Tab(3).Control(5)=   "Label2(0)"
         Tab(3).Control(6)=   "Label8(1)"
         Tab(3).ControlCount=   7
         Begin VB.TextBox txtEstadoPersona 
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
            Left            =   -73560
            MaxLength       =   15
            TabIndex        =   24
            ToolTipText     =   "Campo para la Cédula de Identidad"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtSexo 
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
            Left            =   -70920
            MaxLength       =   15
            TabIndex        =   23
            ToolTipText     =   "Campo para la Cédula de Identidad"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtEstado 
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
            Left            =   -73560
            MaxLength       =   15
            TabIndex        =   22
            ToolTipText     =   "Campo para la Cédula de Identidad"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtSector 
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
            Left            =   -68160
            MaxLength       =   15
            TabIndex        =   21
            ToolTipText     =   "Campo para la Cédula de Identidad"
            Top             =   1920
            Width           =   1815
         End
         Begin VB.TextBox txtProfesion 
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
            Left            =   -73560
            MaxLength       =   45
            TabIndex        =   20
            ToolTipText     =   "Campo para la Cédula de Identidad"
            Top             =   1590
            Width           =   4335
         End
         Begin VB.TextBox txtHijos 
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
            Left            =   -68160
            MaxLength       =   15
            TabIndex        =   19
            ToolTipText     =   "Campo para la Cédula de Identidad"
            Top             =   1560
            Width           =   1815
         End
         Begin VB.TextBox txtBoleta 
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
            Left            =   -70920
            MaxLength       =   15
            TabIndex        =   18
            Text            =   "1"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtCodPromotor 
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
            Left            =   -73560
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtNombrePromotor 
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
            Left            =   -72600
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1920
            Width           =   3375
         End
         Begin VB.ComboBox cboEtiquetas 
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
            ItemData        =   "frmAF_SeguimientoRevisionesTag.frx":68C2
            Left            =   -73320
            List            =   "frmAF_SeguimientoRevisionesTag.frx":68C4
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   5295
         End
         Begin VB.TextBox txtObservacion 
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
            Height          =   1575
            Left            =   -73320
            MaxLength       =   995
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   4080
            Width           =   9135
         End
         Begin FPSpreadADO.fpSpread vGrid 
            Height          =   5655
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   10815
            _Version        =   524288
            _ExtentX        =   19076
            _ExtentY        =   9975
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
            MaxCols         =   7
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "frmAF_SeguimientoRevisionesTag.frx":68C6
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridSeguimiento 
            Height          =   5775
            Left            =   -74760
            TabIndex        =   9
            Top             =   600
            Width           =   10575
            _Version        =   524288
            _ExtentX        =   18653
            _ExtentY        =   10186
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
            MaxCols         =   487
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "frmAF_SeguimientoRevisionesTag.frx":73BF
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin MSComctlLib.Toolbar tlbAplicar 
            Height          =   570
            Left            =   -65760
            TabIndex        =   10
            Top             =   5880
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   1005
            ButtonWidth     =   2117
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Aplicar"
                  Key             =   "Aplicar"
                  Object.ToolTipText     =   "Aplicar Etiqueta"
                  ImageIndex      =   1
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lswErrores 
            Height          =   2655
            Left            =   -73320
            TabIndex        =   11
            Top             =   1200
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   4683
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
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
               Text            =   "Código"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Aplicado"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Mensaje"
               Object.Width           =   8819
            EndProperty
         End
         Begin TabDlg.SSTab ssTabSubX 
            Height          =   3615
            Left            =   -74520
            TabIndex        =   25
            Top             =   2640
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   6376
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
            TabPicture(0)   =   "frmAF_SeguimientoRevisionesTag.frx":79B7
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
            TabCaption(1)   =   "Trabajo"
            TabPicture(1)   =   "frmAF_SeguimientoRevisionesTag.frx":E219
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lblCentroTrabajo"
            Tab(1).Control(1)=   "lblDepartamento"
            Tab(1).Control(2)=   "lblSeccion"
            Tab(1).Control(3)=   "Label8(3)"
            Tab(1).Control(4)=   "txtCTCodigo"
            Tab(1).Control(5)=   "txtCTDesc"
            Tab(1).Control(6)=   "txtDeptDesc"
            Tab(1).Control(7)=   "txtDeptCodigo"
            Tab(1).Control(8)=   "txtSecDesc"
            Tab(1).Control(9)=   "txtSecCodigo"
            Tab(1).Control(10)=   "txtInstitucion"
            Tab(1).ControlCount=   11
            TabCaption(2)   =   "Nombramientos"
            TabPicture(2)   =   "frmAF_SeguimientoRevisionesTag.frx":14A7B
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label19"
            Tab(2).Control(1)=   "Label1(2)"
            Tab(2).Control(2)=   "Label1(1)"
            Tab(2).Control(3)=   "Label17"
            Tab(2).Control(4)=   "Label16(5)"
            Tab(2).Control(5)=   "Label16(4)"
            Tab(2).Control(6)=   "Line9(4)"
            Tab(2).Control(7)=   "Line9(5)"
            Tab(2).Control(8)=   "dtpNombramiento"
            Tab(2).Control(9)=   "lswNombramiento"
            Tab(2).Control(10)=   "txtAniosSerivicio"
            Tab(2).Control(11)=   "optNombramiento(1)"
            Tab(2).Control(12)=   "optNombramiento(0)"
            Tab(2).Control(13)=   "txtNumeroPagos"
            Tab(2).ControlCount=   14
            Begin VB.TextBox txtInstitucion 
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
               Left            =   -73560
               TabIndex        =   47
               Top             =   600
               Width           =   5175
            End
            Begin VB.TextBox txtSecCodigo 
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
               Left            =   -73560
               MaxLength       =   20
               TabIndex        =   46
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   1500
               Width           =   615
            End
            Begin VB.TextBox txtSecDesc 
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
               MaxLength       =   20
               TabIndex        =   45
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   1500
               Width           =   4575
            End
            Begin VB.TextBox txtDeptCodigo 
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
               Left            =   -73560
               MaxLength       =   20
               TabIndex        =   44
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   1140
               Width           =   615
            End
            Begin VB.TextBox txtDeptDesc 
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
               MaxLength       =   20
               TabIndex        =   43
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   1140
               Width           =   4575
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
               TabIndex        =   35
               Top             =   420
               Width           =   8055
               Begin VB.TextBox txtDistrito 
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
                  Left            =   840
                  TabIndex        =   39
                  Top             =   1080
                  Width           =   2055
               End
               Begin VB.TextBox txtCanton 
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
                  Left            =   840
                  TabIndex        =   38
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.TextBox txtProvincia 
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
                  Left            =   840
                  TabIndex        =   37
                  Top             =   360
                  Width           =   2055
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
                  Height          =   1060
                  Left            =   3000
                  MaxLength       =   100
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   36
                  ToolTipText     =   "Dirección exacta Aqui"
                  Top             =   360
                  Width           =   4935
               End
               Begin VB.Label Label9 
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
                  TabIndex        =   42
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.Label Label8 
                  Caption         =   "Canton"
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
                  TabIndex        =   41
                  Top             =   720
                  Width           =   735
               End
               Begin VB.Label Label7 
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
                  TabIndex        =   40
                  Top             =   360
                  Width           =   735
               End
            End
            Begin VB.TextBox txtApartado 
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
               Left            =   1680
               MaxLength       =   15
               TabIndex        =   34
               ToolTipText     =   "Campo para la Cédula de Identidad"
               Top             =   2580
               Width           =   5895
            End
            Begin VB.TextBox txtEmail 
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
               Left            =   1680
               MaxLength       =   45
               TabIndex        =   33
               ToolTipText     =   "Campo para la Cédula de Identidad"
               Top             =   2220
               Width           =   5895
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
               Height          =   495
               Left            =   1680
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   32
               Top             =   2940
               Width           =   5895
            End
            Begin VB.TextBox txtNumeroPagos 
               Alignment       =   2  'Center
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
               Left            =   -73080
               MaxLength       =   1
               TabIndex        =   31
               Text            =   "2"
               ToolTipText     =   "Número de Pagos Mensuales del socio"
               Top             =   2220
               Width           =   750
            End
            Begin VB.OptionButton optNombramiento 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   -73680
               TabIndex        =   30
               Top             =   1020
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.OptionButton optNombramiento 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   -73680
               TabIndex        =   29
               Top             =   1380
               Width           =   1575
            End
            Begin VB.TextBox txtAniosSerivicio 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
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
               Left            =   -73080
               MaxLength       =   45
               TabIndex        =   28
               ToolTipText     =   "Campo para la Cédula de Identidad"
               Top             =   2580
               Width           =   735
            End
            Begin VB.TextBox txtCTDesc 
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
               MaxLength       =   20
               TabIndex        =   27
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   2100
               Width           =   4575
            End
            Begin VB.TextBox txtCTCodigo 
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
               Left            =   -73560
               MaxLength       =   20
               TabIndex        =   26
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   2100
               Width           =   615
            End
            Begin MSComctlLib.ListView lswNombramiento 
               Height          =   2175
               Left            =   -72000
               TabIndex        =   48
               Top             =   900
               Width           =   4815
               _ExtentX        =   8493
               _ExtentY        =   3836
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
               Left            =   -73680
               TabIndex        =   49
               ToolTipText     =   "Fecha de Ingreso al sistema"
               Top             =   1860
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
               Format          =   290848771
               CurrentDate     =   38899
               MaxDate         =   55153
               MinDate         =   14611
            End
            Begin VB.Line Line9 
               BorderColor     =   &H00FFFFFF&
               Index           =   5
               X1              =   -72000
               X2              =   -68280
               Y1              =   780
               Y2              =   780
            End
            Begin VB.Line Line9 
               BorderColor     =   &H00FFFFFF&
               Index           =   4
               X1              =   -74880
               X2              =   -72480
               Y1              =   780
               Y2              =   780
            End
            Begin VB.Label Label8 
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
               Left            =   -74760
               TabIndex        =   63
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label lblSeccion 
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
               Left            =   -74760
               TabIndex        =   62
               Top             =   1500
               Width           =   1575
            End
            Begin VB.Label Label8 
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
               Index           =   2
               Left            =   -74760
               TabIndex        =   61
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label lblDepartamento 
               Caption         =   "Departam"
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
               Left            =   -74760
               TabIndex        =   60
               Top             =   1140
               Width           =   1335
            End
            Begin VB.Label Label11 
               Caption         =   "Apto. Postal"
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
               Left            =   480
               TabIndex        =   59
               Top             =   2580
               Width           =   975
            End
            Begin VB.Label Label10 
               Caption         =   "Email"
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
               Left            =   480
               TabIndex        =   58
               Top             =   2220
               Width           =   735
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
               Left            =   480
               TabIndex        =   57
               Top             =   2940
               Width           =   1005
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
               Left            =   -74880
               TabIndex        =   56
               Top             =   540
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
               Left            =   -72000
               TabIndex        =   55
               Top             =   540
               Width           =   1455
            End
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Pagos Mensuales"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   -74760
               TabIndex        =   54
               Top             =   2220
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Estado"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   -74760
               TabIndex        =   53
               Top             =   1020
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "A partir del"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   -74760
               TabIndex        =   52
               Top             =   1860
               Width           =   1095
            End
            Begin VB.Label Label19 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Años de Servicio"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   -74760
               TabIndex        =   51
               Top             =   2580
               Width           =   1575
            End
            Begin VB.Label lblCentroTrabajo 
               Caption         =   "Centro de Trabajo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -74760
               TabIndex        =   50
               Top             =   1980
               Width           =   1095
            End
         End
         Begin MSComCtl2.DTPicker dtpFechaIngreso 
            Height          =   315
            Left            =   -67920
            TabIndex        =   64
            ToolTipText     =   "Fecha de Ingreso al sistema"
            Top             =   1080
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   519307267
            CurrentDate     =   38899
            MaxDate         =   55153
            MinDate         =   14611
         End
         Begin MSComCtl2.DTPicker dtpNacimiento 
            Height          =   315
            Left            =   -67920
            TabIndex        =   65
            ToolTipText     =   "Fecha de Ingreso al sistema"
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   519307267
            CurrentDate     =   36059
         End
         Begin MSComctlLib.ImageList imgRefresh 
            Left            =   8760
            Top             =   240
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAF_SeguimientoRevisionesTag.frx":1B2DD
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar tlbRefresh 
            Height          =   330
            Left            =   9480
            TabIndex        =   77
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            ButtonWidth     =   1984
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imgRefresh"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Refrescar"
                  Key             =   "Refrescar"
                  Object.ToolTipText     =   "Volver a cargar la información"
                  ImageIndex      =   1
               EndProperty
            EndProperty
            MousePointer    =   1
         End
         Begin VB.Label Label5 
            Caption         =   "Estado"
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
            Left            =   -74520
            TabIndex        =   75
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Ingreso"
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
            Left            =   -68880
            TabIndex        =   74
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Nacimiento"
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
            Left            =   -68880
            TabIndex        =   73
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label14 
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
            Left            =   -71640
            TabIndex        =   72
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label15 
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
            Left            =   -74520
            TabIndex        =   71
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Profesion"
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
            Left            =   -74520
            TabIndex        =   70
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Promotor"
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
            Index           =   2
            Left            =   -74520
            TabIndex        =   69
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "# Dependi."
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
            Left            =   -69120
            TabIndex        =   68
            ToolTipText     =   "Número de Dependientes"
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Sector"
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
            Left            =   -69000
            TabIndex        =   67
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "# Boleta"
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
            Left            =   -71640
            TabIndex        =   66
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "Omisiones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   14
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Etiqueta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   -74760
            TabIndex        =   13
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Observación"
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
            Left            =   -74760
            TabIndex        =   12
            Top             =   4080
            Width           =   1455
         End
      End
   End
   Begin VB.Frame fraOperacion 
      BorderStyle     =   0  'None
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
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   11295
      Begin VB.TextBox txtId 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   76
         ToolTipText     =   "Id Registro"
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox txtCedula 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label LblOperacion 
         Caption         =   "Cedula"
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
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         ToolTipText     =   "Nombre"
         Top             =   120
         Width           =   6135
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11520
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":1B402
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":21C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":284C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":2ED28
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":3558A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":3BDEC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   10080
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":4264E
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":4276C
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":42892
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":429BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":42ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":42BE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":42CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":42E1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":42F32
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   9120
      Top             =   0
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
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":43056
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":498B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":50234
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":50352
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNombreUsuario 
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
      Left            =   2280
      TabIndex        =   15
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   1800
      Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5044E
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmAF_SeguimientoRevisionesTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCedula As String, vConsecutivo As String, vPaso As Boolean

Private Sub cboEtiquetas_Click()
If vPaso Then Exit Sub
Call sbCargarObservacion
End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()
vModulo = 8


lblNombreUsuario.Caption = glogon.Usuario

SSTab1.Tab = 0

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub lswErrores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
    If Item.SubItems(2) = "S" Then
        Item.Checked = True
        If MsgBox("El error ya fué aplicado desea agregar únicamente la nota", vbOKCancel) = vbOK Then
            If txtObservacion = Empty Then
              txtObservacion.Text = Item.SubItems(1)
            Else
              txtObservacion.Text = txtObservacion.Text & vbCrLf & " - " & Item.SubItems(1)
            End If
        End If
        Exit Sub
    End If
    
    If Item.Checked Then
    
      strSQL = "insert SIF_OMISIONESG (cedula,ID_ERROR,MODULO,CODIGO,DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO) values('" & txtCedula.Text _
             & "'," & Item.Text & ",'AFI','" & txtCedula.Text & "','" & txtId.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
      Call ConectionExecute(strSQL)
             
      strSQL = "select max(LINEA_ERR) as 'Linea' from SIF_OMISIONESG where codigo = '" & txtCedula.Text & "' and Documento = '" & txtId & "' and ID_ERROR = " & Item.Text
      Call OpenRecordSet(rs, strSQL)
          Item.Tag = rs!Linea
      rs.Close
      
      If txtObservacion = Empty Then
        txtObservacion.Text = " - " & Item.SubItems(1)
      Else
        txtObservacion.Text = txtObservacion.Text & vbCrLf & " - " & Item.SubItems(1)
      End If
      
    Else
      strSQL = "delete SIF_OMISIONESG where LINEA_ERR = " & Item.Tag
      Call ConectionExecute(strSQL)
      
      Item.Tag = ""
      
     Call sbCargarObservacion
    End If
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
   Case 2
     Call sbCargarGridSeguimiento(txtCedula.Text, txtId.Text)
   Case 3
     Call sbCargarListaErrores
     Call sbCargarCombosEtiquetas
End Select
End Sub



Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbCargaListaAfiliaciones

End Sub


Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError
    Me.MousePointer = vbHourglass
    
    If Trim(cboEtiquetas.Text) = Empty Then
        MsgBox "Debe seleccionar la etiqueta que desea plicar"
        Me.MousePointer = vbDefault
        Exit Sub
    End If

    If MsgBox("Está seguro que sea aplicar la etiqueta en las afiliaciones seleccionadas", vbExclamation + vbYesNo) = vbNo Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Call sbSIFRegistraTags(txtCedula.Text, SIFGlobal.fxCodText(cboEtiquetas.Text), txtObservacion, vConsecutivo, "AFI", txtCedula.Text, vConsecutivo)


    Call sbAplicarErrores
    Call sbCargaListaAfiliaciones
    txtCedula.SetFocus
    SSTab1.Tab = 0
    
    
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbRefresh_ButtonClick(ByVal Button As MSComctlLib.Button)
Call TimerX_Timer
End Sub

Private Sub txtCedula_GotFocus()
SSTab1.Tab = 0
vCedula = Empty
vConsecutivo = Empty

lblNombre = Empty
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False


End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Trim(txtCedula) <> "" Then Call sbCurrentRecord(txtCedula)
End Sub


Private Sub txtCedula_LostFocus()
If Trim(txtCedula) <> "" Then Call sbCurrentRecord(txtCedula)
End Sub


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

    vGrid.Col = 2
    vGrid.Row = Row

    vCedula = vGrid.Text
    vGrid.Col = 3
    lblNombre.Caption = vGrid.Text
    vGrid.Col = 7
    vConsecutivo = vGrid.Text
    txtId.Text = vConsecutivo
    
    If Len(Trim(vCedula)) > 0 Then
        Call sbCurrentRecord(vCedula)
    End If

End Sub


Private Sub sbCargaListaAfiliaciones(Optional pCedula As String = "")
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select Top 3000 'B', A.cedula, S.nombre,A.usuario,A.cod_remesa,R.usuario,A.consec" _
        & " from afi_ingresos A  inner join Socios S on A.Cedula = S.cedula " _
        & " left join AFI_REMESAS_ING R on A.cod_remesa = R.cod_remesa" _
        & " Where a.ANALISTA_REVISION  Is Null and S.EstadoActual in('S','A','P')"

If pCedula = "" Then
   strSQL = strSQL & " and A.cedula like '%" & pCedula & "%'"
End If

vPaso = True
Call sbCargaGrid(vGrid, 7, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1
vPaso = False

Me.MousePointer = vbDefault

Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCurrentRecord(vCedula As String)
Dim strSQL As String
Dim rs As New ADODB.Recordset, rsTemp As New ADODB.Recordset
Dim i As Integer, vEspacio As Integer
Dim vFechaActual As Date

On Error Resume Next

vFechaActual = fxFechaServidor

If Not fxSIFValidaCadena(vCedula) Then
   Exit Sub
End If


If Not GLOBALES.SysASEVersion Then
    strSQL = "Select S.*,Est.Descripcion as 'EstadoPersonaDesc',Est.Cod_Estado + ' - ' + Est.Descripcion as 'EstadoPersona'" _
           & ",I.descripcion as DescInst,D.descripcion as DescDept,X.descripcion as DescSec,P.nombre as Promotor,R.descripcion as ProfesionX" _
           & ",Q.descripcion as Sector,dbo.fxAFIAnioServicio(cedula,'" & Format(vFechaActual, "yyyy/mm/dd") & "') as AnioServicio" _
           & ",rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
           & ",Tid.Descripcion as TipoIdDesc,Tid.Tipo_Personeria,Soc.cod_sociedad + ' - ' + rtrim(Soc.descripcion) as 'SociedadDesc'" _
           & ",Act.cod_actividad + ' - ' + rtrim(Act.descripcion) as 'ActividadDesc',O.descripcion as Oficina" _
           & " From socios S inner join Instituciones I on S.cod_institucion = I.cod_institucion" _
           & " left join AFDepartamentos D on S.cod_institucion = D.cod_institucion and S.cod_departamento = D.cod_departamento" _
           & " left join AFSecciones X on S.cod_institucion = X.cod_institucion" _
           & "  and S.cod_departamento = X.cod_departamento and S.cod_seccion = X.cod_seccion" _
           & " inner join promotores P on S.id_promotor = P.id_promotor" _
           & " inner join afi_profesiones R on S.cod_profesion = R.cod_profesion" _
           & " inner join afi_sectores Q on S.cod_sector = Q.cod_sector" _
           & " inner join AFI_ESTADOS_PERSONA Est on S.EstadoActual = Est.Cod_Estado" _
           & " left join Provincias Prov on S.Provincia = Prov.Provincia" _
           & " left join sif_oficinas O on S.cod_oficina = O.cod_oficina " _
           & " left join Cantones Cant on S.Provincia = Cant.Provincia and S.Canton = Cant.Canton" _
           & " left join Distritos Dist on S.Provincia = Dist.Provincia and convert(int,S.Canton) = convert(int,Dist.Canton) and S.distrito = Dist.distrito" _
           & " left join AFI_TIPOS_IDS Tid on S.tipo_id = Tid.tipo_id" _
           & " left join AFI_SOCIEDADES_TIPOS Soc on S.cod_sociedad = Soc.cod_sociedad" _
           & " left join AFI_ACTIVIDADES_ECO Act on S.cod_actividad = Act.cod_actividad" _
           & " where cedula='" & Trim(vCedula) & "'"
Else
   'Modo de ASECCSS
    strSQL = "Select S.*,UT as 'Cod_Seccion',UP as 'Cod_Departamento',C.descripcion as 'CentroDesc',O.descripcion as Oficina" _
           & ",Est.Descripcion as 'EstadoPersonaDesc',Est.Cod_Estado + ' - ' + Est.Descripcion as 'EstadoPersona'" _
           & ",I.descripcion as DescInst,D.descripcion as DescDept,X.ut_descripcion as DescSec,P.nombre as Promotor,R.descripcion as ProfesionX" _
           & ",Q.descripcion as Sector,dbo.fxAFIAnioServicio(cedula,'" & Format(vFechaActual, "yyyy/mm/dd") & "') as AnioServicio" _
           & ",rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
           & ",Tid.Descripcion as TipoIdDesc,Tid.Tipo_Personeria,Soc.cod_sociedad + ' - ' + rtrim(Soc.descripcion) as 'SociedadDesc'" _
           & ",Act.cod_actividad + ' - ' + rtrim(Act.descripcion) as 'ActividadDesc'" _
           & " From socios S inner join Instituciones I on S.cod_institucion = I.cod_institucion" _
           & " left join uprogramatica D on S.UP = D.codigo" _
           & " left join utrabajo X on S.ut = X.ut_codigo" _
           & " left join uprogramatica C on S.CT = C.codigo" _
           & " inner join promotores P on S.id_promotor = P.id_promotor" _
           & " inner join afi_profesiones R on S.cod_profesion = R.cod_profesion" _
           & " inner join afi_sectores Q on S.cod_sector = Q.cod_sector" _
           & " inner join AFI_ESTADOS_PERSONA Est on S.EstadoActual = Est.Cod_Estado" _
           & " left join sif_oficinas O on S.cod_oficina = O.cod_oficina " _
           & " left join Provincias Prov on S.Provincia = Prov.Provincia" _
           & " left join Cantones Cant on S.Provincia = Cant.Provincia and S.Canton = Cant.Canton" _
           & " left join Distritos Dist on S.Provincia = Dist.Provincia and S.Canton = Dist.Canton and S.distrito = Dist.distrito" _
           & " left join AFI_TIPOS_IDS Tid on S.tipo_id = Tid.tipo_id" _
           & " left join AFI_SOCIEDADES_TIPOS Soc on S.cod_sociedad = Soc.cod_sociedad" _
           & " left join AFI_ACTIVIDADES_ECO Act on S.cod_actividad = Act.cod_actividad" _
           & " where cedula='" & Trim(vCedula) & "'"
   
End If

Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
   
   vCedula = Trim(rs!Cedula)
   
   txtCedula.Text = Trim(rs!Cedula)
   lblNombre.Caption = Trim(rs!Nombre)

   txtBoleta = rs!id_Boleta_AF & ""
     
   txtEstadoPersona = rs!EstadoPersona
  
   dtpFechaIngreso = rs!FechaIngreso
   dtpNacimiento = rs!fecha_nac
   txtSexo = IIf(rs!sexo = "M", "Masculino", "Femenino")
     
   txtEstado = fxEstadoCivil(rs!estadoCivil)
     
     
   txtProvincia = rs!ProvDesc
   txtCanton = rs!CantonDesc
   txtDistrito = rs!DistDesc
     
      
   txtDireccion = Trim(rs!Direccion) & ""
   txtEmail = Trim(rs!AF_Email) & ""
   txtApartado = Trim(rs!apto) & ""
   
   If IIf(IsNull(rs!estadoLaboral), 1, rs!estadoLaboral) = 1 Then
     optNombramiento.Item(0).Value = True
   Else
     optNombramiento.Item(1).Value = True
   End If
   
   dtpNombramiento.Value = IIf(IsNull(rs!nombramiento_fecha), dtpFechaIngreso.Value, rs!nombramiento_fecha)
   lswNombramiento.ListItems.Clear
   
   txtAniosSerivicio.Text = Trim(rs!AnioServicio)
   
   
   
   txtCodPromotor.Text = rs!id_promotor
   txtNombrePromotor.Text = Trim(rs!promotor)
   
   txtNotificaciones.Text = Trim(rs!Notificaciones & "")
   
   txtInstitucion.Text = Trim(rs!DescInst)
   txtProfesion.Text = Trim(rs!profesionX)
   txtSector.Text = Trim(rs!sector)
   
   txtDeptCodigo = rs!cod_departamento & ""
   txtDeptDesc = Trim(rs!descDept & "")
   
   txtSecCodigo = rs!cod_seccion & ""
   txtSecDesc = Trim(rs!DescSec & "")
   
   lblCentroTrabajo.Visible = False
   txtCTCodigo.Visible = False
   txtCTDesc.Visible = False
   
   If GLOBALES.SysASEVersion Then
        lblCentroTrabajo.Visible = True
        txtCTCodigo.Visible = True
        txtCTDesc.Visible = True
        
        txtCTCodigo.Text = rs!CT & ""
        txtCTDesc.Visible = rs!CentroDesc & ""
   End If
   
   txtHijos.Text = IIf(IsNull(rs!hijos), 0, rs!hijos)
   txtNumeroPagos = IIf(IsNull(rs!af_npagos), 0, rs!af_npagos)

End If
rs.Close

If vConsecutivo = Empty Then Call sbCargaConsecutivo

SSTab1.TabEnabled(1) = True
SSTab1.TabEnabled(2) = True
SSTab1.TabEnabled(3) = True





End Sub

Private Sub sbCargarGridSeguimiento(pCedula As String, pBoleta As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
vGridSeguimiento.MaxCols = 4
vGridSeguimiento.MaxRows = 0
    

Me.MousePointer = vbHourglass

strSQL = "select T.DESCRIPCION, OT.NOTAS, OT.REGISTRO_FECHA, OT.REGISTRO_USUARIO" _
    & " from SIF_CONTROL_TAGS OT inner join SIF_TAGS T on OT.TAG_CODIGO = T.TAG_CODIGO" _
    & " where OT.codigo = '" & pCedula & "' and OT.cod_Modulo = 'AFI' and OT.Documento = '" & pBoleta & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    vGridSeguimiento.MaxRows = vGridSeguimiento.MaxRows + 1
    vGridSeguimiento.Row = vGridSeguimiento.MaxRows
  
    vGridSeguimiento.Col = 1
    vGridSeguimiento.Text = rs!Descripcion
    vGridSeguimiento.TextTip = TextTipFixed
    vGridSeguimiento.TextTipDelay = 1000
    vGridSeguimiento.CellNote = "Usuario: " & rs!registro_usuario & "[" & rs!Registro_Fecha & "]"
            
    vGridSeguimiento.Col = 2
    vGridSeguimiento.Value = IIf(IsNull(rs!notas), "", rs!notas)
    
    vGridSeguimiento.Col = 3
    vGridSeguimiento.Value = IIf(IsNull(rs!Registro_Fecha), "", rs!Registro_Fecha)
    
    vGridSeguimiento.Col = 4
    vGridSeguimiento.Value = IIf(IsNull(rs!registro_usuario), "", rs!registro_usuario)
    
    vGridSeguimiento.RowHeight(vGridSeguimiento.Row) = vGridSeguimiento.MaxTextRowHeight(vGridSeguimiento.Row)
    rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCargarCombosEtiquetas()
Dim strSQL As String

On Error GoTo vError

    
    strSQL = "SELECT CT.TAG_CODIGO + ' - ' +  rtrim(CT.DESCRIPCION) as 'ItmX'" _
            & " FROM SIF_TAGS CT INNER JOIN SIF_TAGS_GRUPOS CTG ON CT.TAG_CODIGO = CTG.TAG_CODIGO" _
            & " INNER JOIN SIF_GRPUSERS CGU ON CTG.COD_GRUPO = CGU.COD_GRUPO" _
            & " WHERE CT.ACTIVO = 1 AND CGU.USUARIO = '" & glogon.Usuario _
            & "' and  CT.TAG_CODIGO in(select TAG_CODIGO from SIF_TAGS_MODULOS where cod_modulo = 'AFI')" _
            & " order by CT.TAG_CODIGO"
    vPaso = True
    Call sbLlenaCbo(cboEtiquetas, strSQL, False, False)
    vPaso = False
    Call cboEtiquetas_Click
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCargarListaErrores()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If txtCedula = Empty Then
    Exit Sub
End If

With lswErrores
 .ListItems.Clear
 
 strSQL = "select E.ID_ERROR,E.DESCRIPCION,ER.ID_ERROR as asignado, ISNULL(ER.APLICADO,'N') AS APLICADO, E.MENSAJE, ER.LINEA_ERR" _
        & " from sif_Omisiones E left join SIF_OMISIONESG ER on E.ID_ERROR = ER.ID_ERROR" _
        & " and ER.cedula = '" & txtCedula.Text & "' and ER.Modulo = 'AFI' and ER.Codigo = '" & txtCedula.Text _
        & "' and ER.Documento = '" & txtId.Text & "'" _
        & " where E.ACTIVO = '1'  and E.ID_ERROR in(select ID_ERROR from SIF_OMISIONES_MODULOS where cod_modulo = 'AFI') " _
        & " order by E.ID_ERROR"
        
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!ID_ERROR)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
         itmX.Tag = rs!LINEA_ERR
      End If
      itmX.SubItems(2) = rs!APLICADO
      itmX.SubItems(3) = rs!Mensaje
  rs.MoveNext
 Loop
 rs.Close
End With
End Sub

Private Sub sbCargarObservacion()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError
    
    strSQL = "select ISNULL(MENSAJE,'') from SIF_TAGS_AVISOS where TAG_CODIGO = '" & SIFGlobal.fxCodText(cboEtiquetas.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        txtObservacion = rs.Fields(0) & vbNewLine
    Else
        txtObservacion = Empty
    End If
    
    For i = 1 To lswErrores.ListItems.Count
        If lswErrores.ListItems(i).Checked = True Then
            If lswErrores.ListItems(i).SubItems(2) = "N" Then
                If txtObservacion = Empty Then
                    txtObservacion.Text = "-" & lswErrores.ListItems(i).SubItems(3)
                Else
                    txtObservacion.Text = txtObservacion.Text & vbNewLine & "-" & lswErrores.ListItems(i).SubItems(3)
                End If
            End If
        End If
    Next
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbAplicarErrores()
'' Procedimiento para colocar los errores ingresados en aplicados
Dim Linea As String, strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    strSQL = "update SIF_OMISIONESG SET APLICADO = 'S' WHERE cedula = '" & txtCedula.Text _
           & "' AND MODULO = 'AFI' AND CODIGO = '" & txtCedula.Text & "' AND DOCUMENTO = '" & txtId.Text & "'"
    Call ConectionExecute(strSQL)

    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaConsecutivo()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select max(consec) as consecutivo from afi_ingresos where cedula = '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  vConsecutivo = rs!consecutivo
  txtId.Text = rs!consecutivo
End If
rs.Close
End Sub
