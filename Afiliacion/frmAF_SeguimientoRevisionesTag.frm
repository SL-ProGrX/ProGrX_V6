VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_SeguimientoRevisionesTag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisión de Afliaciones"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraControles 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   0
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
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Afiliaciones"
         TabPicture(0)   =   "frmAF_SeguimientoRevisionesTag.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "vGridAfiliaciones"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detalle"
         TabPicture(1)   =   "frmAF_SeguimientoRevisionesTag.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSTab2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Seguimiento"
         TabPicture(2)   =   "frmAF_SeguimientoRevisionesTag.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "vGridSeguimiento"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Revisión"
         TabPicture(3)   =   "frmAF_SeguimientoRevisionesTag.frx":0054
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "Label8(1)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Label2(0)"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Label27"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "lswErrores"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "tlbAplicar"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "txtObservacion"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "cboEtiquetas"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).ControlCount=   7
         Begin TabDlg.SSTab SSTab2 
            Height          =   5655
            Left            =   -74400
            TabIndex        =   16
            Top             =   480
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   9975
            _Version        =   393216
            TabOrientation  =   1
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            ForeColor       =   16711680
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
            TabPicture(0)   =   "frmAF_SeguimientoRevisionesTag.frx":0070
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Line1"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label13"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label12(1)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label12(0)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label10(2)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label10(1)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label15(0)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label14"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Label1(0)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Label6"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Label5"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "dtpNacimiento"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "dtpFechaIngreso"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "ssTabSubX"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "txtNombrePromotor"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "txtCodPromotor"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "txtBoleta"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "txtHijos"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "txtProfesion"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "txtSector"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "txtEstado"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "txtSexo"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "txtEstadoPersona"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).ControlCount=   23
            TabCaption(1)   =   "Otros"
            TabPicture(1)   =   "frmAF_SeguimientoRevisionesTag.frx":68D2
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lblEtiqueta"
            Tab(1).Control(1)=   "lswOtros"
            Tab(1).Control(2)=   "optX(2)"
            Tab(1).Control(3)=   "optX(1)"
            Tab(1).Control(4)=   "optX(0)"
            Tab(1).ControlCount=   5
            Begin VB.TextBox txtEstadoPersona 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1080
               MaxLength       =   15
               TabIndex        =   87
               ToolTipText     =   "Campo para la Cédula de Identidad"
               Top             =   690
               Width           =   1575
            End
            Begin VB.TextBox txtSexo 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   3720
               MaxLength       =   15
               TabIndex        =   86
               ToolTipText     =   "Campo para la Cédula de Identidad"
               Top             =   330
               Width           =   1695
            End
            Begin VB.TextBox txtEstado 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1080
               MaxLength       =   15
               TabIndex        =   85
               ToolTipText     =   "Campo para la Cédula de Identidad"
               Top             =   330
               Width           =   1575
            End
            Begin VB.TextBox txtSector 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   6480
               MaxLength       =   15
               TabIndex        =   84
               ToolTipText     =   "Campo para la Cédula de Identidad"
               Top             =   1530
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
               Left            =   1080
               MaxLength       =   45
               TabIndex        =   83
               ToolTipText     =   "Campo para la Cédula de Identidad"
               Top             =   1200
               Width           =   4335
            End
            Begin VB.OptionButton optX 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Telefonos"
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
               Height          =   1035
               Index           =   0
               Left            =   -74520
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmAF_SeguimientoRevisionesTag.frx":D134
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   4530
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optX 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Beneficarios"
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
               Height          =   1035
               Index           =   1
               Left            =   -73320
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmAF_SeguimientoRevisionesTag.frx":13986
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   4530
               Width           =   1215
            End
            Begin VB.OptionButton optX 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Cuentas"
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
               Height          =   1035
               Index           =   2
               Left            =   -72120
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmAF_SeguimientoRevisionesTag.frx":1A1D8
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   4530
               Width           =   1215
            End
            Begin VB.TextBox txtHijos 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   6480
               MaxLength       =   15
               TabIndex        =   54
               ToolTipText     =   "Campo para la Cédula de Identidad"
               Top             =   1170
               Width           =   1815
            End
            Begin VB.TextBox txtBoleta 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   3720
               MaxLength       =   15
               TabIndex        =   53
               Text            =   "1"
               Top             =   690
               Width           =   1695
            End
            Begin VB.TextBox txtCodPromotor 
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   1530
               Width           =   975
            End
            Begin VB.TextBox txtNombrePromotor 
               Height          =   285
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   1530
               Width           =   3375
            End
            Begin TabDlg.SSTab ssTabSubX 
               Height          =   3375
               Left            =   240
               TabIndex        =   19
               Top             =   1920
               Width           =   8175
               _ExtentX        =   14420
               _ExtentY        =   5953
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
               TabPicture(0)   =   "frmAF_SeguimientoRevisionesTag.frx":20A2A
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
               TabPicture(1)   =   "frmAF_SeguimientoRevisionesTag.frx":2728C
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
               TabPicture(2)   =   "frmAF_SeguimientoRevisionesTag.frx":2DAEE
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
                  Height          =   285
                  Left            =   -73560
                  TabIndex        =   82
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
                  TabIndex        =   38
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
                  TabIndex        =   37
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
                  TabIndex        =   36
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
                  TabIndex        =   35
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
                  TabIndex        =   30
                  Top             =   300
                  Width           =   7095
                  Begin VB.TextBox txtDistrito 
                     Height          =   285
                     Left            =   840
                     TabIndex        =   81
                     Top             =   1080
                     Width           =   2055
                  End
                  Begin VB.TextBox txtCanton 
                     Height          =   285
                     Left            =   840
                     TabIndex        =   80
                     Top             =   720
                     Width           =   2055
                  End
                  Begin VB.TextBox txtProvincia 
                     Height          =   285
                     Left            =   840
                     TabIndex        =   79
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
                     TabIndex        =   31
                     ToolTipText     =   "Dirección exacta Aqui"
                     Top             =   360
                     Width           =   3975
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
                     TabIndex        =   34
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
                     TabIndex        =   33
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
                     TabIndex        =   32
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
                  Left            =   1320
                  MaxLength       =   15
                  TabIndex        =   29
                  ToolTipText     =   "Campo para la Cédula de Identidad"
                  Top             =   2340
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
                  Left            =   1320
                  MaxLength       =   45
                  TabIndex        =   28
                  ToolTipText     =   "Campo para la Cédula de Identidad"
                  Top             =   1980
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
                  Left            =   1320
                  MaxLength       =   255
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   27
                  Top             =   2700
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
                  TabIndex        =   25
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
                  TabIndex        =   24
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
                  TabIndex        =   23
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
                  TabIndex        =   22
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
                  TabIndex        =   21
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
                  TabIndex        =   20
                  ToolTipText     =   "Presione F4 para Consultar"
                  Top             =   2100
                  Width           =   615
               End
               Begin MSComctlLib.ListView lswNombramiento 
                  Height          =   2175
                  Left            =   -72000
                  TabIndex        =   26
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
                  TabIndex        =   39
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
                  Format          =   83492867
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
                  TabIndex        =   78
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
                  TabIndex        =   52
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
                  TabIndex        =   51
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
                  TabIndex        =   50
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
                  Left            =   120
                  TabIndex        =   49
                  Top             =   2340
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
                  Left            =   120
                  TabIndex        =   48
                  Top             =   1980
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
                  Left            =   120
                  TabIndex        =   47
                  Top             =   2700
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
                  TabIndex        =   46
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
                  TabIndex        =   45
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
                  TabIndex        =   44
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
                  TabIndex        =   43
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
                  TabIndex        =   42
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
                  TabIndex        =   41
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
                  TabIndex        =   40
                  Top             =   1980
                  Width           =   1095
               End
            End
            Begin MSComCtl2.DTPicker dtpFechaIngreso 
               Height          =   315
               Left            =   6720
               TabIndex        =   58
               ToolTipText     =   "Fecha de Ingreso al sistema"
               Top             =   690
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
               Format          =   83492867
               CurrentDate     =   38899
               MaxDate         =   55153
               MinDate         =   14611
            End
            Begin MSComCtl2.DTPicker dtpNacimiento 
               Height          =   315
               Left            =   6720
               TabIndex        =   59
               ToolTipText     =   "Fecha de Ingreso al sistema"
               Top             =   330
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
               Format          =   83492867
               CurrentDate     =   36059
            End
            Begin MSComctlLib.ListView lswOtros 
               Height          =   3975
               Left            =   -74520
               TabIndex        =   60
               Top             =   450
               Width           =   7815
               _ExtentX        =   13785
               _ExtentY        =   7011
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
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
               NumItems        =   0
            End
            Begin VB.Label Label16 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Aplicar Deducción Doble en Planillas"
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
               Index           =   1
               Left            =   -74760
               TabIndex        =   77
               Top             =   2520
               Width           =   3255
            End
            Begin VB.Label Label16 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cambio de Primer Deducción de Aportes"
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
               Index           =   0
               Left            =   -74760
               TabIndex        =   76
               Top             =   480
               Width           =   3495
            End
            Begin VB.Label Label16 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Desactivar Cálculo de Aporte Patronal"
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
               Index           =   2
               Left            =   -74760
               TabIndex        =   75
               Top             =   1440
               Width           =   3255
            End
            Begin VB.Label Label16 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Bloqueo a Créditos y Notas de Advertencia"
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
               Height          =   300
               Index           =   3
               Left            =   -74760
               TabIndex        =   74
               Top             =   3960
               Width           =   3735
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Caption         =   "Apellido 1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   -74880
               TabIndex        =   73
               Top             =   360
               Width           =   2295
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
               Left            =   120
               TabIndex        =   72
               Top             =   690
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
               Left            =   5760
               TabIndex        =   71
               Top             =   690
               Width           =   735
            End
            Begin VB.Label lblEtiqueta 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   -74520
               TabIndex        =   70
               Top             =   210
               Width           =   7815
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
               Left            =   5760
               TabIndex        =   69
               Top             =   330
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
               Left            =   3000
               TabIndex        =   68
               Top             =   330
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
               Left            =   120
               TabIndex        =   67
               Top             =   330
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
               Left            =   120
               TabIndex        =   66
               Top             =   1170
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
               Left            =   120
               TabIndex        =   65
               Top             =   1530
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
               Left            =   5640
               TabIndex        =   64
               ToolTipText     =   "Número de Dependientes"
               Top             =   1170
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
               Left            =   5640
               TabIndex        =   63
               Top             =   1530
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
               Left            =   3000
               TabIndex        =   62
               Top             =   690
               Width           =   735
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               X1              =   8280
               X2              =   0
               Y1              =   1050
               Y2              =   1050
            End
            Begin VB.Label lblOficina 
               Alignment       =   1  'Right Justify
               Caption         =   ".."
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4080
               TabIndex        =   61
               Top             =   -330
               Width           =   3360
            End
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
            ItemData        =   "frmAF_SeguimientoRevisionesTag.frx":34350
            Left            =   1680
            List            =   "frmAF_SeguimientoRevisionesTag.frx":34352
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   5295
         End
         Begin VB.TextBox txtObservacion 
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
            Left            =   1680
            MaxLength       =   995
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   4080
            Width           =   9135
         End
         Begin FPSpreadADO.fpSpread vGridAfiliaciones 
            Height          =   5775
            Left            =   -74880
            TabIndex        =   8
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   7
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "frmAF_SeguimientoRevisionesTag.frx":34354
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   487
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "frmAF_SeguimientoRevisionesTag.frx":34D11
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin MSComctlLib.Toolbar tlbAplicar 
            Height          =   570
            Left            =   1560
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
            Left            =   1680
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
            Left            =   240
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
            Left            =   240
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
            Left            =   240
            TabIndex        =   12
            Top             =   4080
            Width           =   1455
         End
      End
   End
   Begin VB.Frame fraOperacion 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   11295
      Begin VB.TextBox txtCedula 
         Alignment       =   2  'Center
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
         Left            =   960
         TabIndex        =   1
         Top             =   120
         Width           =   2415
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
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         ToolTipText     =   "Nombre"
         Top             =   120
         Width           =   5895
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
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":352BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":3BB21
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":42383
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":48BE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":4F447
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":55CA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   11280
      Top             =   240
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
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5C50B
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5C629
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5C74F
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5C879
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5C98B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5CAA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5CBA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5CCDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5CDEF
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":5CF13
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":63775
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":69FD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_SeguimientoRevisionesTag.frx":6A0F1
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
      Left            =   2400
      TabIndex        =   15
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmAF_SeguimientoRevisionesTag.frx":6A20F
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   1920
      Picture         =   "frmAF_SeguimientoRevisionesTag.frx":6A404
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmAF_SeguimientoRevisionesTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCedula As String, mConsecutivo As String

Private Sub cboEtiquetas_Click()
Call sbCargarObservacion
End Sub

Private Sub Form_Load()
Call sbCargarListaAfliaiciones

End Sub


Private Sub lswErrores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
    If Item.SubItems(2) = "S" Then
        Item.Checked = True
        If MsgBox("El error ya fué aplicado desea agregar únicamente la nota", vbOKCancel) = vbOK Then
            If txtObservacion = Empty Then
              txtObservacion.Text = Item.SubItems(3)
            Else
              txtObservacion.Text = txtObservacion.Text & Item.SubItems(3)
            End If
        End If
        Exit Sub
    End If
    
    If Item.Checked Then
    
      strSQL = "insert SIF_OMISIONESG (cedula,ID_ERROR) values('" & mCedula _
             & "'," & Item.Text & ")"
             
      If txtObservacion = Empty Then
        txtObservacion.Text = "-" & Item.SubItems(1)
      Else
        txtObservacion.Text = txtObservacion.Text & Item.SubItems(1)
      End If
      
    Else
      strSQL = "delete SIF_OMISIONESG where cedula = '" & mCedula & "'" _
              & " and ID_ERROR = " & Item.Text
             
     Call sbCargarObservacion
    End If
    glogon.Conection.Execute strSQL
    
    Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub OptX_Click(Index As Integer)
Dim rs As New ADODB.Recordset
Dim itmX As ListItem

lswOtros.ListItems.Clear
lswOtros.ColumnHeaders.Clear


lblEtiqueta.Caption = UCase(optX.Item(Index).Caption)

Select Case Index
  Case 0 'Telefonos
       
    lswOtros.ColumnHeaders.Add 1, , "Numero", 1500
    lswOtros.ColumnHeaders.Add 2, , "Tipo", 1500
    lswOtros.ColumnHeaders.Add 3, , "Extension", 1500
    lswOtros.ColumnHeaders.Add 4, , "Contacto", 2500
    
    lblEtiqueta = "TELEFONOS"
    
    strSQL = "Select * From Telefonos where " _
           & "Cedula='" & Trim(txtCedula) & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    Do While Not rs.EOF
       Set itmX = lswOtros.ListItems.Add(, , Trim(rs!Numero))
           itmX.SubItems(1) = fxTipoTelefono(rs!Tipo)
           itmX.SubItems(2) = Trim(rs!Ext) & ""
           itmX.SubItems(3) = Trim(rs!contacto) & ""
       rs.MoveNext
    Loop
    rs.Close
  
  Case 1 'Beneficiario
  
    lswOtros.ColumnHeaders.Add 1, , "Cedula", 1500
    lswOtros.ColumnHeaders.Add 2, , "Nombre", 3500
    lswOtros.ColumnHeaders.Add 3, , "Porcentaje", 1100
    
    strSQL = "Select CedulaBn,Nombre,Porcentaje From Beneficiarios where " _
           & " Cedula = '" & Trim(txtCedula) & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    Do While Not rs.EOF
       Set itmX = lswOtros.ListItems.Add(, , rs!cedulaBn)
           itmX.SubItems(1) = Trim(rs!Nombre)
           itmX.SubItems(2) = Trim(rs!Porcentaje)
       rs.MoveNext
    Loop
    rs.Close
  
  
  Case 2 'Cuentas de Ahorros
    lswOtros.ColumnHeaders.Add 1, , "Cuenta", 1500
    lswOtros.ColumnHeaders.Add 2, , "Banco", 3500
    lswOtros.ColumnHeaders.Add 3, , "Tipo", 1100

           
    strSQL = "select A.cuenta,B.descripcion,A.tipo" _
           & " from cuentas_ahorros A inner join Tes_Bancos B on A.id_banco = B.id_Banco" _
           & " where A.cedula = '" & Trim(txtCedula) & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    Do While Not rs.EOF
       Set itmX = lswOtros.ListItems.Add(, , rs!Cuenta)
           itmX.SubItems(1) = Trim(rs!Descripcion)
       Select Case rs!Tipo
         Case 0
            itmX.SubItems(2) = "Cuenta Corriente"
         Case 1
            itmX.SubItems(2) = "Cuenta de Ahorros"
         Case 2
            itmX.SubItems(2) = "Tarjeta de Crédito"
       End Select
       rs.MoveNext
    Loop
    rs.Close
  
  

End Select


End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
   Case 2
     Call sbCargarGridSeguimiento(txtCedula)
   Case 3
     Call sbCargarListaErrores
     Call sbCargarCombosEtiquetas
End Select
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
If SSTab2.Tab = 1 Then Call OptX_Click(0)

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
    Call sbGuardar(txtCedula, mConsecutivo)
   
''   Se pasa al sp al insertar el tag
'    If SIFGlobal.fxSIFCodText(cboEtiquetas) = mTagRevision Then
'      Call sbCambiaCreditoRevisado
'    End If
    
    Call sbAplicarErrores
    Call sbCargarListaAfliaiciones
    txtCedula.SetFocus
    SSTab1.Tab = 0
    
    
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
    Resume
End Sub

Private Sub txtCedula_GotFocus()
SSTab1.Tab = 0
mCedula = Empty
mConsecutivo = Empty
Call sbLimpiaControles
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


Private Sub vGridAfiliaciones_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    vGridAfiliaciones.Col = 2
    vGridAfiliaciones.Row = Row

'    Call sbLimpiarDatosCreditos(True)
    mCedula = vGridAfiliaciones.Text
    vGridAfiliaciones.Col = 3
    lblNombre = vGridAfiliaciones.Text
    vGridAfiliaciones.Col = 7
    mConsecutivo = vGridAfiliaciones.Text
    If Len(Trim(mCedula)) > 0 Then
        Call sbCurrentRecord(mCedula)
    End If

End Sub


Private Sub sbCargarListaAfliaiciones(Optional ByVal strCedula As String = Empty)
' Carga Lista de afiliaciones
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo error
Me.MousePointer = vbHourglass

If Trim(strCedula) = Empty Then

    strSQL = "select Top 3000 A.cedula, S.nombre,A.usuario,A.cod_remesa,R.usuario,A.consec from afi_ingresos A" _
            & " inner join Socios S on A.Cedula = S.cedula left join AFI_REMESAS_ING R on A.cod_remesa = R.cod_remesa" _
            & " Where a.ANALISTA_REVISION  Is Null"

Call sbCargaGridCheckIni(vGridAfiliaciones, 6, strSQL)
Me.MousePointer = vbDefault
End If
Exit Sub
    
error:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical

End Sub



Private Sub sbCurrentRecord(vCedula As String)
Dim rs As New ADODB.Recordset, rsTemp As New ADODB.Recordset
Dim vApellido1 As String, vApellido2 As String
Dim vNombre1 As String, vNombre2 As String
Dim i As Integer, vEspacio As Integer

On Error Resume Next

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

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
   vEditar = True
   
   
'   Call sbToolBar(Me.tlb, "activo")
'   Call RefrescaTags(Me)
'   Call sbLockControles("L")
'   Call sbLimpiaDatos 'Inicializa Datos
   
   vCedula = Trim(rs!Cedula)
   txtCedula.Text = Trim(rs!Cedula)
   
   If Not IsNull(rs!TipoIdDesc) Then
       vPaso = True
           cboTipoId.Text = Trim(rs!TipoIdDesc)
       vPaso = False
   End If
   
     
   If rs!Tipo_Personeria = "J" Then
      fraTipo.Visible = True
      txtNombreComercial.Text = Trim(rs!Nombre)
      txtRazonSocial.Text = Trim(rs!Razon_Social & "")
'      If Not IsNull(rs!Cod_actividad) Then
'         Call sbCboAsignaDato(cboActividad, rs!ActividadDesc)
'      End If
'      If Not IsNull(rs!cod_sociedad) Then
'         Call sbCboAsignaDato(cboSociedad, rs!SociedadDesc)
'      End If
      
   Else
      fraTipo.Visible = False
        
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
   
        'txtApellido1 = vApellido1
        'txtApellido2 = vApellido2
        lblNombre = vApellido1 & " " & " " & vApellido2 & " " & vNombre1 & " " & vNombre2
   
   End If
    
     
   txtCedAlternativa = Trim(rs!cedular & "")
   
     
   txtBoleta = rs!id_Boleta_AF & ""
     

   'Carga Información del Estado de la Persona y sus posibles Acciones
   tlbIngreso.Buttons.Item(1).Enabled = False 'Reingreso
   tlbIngreso.Buttons.Item(2).Enabled = False 'Activacion
   
   strSQL = "select COD_MOVIMIENTO from AFI_ESTADOS_CAMBIO" _
          & " where COD_ESTADO = '" & rs!EstadoActual & "' and COD_MOVIMIENTO IN('REI','ACT')"
   rsTemp.Open strSQL, glogon.Conection, adOpenStatic
   Do While Not rsTemp.EOF
    If rsTemp!COD_MOVIMIENTO = "REI" Then tlbIngreso.Buttons.Item(1).Enabled = True
    If rsTemp!COD_MOVIMIENTO = "ACT" Then tlbIngreso.Buttons.Item(1).Enabled = True
    rsTemp.MoveNext
   Loop
   rsTemp.Close
   
   'Si el Estado es de Ingreso, Puede usarse con otros estados de ingreso, caso contrario limpiar la lista
   strSQL = "select count(*) as Ingreso from AFI_ESTADOS_CAMBIO" _
          & " where COD_ESTADO = '" & rs!EstadoActual & "' and COD_MOVIMIENTO IN('ING')"
   rsTemp.Open strSQL, glogon.Conection, adOpenStatic
   If rsTemp!Ingreso = 0 Then
       cboEstadoPersona.Clear
   End If
   rsTemp.Close
   
   txtEstadoPersona = rs!EstadoPersona
   
   
   If IsNull(rs!Prideduc) Then
      txtPriDeduc.Text = GLOBALES.glngFechaCR
   Else
      txtPriDeduc.Text = rs!Prideduc
   End If
     
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
   
   txtConyugeCedula.Text = Trim(rs!conyuge_cedula & "")
   txtConyugeNombre.Text = Trim(rs!conyuge_nombre & "")
   txtConyugeTelCelular.Text = Trim(rs!conyuge_TelCell & "")
   txtConyugeTelTrabajo.Text = Trim(rs!conyuge_TelTra & "")
   txtConyugeTelTrabajoExt.Text = Trim(rs!conyuge_TelTraExt & "")
   
   txtAlbaceaCedula.Text = Trim(rs!albacea_Cedula & "")
   txtAlbaceaNombre.Text = Trim(rs!albacea_nombre & "")
   
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
   lblOficina.Caption = IIf(IsNull(rs!Oficina), "Sin Descripción", rs!Oficina)
   
   If lblOficina.Caption = "Sin Descripción" Then
    lblOficina.Tag = 0
   Else
    lblOficina.Tag = rs!COD_OFICINA
   End If
   
   chkBienes.Value = IIf(IsNull(rs!ind_propiedades), 0, rs!ind_propiedades)
   
   chkBloqueo.Value = IIf(IsNull(rs!bloqueo), 0, rs!bloqueo)
   chkDesactivaAporte.Value = IIf(IsNull(rs!ind_sinAporte), 0, rs!ind_sinAporte)
   chkDobleDeduccion.Value = IIf(IsNull(rs!IND_DOBLE_DEDUCCION), 0, rs!IND_DOBLE_DEDUCCION)

   
   txtNotasAdv.Text = rs!notas & ""
   
   'txtCedula.SetFocus
   
   StatusBarX.Panels.Item(1) = rs!reg_user & ""
   StatusBarX.Panels.Item(2) = rs!reg_fecha & ""
   StatusBarX.Panels.Item(3) = rs!ActualizaUser & ""
   StatusBarX.Panels.Item(4) = rs!ActualizaFecha & ""
   
   vCambios.vFecNac = dtpNacimiento.Value ' carga fecha Nac. para verificacion
   vCambios.vEstado = cboEstado.Text ' carga Estado civil para verificacion
   vCambios.vPromotor = txtNombrePromotor.Text ' carga promotor para verificacion

rs.Close

If mConsecutivo = Empty Then Call sbCargaConsecutivo
SSTab1.TabEnabled(1) = True
SSTab1.TabEnabled(2) = True
SSTab1.TabEnabled(3) = True

End If



End Sub

Private Sub sbLimpiaControles()

For Each vControl In Me
  If TypeOf vControl Is TextBox Then
     vControl.Text = ""
  End If
Next
End Sub



Private Sub sbCargarGridSeguimiento(ByVal vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    If mCedula = Empty Then Exit Sub

    Me.MousePointer = vbHourglass

    strSQL = "select T.DESCRIPCION, OT.NOTAS, OT.REGISTRO_FECHA, OT.REGISTRO_USUARIO from SIF_CONTROL_TAGS OT" _
           & " inner join SIF_TAGS T on OT.TAG_CODIGO = T.TAG_CODIGO where OT.cedula = '" & vCedula & "'"
            
    vGridSeguimiento.MaxCols = 4
    vGridSeguimiento.MaxRows = 0


rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
    vGridSeguimiento.MaxRows = vGridSeguimiento.MaxRows + 1
    vGridSeguimiento.Row = vGridSeguimiento.MaxRows
  
    vGridSeguimiento.Col = 1
    vGridSeguimiento.Text = rs!Descripcion
    vGridSeguimiento.TextTip = TextTipFixed
    vGridSeguimiento.TextTipDelay = 1000
    vGridSeguimiento.CellNote = "Usuario: " & rs!Registro_Usuario & "[" & rs!Registro_Fecha & "]"
            
    vGridSeguimiento.Col = 2
    vGridSeguimiento.Value = IIf(IsNull(rs!notas), "", rs!notas)
    
    vGridSeguimiento.Col = 3
    vGridSeguimiento.Value = IIf(IsNull(rs!Registro_Fecha), "", rs!Registro_Fecha)
    
    vGridSeguimiento.Col = 4
    vGridSeguimiento.Value = IIf(IsNull(rs!Registro_Usuario), "", rs!Registro_Usuario)
    
    vGridSeguimiento.RowHeight(vGridSeguimiento.Row) = vGridSeguimiento.MaxTextRowHeight(vGridSeguimiento.Row)
    rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault
Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
End Sub


Private Sub sbCargarCombosEtiquetas()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

    cboEtiquetas.Clear
    cboEtiquetas.AddItem " "
    
    strSQL = "SELECT CT.TAG_CODIGO as llave,CT.DESCRIPCION as describe FROM SIF_TAGS CT INNER JOIN SIF_TAGS_GRUPOS CTG ON CT.TAG_CODIGO = CTG.TAG_CODIGO" _
       & " INNER JOIN SIF_GRPUSERS CGU ON CTG.COD_GRUPO = CGU.COD_GRUPO" _
       & " WHERE CT.ACTIVO = 1 AND CGU.USUARIO = '" & glogon.Usuario _
       & "' order by CT.TAG_CODIGO"
    rs.Open strSQL, glogon.Conection, adOpenStatic

    Do While Not rs.EOF
      cboEtiquetas.AddItem Trim(rs!llave) & " - " & Trim(rs!describe)
      rs.MoveNext
    Loop
    rs.Close
    
    cboEtiquetas.Text = " "
    
    Exit Sub
vError:
 MsgBox Err.Description, vbCritical

End Sub



Private Sub sbCargarListaErrores()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If mCedula = Empty Then
    Exit Sub
End If

With lswErrores
 .ListItems.Clear
  
 strSQL = "select E.ID_ERROR,E.DESCRIPCION,ER.ID_ERROR as asignado, ISNULL(ER.APLICADO,'N') AS APLICADO, E.MENSAJE" _
        & " from sif_Omisiones E left join SIF_OMISIONESG ER on E.ID_ERROR = ER.ID_ERROR" _
        & " and ER.cedula = '" & mCedula & "'" _
        & " where E.ACTIVO = '1'" _
        & " order by E.ID_ERROR"
        
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!ID_ERROR)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
      itmX.SubItems(2) = rs!APLICADO
      itmX.SubItems(3) = rs!Mensaje
  rs.MoveNext
 Loop
 rs.Close
End With
End Sub




Private Sub sbGuardar(vCedula As String, vConsecutivo As String)

'Call sbRegistraTags(mCedula, SIFGlobal.fxSIFCodText(cboEtiquetas.Text), txtObservacion, mConsecutivo)

End Sub



Private Sub sbCargarObservacion()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError
    
    strSQL = "select ISNULL(MENSAJE,'') from SIF_TAGS_AVISOS where TAG_CODIGO = '" & SIFGlobal.fxSIFCodText(cboEtiquetas.Text) & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
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
  MsgBox Err.Description, vbCritical
End Sub


Private Sub sbAplicarErrores()
'' Procedimiento para colocar los errores ingresados en aplicados
Dim Linea As String, strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    If mCedula = Empty Then
        Exit Sub
    End If
    
    strSQL = "update SIF_OMISIONESG SET APLICADO = 'S' WHERE cedula = '" & mCedula & "'"
    glogon.Conection.Execute strSQL

    Exit Sub
vError:
    MsgBox Err.Description, vbCritical

End Sub


Private Sub sbCargaConsecutivo()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select max(consec) as consecutivo from afi_reingresos where cedula = '" & txtCedula & "' and ANALISTA_REVISION  is null "
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF Then
  mConsecutivo = rs!consecutivo
End If
rs.Close
End Sub
