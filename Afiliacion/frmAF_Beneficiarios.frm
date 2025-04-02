VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_Beneficiarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Beneficiarios y Otros Contactos"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7455
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   9855
      _Version        =   1441793
      _ExtentX        =   17383
      _ExtentY        =   13150
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
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Listado"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   44
      Item(1).Control(0)=   "txtObservacion"
      Item(1).Control(1)=   "txtApellido1"
      Item(1).Control(2)=   "txtApellido2"
      Item(1).Control(3)=   "txtNombre"
      Item(1).Control(4)=   "txtDireccion"
      Item(1).Control(5)=   "txtApartadoPostal"
      Item(1).Control(6)=   "txtTelefono2"
      Item(1).Control(7)=   "txtTelefono1"
      Item(1).Control(8)=   "txtPorcentaje"
      Item(1).Control(9)=   "txtCedula"
      Item(1).Control(10)=   "dtpFechaNacimiento"
      Item(1).Control(11)=   "Label7(1)"
      Item(1).Control(12)=   "Label16"
      Item(1).Control(13)=   "Label15(0)"
      Item(1).Control(14)=   "Label14"
      Item(1).Control(15)=   "Label8"
      Item(1).Control(16)=   "Label4(0)"
      Item(1).Control(17)=   "Lbl5"
      Item(1).Control(18)=   "Lbl4"
      Item(1).Control(19)=   "Lbl3(0)"
      Item(1).Control(20)=   "Lbl1"
      Item(1).Control(21)=   "cboParentesco"
      Item(1).Control(22)=   "cboRelacion"
      Item(1).Control(23)=   "chkSeguros"
      Item(1).Control(24)=   "Label15(1)"
      Item(1).Control(25)=   "txtEmail"
      Item(1).Control(26)=   "txtCodigo"
      Item(1).Control(27)=   "txtAlbaceaNombre"
      Item(1).Control(28)=   "txtAlbaceaCedula"
      Item(1).Control(29)=   "txtAlbaceaTelTrabajo"
      Item(1).Control(30)=   "txtAlbaceaTelTrabajoExt"
      Item(1).Control(31)=   "Label(4)"
      Item(1).Control(32)=   "Label(3)"
      Item(1).Control(33)=   "Label(2)"
      Item(1).Control(34)=   "Label(0)"
      Item(1).Control(35)=   "Label(1)"
      Item(1).Control(36)=   "txtAlbaceaTelCelular"
      Item(1).Control(37)=   "ShortcutCaption3(1)"
      Item(1).Control(38)=   "ShortcutCaption3(0)"
      Item(1).Control(39)=   "ShortcutCaption3(2)"
      Item(1).Control(40)=   "ShortcutCaption3(3)"
      Item(1).Control(41)=   "chkAlbacea"
      Item(1).Control(42)=   "cboTipoId"
      Item(1).Control(43)=   "Label1"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   6855
         Left            =   -70000
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   9855
         _Version        =   1441793
         _ExtentX        =   17383
         _ExtentY        =   12091
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkAlbacea 
         Height          =   210
         Left            =   9000
         TabIndex        =   44
         Top             =   5355
         Width           =   210
         _Version        =   1441793
         _ExtentX        =   362
         _ExtentY        =   370
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkSeguros 
         Height          =   255
         Left            =   7200
         TabIndex        =   17
         Top             =   840
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica Seguros?"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboParentesco 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   2040
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboRelacion 
         Height          =   315
         Left            =   3960
         TabIndex        =   14
         Top             =   840
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   330
         Left            =   1680
         TabIndex        =   16
         Top             =   840
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaNacimiento 
         Height          =   315
         Left            =   4920
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   556
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   795
         Left            =   1680
         TabIndex        =   19
         Top             =   3120
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
         _ExtentY        =   1402
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtObservacion 
         Height          =   915
         Left            =   1680
         TabIndex        =   20
         Top             =   4320
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
         _ExtentY        =   1614
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   330
         Left            =   1680
         TabIndex        =   21
         Top             =   3960
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApartadoPostal 
         Height          =   330
         Left            =   7320
         TabIndex        =   22
         Top             =   2640
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono1 
         Height          =   330
         Left            =   1680
         TabIndex        =   23
         Top             =   2640
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono2 
         Height          =   330
         Left            =   4320
         TabIndex        =   24
         Top             =   2640
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   330
         Left            =   6240
         TabIndex        =   25
         Top             =   1560
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido2 
         Height          =   330
         Left            =   3960
         TabIndex        =   26
         Top             =   1560
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido1 
         Height          =   330
         Left            =   1680
         TabIndex        =   27
         Top             =   1560
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
         Height          =   315
         Left            =   8520
         TabIndex        =   28
         Top             =   2040
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   315
         Left            =   8520
         TabIndex        =   29
         Top             =   7080
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAlbaceaNombre 
         Height          =   315
         Left            =   3480
         TabIndex        =   30
         Top             =   6000
         Width           =   5895
         _Version        =   1441793
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAlbaceaCedula 
         Height          =   315
         Left            =   1680
         TabIndex        =   31
         Top             =   6000
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAlbaceaTelTrabajo 
         Height          =   315
         Left            =   5520
         TabIndex        =   32
         Top             =   6600
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAlbaceaTelTrabajoExt 
         Height          =   315
         Left            =   7560
         TabIndex        =   33
         Top             =   6600
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAlbaceaTelCelular 
         Height          =   315
         Left            =   3480
         TabIndex        =   39
         Top             =   6600
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipoId 
         Height          =   330
         Left            =   1680
         TabIndex        =   45
         Top             =   480
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Id"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   480
         Width           =   1215
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   360
         Index           =   3
         Left            =   6240
         TabIndex        =   43
         Top             =   1200
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
         _ExtentY        =   635
         _StockProps     =   14
         Caption         =   "Nombre"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   360
         Index           =   2
         Left            =   3960
         TabIndex        =   42
         Top             =   1200
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   635
         _StockProps     =   14
         Caption         =   "Apellido 2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   360
         Index           =   0
         Left            =   1680
         TabIndex        =   41
         Top             =   1200
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   635
         _StockProps     =   14
         Caption         =   "Apellido 1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   40
         Top             =   5280
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Datos del Albacea Especifico en Caso de que el Beneficiario sea menor de edad"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3480
         TabIndex        =   38
         Top             =   5760
         Width           =   1245
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   37
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. Móvil"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   36
         Top             =   6360
         Width           =   1275
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Extensión"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7560
         TabIndex        =   35
         Top             =   6390
         Width           =   1035
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. Trabajo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5520
         TabIndex        =   34
         Top             =   6390
         Width           =   1275
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Linea Id:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   18
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Label Lbl1 
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Lbl3 
         Caption         =   "Parentesco"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Lbl4 
         Caption         =   "Fec. Nac."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Lbl5 
         Caption         =   "Porcentaje de Beneficio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Teléfono 2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   8
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Apto. Postal"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   6
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Teléfono 1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   2640
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
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
End
Attribute VB_Name = "frmAF_Beneficiarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vCedulaPrincipal As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vPaso As Boolean
Dim vAlbaceaGeneral As Boolean, vFecha As Date


Private Sub sbCargaLsw()

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "exec spAFI_PERSONA_BENEFICIARIOS_Consulta '" & vCedulaPrincipal & "',0"
Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Linea_Id)
    itmX.SubItems(1) = rs!cedula_Beneficiario
    itmX.SubItems(2) = rs!Nombre
    itmX.SubItems(3) = Format(rs!fecha_nac, "dd/mm/yyyy")
    itmX.SubItems(4) = rs!parentesco
    itmX.SubItems(5) = rs!Porcentaje & " %"
    itmX.SubItems(6) = rs!Relacion_Desc
   rs.MoveNext
Loop
     
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()

vModulo = 1

On Error GoTo vError

vCedulaPrincipal = GLOBALES.gCedulaActual
 
vAlbaceaGeneral = False
If GLOBALES.gTag = "S" Then
    vAlbaceaGeneral = True
End If
 
 vEdita = True
 Call sbToolBarIconos(tlbPrincipal)
 Call sbToolBar(tlbPrincipal, "nuevo")
 
 cboRelacion.AddItem "Beneficiario"
 cboRelacion.ItemData(cboRelacion.ListCount - 1) = "B"
 cboRelacion.AddItem "Contacto"
 cboRelacion.ItemData(cboRelacion.ListCount - 1) = "C"
 
 cboRelacion.Text = "Beneficiario"

 strSQL = "select rtrim(cod_Parentesco) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
        & " from sys_Parentescos where activo = 1"
 Call sbCbo_Llena_New(cboParentesco, strSQL, False, True)
 
'Carga Tipos de Identificacion
vPaso = True
strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " Where TIPO_PERSONERIA = 'F' order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
vPaso = False
 
 
 
 vFecha = fxFechaServidor
  
 With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Identificación", 1400
    .Add , , "Nombre", 3400
    .Add , , "Fec.Nac.", 1400, vbCenter
    .Add , , "Parentesco", 1400
    .Add , , "Porcentaje", 1200, vbCenter
    .Add , , "Relación", 1200, vbCenter
    
 End With
 
 Call sbLimpiaPantalla(0)

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla(Optional Index As Integer = 1)


tcMain.Item(Index).Selected = True

Select Case Index
    Case 0 'Lista
        Call sbCargaLsw
    Case 1 'Caso
        vCodigo = 0
        txtCodigo.Text = ""
        txtCedula.Text = ""
        
'        strSQL = "select isnull(count(*),0) + 1 as Consec from AFI_PERSONA_BENEFICIARIOS where cedula = '" & vCedulaPrincipal & "'"
'        Call OpenRecordSet(rs, strSQL)
'          txtCedula = Trim(vCedulaPrincipal) & "-" & Format(rs!consec, "00")
'        rs.Close
        
        txtApellido1 = ""
        txtApellido2 = ""
        txtNombre = ""
        
        dtpFechaNacimiento.MaxDate = fxFechaServidor
        dtpFechaNacimiento.Value = dtpFechaNacimiento.MaxDate
        
        txtPorcentaje.Text = "0"
        
        txtObservacion.Text = ""
        txtDireccion.Text = ""
        txtApartadoPostal.Text = ""
        txtEmail.Text = ""
        txtTelefono1.Text = ""
        txtTelefono2.Text = ""
        
        chkAlbacea.Value = xtpUnchecked
        txtAlbaceaCedula.Text = ""
        txtAlbaceaNombre.Text = ""
        txtAlbaceaTelCelular.Text = ""
        txtAlbaceaTelTrabajo.Text = ""
        txtAlbaceaTelTrabajoExt.Text = ""
        
End Select

End Sub


Private Sub lsw_DblClick()
If lsw.ListItems.Count > 0 Then
   Call sbConsulta(lsw.SelectedItem.Text)
End If
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 0 Then

    Call sbCargaLsw

End If

End Sub



Private Sub sbBoleta()

On Error GoTo vError

Me.MousePointer = vbHourglass


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Personas"
 
 .Connect = glogon.ConectRPT
      
 .ReportFileName = SIFGlobal.fxPathReportes("Personas_Boleta_Beneficiarios.rpt")
    
 .StoredProcParam(0) = vCedulaPrincipal
 .StoredProcParam(1) = 1
    
 .SubreportToChange = "sbBeneficiarios"
 .StoredProcParam(0) = vCedulaPrincipal
    
    
 .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:

    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      tcMain.Item(1).Selected = True
      txtCedula.SetFocus
      Call sbToolBar(tlbPrincipal, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      tcMain.Item(1).Selected = True
      txtCedula.SetFocus
      Call sbToolBar(tlbPrincipal, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlbPrincipal, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlbPrincipal, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select consec,cedulaBN,nombre from beneficiarios"
       gBusquedas.Filtro = " and cedula = '" & vCedulaPrincipal & "'"
       frmBusquedas.Show vbModal
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       Call sbConsulta(txtCodigo)
        
    
    Case "REPORTES"
      Call sbBoleta
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vApellido1 As String, vApellido2 As String, vNombre1 As String, vNombre2 As String
Dim vEspacio As Integer, i As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_PERSONA_BENEFICIARIOS_Consulta '" & vCedulaPrincipal & "'," & lngCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlbPrincipal, "activo")
  
  tcMain.Item(1).Selected = True
      
  vEdita = True
  vCodigo = rs!Linea_Id
  txtCodigo = rs!Linea_Id
  
  txtCedula.Text = Trim(rs!cedula_Beneficiario)
  
  Call sbCboAsignaDato(cboTipoId, rs!Tipo_Id_Desc, False, rs!Tipo_Id_R)
 
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
   txtApellido1 = vApellido1
   txtApellido2 = vApellido2
   txtNombre = vNombre1 & " " & vNombre2
   
   txtPorcentaje.Text = Format(rs!Porcentaje, "##0.00")
   dtpFechaNacimiento.Value = rs!fecha_nac
   txtObservacion = Trim(rs!Notas & "")
       
    txtDireccion = Trim(rs!direccion & "")
    txtApartadoPostal = Trim(rs!apto_postal & "")
    txtEmail = Trim(rs!Email & "")
    
    txtTelefono1 = Trim(rs!telefono1 & "")
    txtTelefono2 = Trim(rs!telefono2 & "")
    
    
    Call sbCboAsignaDato(cboRelacion, rs!Relacion_Desc, True, rs!Tipo_Relacion)
    Call sbCboAsignaDato(cboParentesco, rs!parentesco, True, rs!Cod_Parentesco)
    
    chkSeguros.Value = rs!Aplica_Seguros
    
    chkAlbacea.Value = rs!Albacea_Check
    txtAlbaceaCedula.Text = rs!albacea_Cedula & ""
    txtAlbaceaNombre.Text = rs!albacea_nombre & ""
    txtAlbaceaTelCelular.Text = rs!ALBACEA_MOVIL & ""
    txtAlbaceaTelTrabajo.Text = rs!ALBACEA_TELTRA & ""
    txtAlbaceaTelTrabajoExt.Text = rs!ALBACEA_TELTRA_EXT & ""
    
    
Else
  
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, pCedulaLargo As Integer

vMensaje = ""
fxValida = True

If cboParentesco.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No se ha seleccionado ningún parentesco..."

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Beneficiario no es válido ..."
If txtApellido1 = "" Then vMensaje = vMensaje & vbCrLf & " - txtApellido 1 del Beneficiario no es válido ..."
If txtApellido2 = "" Then vMensaje = vMensaje & vbCrLf & " - txtApellido 2 del Beneficiario no es válido ..."


'Actualiza el Parametro de Validacion y Luego lo Aplica
strSQL = "select LARGO_MINIMO from AFI_TIPOS_IDS Where TIPO_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    pCedulaLargo = rs!Largo_Minimo
End If
rs.Close

If Len(Trim(txtCedula)) <> pCedulaLargo Then vMensaje = vMensaje & " - Número de Identidad no es válido, se espera que sea de: " & pCedulaLargo _
            & " caracteres, verifique...!" & vbCrLf


'Verificar que el porcentaje no supere el 100 %

If Not vAlbaceaGeneral And dtpFechaNacimiento.Value > DateAdd("yyyy", -18, vFecha) And chkAlbacea.Value = xtpUnchecked Then
    vMensaje = vMensaje & vbCrLf & " - No se ha indicado un Albacea General, y este beneficiario es menor de edad (Indique uno específico)"
End If

If chkAlbacea.Value = xtpChecked And Len(txtAlbaceaCedula.Text) <= 5 Then
    vMensaje = vMensaje & vbCrLf & " - El Albacea indicado no es válido!"
End If

If Len(txtEmail.Text) > 0 Then

    If Not fxEmail_Valida(txtEmail.Text) Then
        vMensaje = vMensaje & vbCrLf & " - Correo Electrónico no es válido..."
    End If
End If

If Not IsNumeric(txtPorcentaje) Then
   vMensaje = vMensaje & vbCrLf & " - El porcentaje no es válido ..."
Else
    If CCur(txtPorcentaje.Text) = 0 Then vMensaje = vMensaje & vbCrLf & " - El porcentaje no es válido ..."
    
    strSQL = "select isnull(sum(porcentaje),0) as Porcentaje from AFI_PERSONA_BENEFICIARIOS" _
           & " where cedula = '" & vCedulaPrincipal & "' and Linea_ID <> " & vCodigo
    Call OpenRecordSet(rs, strSQL)
    If CCur(txtPorcentaje) + rs!Porcentaje > 100 Then
       vMensaje = vMensaje & vbCrLf & " - El porcentaje sobre pasa el total del 100% del total de los beneficiarios ..."
    End If
    rs.Close
End If


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


strSQL = "exec spAFI_PERSONA_BENEFICIARIOS_Registra '" & vCedulaPrincipal & "'," & vCodigo & ",'" & txtCedula.Text _
                & "','" & UCase(Trim(txtApellido1)) & " " & UCase(Trim(txtApellido2)) & " " & UCase(Trim(txtNombre)) _
                & "','" & Format(dtpFechaNacimiento.Value, "yyyy/mm/dd") & "','" & cboRelacion.ItemData(cboRelacion.ListIndex) _
                & "','" & cboParentesco.ItemData(cboParentesco.ListIndex) & "'," & CCur(txtPorcentaje.Text) _
                & "," & chkSeguros.Value & ",'" & txtObservacion.Text & "','" & txtDireccion.Text & "','" & txtApartadoPostal.Text _
                & "','" & txtTelefono1.Text & "','" & txtTelefono2.Text & "','" & txtEmail.Text & "','A','" & glogon.Usuario _
                & "', " & chkAlbacea.Value & ", '" & Trim(txtAlbaceaCedula.Text) & "', '" & Trim(txtAlbaceaNombre.Text) _
                & "','" & Trim(txtAlbaceaTelCelular.Text) & " ', '" & Trim(txtAlbaceaTelTrabajo.Text) & "', '" & Trim(txtAlbaceaTelTrabajoExt.Text) _
                & "', " & cboTipoId.ItemData(cboTipoId.ListIndex)
                
Call OpenRecordSet(rs, strSQL)
If Not glogon.error Then
  vCodigo = rs!LineaId
  txtCodigo.Text = rs!LineaId
  rs.Close

Else
  Exit Sub
End If

If vEdita Then
  
  Call Bitacora("Modifica", "Beneficiario Id.: " & txtCedula.Text & " Consec.: " & vCodigo)
  
  If vParametros.BitacoraEspecial Then
       Call sbgAFIBitacora("13", "Modifica Beneficiario Id.: " & txtCedula.Text & " Consec.: " & vCodigo, Trim(GLOBALES.gCedulaActual))
  End If
  
Else
    
  Call Bitacora("Registra", "Beneficiario Id.: " & txtCedula.Text & " Consec.: " & vCodigo)
 
  If vParametros.BitacoraEspecial Then
       Call sbgAFIBitacora("12", "Registra Beneficiario Id.: " & txtCedula.Text & " Consec.: " & vCodigo, Trim(GLOBALES.gCedulaActual))
  End If
  
End If


   
MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbToolBar(tlbPrincipal, "activo")
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
 
 strSQL = "exec spAFI_PERSONA_BENEFICIARIOS_Registra '" & vCedulaPrincipal & "'," & vCodigo & ",'" & txtCedula.Text _
                & "','" & UCase(Trim(txtApellido1)) & " " & UCase(Trim(txtApellido2)) & " " & UCase(Trim(txtNombre)) _
                & "','" & Format(dtpFechaNacimiento.Value, "yyyy/mm/dd") & "','" & cboRelacion.ItemData(cboRelacion.ListIndex) _
                & "','" & cboParentesco.ItemData(cboParentesco.ListIndex) & "'," & CCur(txtPorcentaje.Text) _
                & "," & chkSeguros.Value & ",'" & txtObservacion.Text & "','" & txtDireccion.Text & "','" & txtApartadoPostal.Text _
                & "','" & txtTelefono1.Text & "','" & txtTelefono2.Text & "','" & txtEmail.Text & "','E','" & glogon.Usuario & "'"
  
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Beneficiario Id.: " & txtCedula.Text & " Consec.: " & vCodigo)
  
  If vParametros.BitacoraEspecial Then
       Call sbgAFIBitacora("14", "Elimina Beneficiario Id.: " & txtCedula.Text & " Consec.: " & vCodigo, Trim(GLOBALES.gCedulaActual))
  End If
  
  Call sbLimpiaPantalla(0)
  Call sbToolBar(tlbPrincipal, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtAlbaceaCedula_LostFocus()
On Error GoTo vError
        
'Consulta Padron
Call gBase_Padron(txtAlbaceaCedula.Text, "General", rs, "CRI")

If rs.RecordCount > 0 Then
   txtAlbaceaNombre.Text = Trim(rs!Apellido_1) & " " & Trim(rs!Apellido_2) & " " & Trim(rs!Nombre)
End If

vError:

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
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAlbaceaTelCelular.SetFocus
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



Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApellido1.SetFocus
End Sub

Private Sub txtApellido1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApellido2.SetFocus
End Sub

Private Sub txtApellido2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub


Private Sub txtCedula_LostFocus()
On Error GoTo vError
        
If txtCodigo.Text = "" Then
    Call gBase_Padron(txtCedula.Text, "General", rs, "CRI")
    
    If rs.RecordCount > 0 Then
       txtApellido1.Text = Trim(rs!Apellido_1)
       txtApellido2.Text = Trim(rs!Apellido_2)
       txtNombre.Text = Trim(rs!Nombre)
        
       dtpFechaNacimiento.Value = rs!fecha_nacimiento
       
    End If
End If

vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboParentesco.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  
  gBusquedas.Col1Name = "Registro Id"
  gBusquedas.Col2Name = "Identificacion"
  gBusquedas.Col3Name = "Nombre Completo"
  
  gBusquedas.Consulta = "select Linea_Id,Cedula_Beneficiario,Nombre from AFI_PERSONA_BENEFICIARIOS"
  gBusquedas.Filtro = " and Cedula = '" & vCedulaPrincipal & "'"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub cboParentesco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFechaNacimiento.SetFocus
End Sub

Private Sub dtpFechaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPorcentaje.SetFocus
End Sub

Private Sub txtPorcentaje_GotFocus()
On Error GoTo vError
 txtPorcentaje = CCur(txtPorcentaje)
vError:
End Sub

Private Sub txtPorcentaje_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono1.SetFocus
End Sub

Private Sub txtPorcentaje_LostFocus()
On Error GoTo vError
 txtPorcentaje = Format(CCur(txtPorcentaje), "##0.00")
vError:
End Sub

Private Sub txtTelefono1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono2.SetFocus
End Sub

Private Sub txtTelefono2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApartadoPostal.SetFocus
End Sub

Private Sub txtApartadoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservacion.SetFocus
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedula.SetFocus
End Sub
