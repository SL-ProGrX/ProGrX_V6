VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmUS_Usuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Usuarios"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10410
   HelpContextID   =   1005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5175
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   10455
      _Version        =   1441793
      _ExtentX        =   18441
      _ExtentY        =   9128
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
      ItemCount       =   3
      Item(0).Caption =   "Usuario"
      Item(0).ControlCount=   20
      Item(0).Control(0)=   "chkContabiliza"
      Item(0).Control(1)=   "txtIdentificacion"
      Item(0).Control(2)=   "txtNombre"
      Item(0).Control(3)=   "txtTelCelular"
      Item(0).Control(4)=   "txtTelTrabajo"
      Item(0).Control(5)=   "txtEmail"
      Item(0).Control(6)=   "txtNotas"
      Item(0).Control(7)=   "txtEstado"
      Item(0).Control(8)=   "dtpIngreso"
      Item(0).Control(9)=   "dtpModificacion"
      Item(0).Control(10)=   "Label2"
      Item(0).Control(11)=   "Label1(4)"
      Item(0).Control(12)=   "Label1(3)"
      Item(0).Control(13)=   "Label1(7)"
      Item(0).Control(14)=   "Label1(6)"
      Item(0).Control(15)=   "Label5"
      Item(0).Control(16)=   "Label4(0)"
      Item(0).Control(17)=   "Label3"
      Item(0).Control(18)=   "btnActivar"
      Item(0).Control(19)=   "gb2FA"
      Item(1).Caption =   "Miembro de"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "Label8(0)"
      Item(1).Control(1)=   "lswMiembros"
      Item(1).Control(2)=   "lswRoles"
      Item(1).Control(3)=   "lblRol"
      Item(2).Caption =   "Bitácora"
      Item(2).ControlCount=   7
      Item(2).Control(0)=   "lswBitacora"
      Item(2).Control(1)=   "cboTransac"
      Item(2).Control(2)=   "txtLineas"
      Item(2).Control(3)=   "dtpBInicio"
      Item(2).Control(4)=   "dtpBCorte"
      Item(2).Control(5)=   "Label8(2)"
      Item(2).Control(6)=   "Label8(3)"
      Begin XtremeSuiteControls.ListView lswBitacora 
         Height          =   4215
         Left            =   -70000
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1441793
         _ExtentX        =   18230
         _ExtentY        =   7435
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ListView lswRoles 
         Height          =   4215
         Left            =   -64720
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   7435
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ListView lswMiembros 
         Height          =   4215
         Left            =   -69880
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   7435
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.GroupBox gb2FA 
         Height          =   975
         Left            =   840
         TabIndex        =   45
         Top             =   3960
         Width           =   7935
         _Version        =   1441793
         _ExtentX        =   13996
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "2FA Doble Factor de Autenticación"
         ForeColor       =   4210752
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
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chk2FA 
            Height          =   255
            Left            =   1440
            TabIndex        =   46
            Top             =   360
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Activar 2FA"
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
            Appearance      =   17
         End
         Begin XtremeSuiteControls.ComboBox cbo2FA 
            Height          =   330
            Left            =   6240
            TabIndex        =   47
            Top             =   360
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
         Begin XtremeSuiteControls.Label Label6 
            Height          =   255
            Left            =   3840
            TabIndex        =   48
            Top             =   360
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Canal de Validación"
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
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.CheckBox chkContabiliza 
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   480
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Contabiliza en Cobranza"
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
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Left            =   2280
         TabIndex        =   9
         Top             =   960
         Width           =   6375
         _Version        =   1441793
         _ExtentX        =   11239
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
      Begin XtremeSuiteControls.FlatEdit txtTelCelular 
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTelTrabajo 
         Height          =   315
         Left            =   6600
         TabIndex        =   11
         Top             =   1920
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   315
         Left            =   2280
         TabIndex        =   12
         Top             =   2280
         Width           =   6375
         _Version        =   1441793
         _ExtentX        =   11239
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1155
         Left            =   2280
         TabIndex        =   13
         Top             =   2640
         Width           =   6375
         _Version        =   1441793
         _ExtentX        =   11245
         _ExtentY        =   2037
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
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   315
         Left            =   6600
         TabIndex        =   14
         Top             =   1440
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpIngreso 
         Height          =   315
         Left            =   2280
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.DateTimePicker dtpModificacion 
         Height          =   315
         Left            =   5040
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.DateTimePicker dtpBInicio 
         Height          =   315
         Left            =   -65200
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.DateTimePicker dtpBCorte 
         Height          =   315
         Left            =   -63880
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.ComboBox cboTransac 
         Height          =   330
         Left            =   -70000
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
         _Version        =   1441793
         _ExtentX        =   6588
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
      Begin XtremeSuiteControls.FlatEdit txtLineas 
         Height          =   315
         Left            =   -60520
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Text            =   "150"
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnActivar 
         Height          =   315
         Left            =   8760
         TabIndex        =   36
         Top             =   1440
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   3
         Left            =   -61840
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Resultados:"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   2
         Left            =   -66040
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fechas"
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
      End
      Begin XtremeSuiteControls.Label lblRol 
         Height          =   255
         Left            =   -64720
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Roles:"
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
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Index           =   0
         Left            =   -69880
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   3495
         _Version        =   1441793
         _ExtentX        =   6165
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Clientes Relacionados:"
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
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Ingreso"
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
         Left            =   720
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Ultimo Movimiento"
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
         Left            =   3960
         TabIndex        =   22
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
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
         Index           =   6
         Left            =   720
         TabIndex        =   21
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   720
         TabIndex        =   20
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Tel. Celular"
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
         Left            =   720
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Left            =   4680
         TabIndex        =   18
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Left            =   720
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
   End
   Begin XtremeSuiteControls.PushButton btnExiste 
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Top             =   600
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Existe?"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   5520
      Top             =   480
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtUserName 
      Height          =   435
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.FlatEdit txtUserID 
      Height          =   435
      Left            =   8520
      TabIndex        =   5
      Top             =   600
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnCuenta 
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   37
      Top             =   0
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cuenta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_UsuariosX.frx":0000
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnCuenta 
      Height          =   375
      Index           =   1
      Left            =   9000
      TabIndex        =   38
      Top             =   0
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Restablece"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_UsuariosX.frx":0719
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   1080
      TabIndex        =   39
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_UsuariosX.frx":0E32
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   2160
      TabIndex        =   40
      ToolTipText     =   "Editar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_UsuariosX.frx":1464
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   2520
      TabIndex        =   41
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_UsuariosX.frx":1A5F
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   3120
      TabIndex        =   42
      ToolTipText     =   "Guardar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_UsuariosX.frx":2003
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   3480
      TabIndex        =   43
      ToolTipText     =   "Deshacer"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_UsuariosX.frx":2734
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   3960
      TabIndex        =   44
      ToolTipText     =   "Reporte"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmUS_UsuariosX.frx":2E34
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   375
      Index           =   1
      Left            =   7320
      TabIndex        =   4
      Top             =   600
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "User Id:"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Usuario:"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUS_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vBusca As Integer, vScroll As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Dim v2FA_Indica As Integer, v2FA_Metodo As String


Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub btnBarra_Click(Index As Integer)
Dim strSQL As String


Select Case Index
    Case 0 'NUEVO
      vEdita = False
      Call sbLimpiaPantalla
      txtUserName.SetFocus
      Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
        vEdita = True
        txtNombre.SetFocus
        Call sbBarra_Accion("Editar")
     
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
     
      Call sbBarra_Accion("Editar")
      If txtUserName.Text = "" Then
        Call sbLimpiaPantalla
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      Else
        Call sbConsulta(txtUserName.Text)
      End If
    
    Case 5 'REPORTES
           frmUS_ReporteUsuarios.Show vbModal
   
End Select

'            gBusquedas.Columna = "nombre"
'            gBusquedas.Orden = "nombre"
'            gBusquedas.Consulta = "select Usuario,Nombre from US_usuarios"
'            gBusquedas.Filtro = ""
'            frmBusquedas.Show vbModal
'            txtUserName = gBusquedas.Resultado
'            txtNombre = gBusquedas.Resultado2
'            Call sbConsulta(txtUserName)
'            txtUserName.SetFocus

End Sub




Private Sub btnActivar_Click()
If txtUserID.Text = "" Then Exit Sub

gEntidad.UserID = txtUserID.Text
gEntidad.Usuario = txtUserName.Text

frmUS_Activacion.Show vbModal

Call sbConsulta(txtUserName)
End Sub

Private Sub btnCuenta_Click(Index As Integer)
If txtUserID = "" Then Exit Sub

gEntidad.Tipo = "U"
gEntidad.UserID = txtUserID
gEntidad.Usuario = Trim(txtUserName)

Select Case Index
  Case 0  'cuenta"
    frmUS_Cuentas.Show vbModal
  Case 1 'Restablece
    frmUS_CuentaRestablece.Show vbModal
End Select

End Sub

Private Sub btnExiste_Click()
Dim vMensaje As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select count(*) as 'Existe' from US_USUARIOS" _
       & " where Usuario = '" & txtUserName.Text & "'"
       
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    vMensaje = "Usuario: Libre!"
Else
    vMensaje = "Usuario: Ocupado!"
End If
rs.Close

Me.MousePointer = vbDefault

MsgBox vMensaje, vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboTransac_Click()
If vPaso Then Exit Sub
Call sbBitacoraLlena
End Sub


Private Sub dtpBCorte_Change()
If vPaso Then Exit Sub
Call sbBitacoraLlena
End Sub

Private Sub dtpBInicio_Change()
If vPaso Then Exit Sub
Call sbBitacoraLlena
End Sub

Private Sub dtpIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpModificacion.SetFocus
End Sub

Private Sub dtpModificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelCelular.SetFocus
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 Usuario from US_USUARIOS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where Usuario > '" & txtUserName & "'"
    Else
       strSQL = strSQL & " where Usuario < '" & txtUserName & "'"
    End If
    
    
    If Not gAdminAccess.Rol_AdminView Then
        strSQL = strSQL & " AND isnull(key_admin,0) = 0"
    End If
    
    If Not gAdminAccess.Rol_DirGlobal Then
        strSQL = strSQL & " AND usuario in(select usuario from PGX_CLIENTES_USERS" _
            & " Where cod_Empresa = " & gPortal.Empresa_Id & ")"
    End If
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " order by Usuario asc"
    Else
       strSQL = strSQL & " order by Usuario desc"
    End If
    
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtUserName = rs!Usuario
      Call sbConsulta(txtUserName)
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbBitacoraLlena()

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "select Top " & txtLineas.Text & "  L.*,T.descripcion " _
       & " from us_transac_log L inner join us_transacciones T on L.cod_transac = T.cod_transac" _
       & " where L.Usuario = '" & txtUserName.Text & "' and L.mov_fecha between '" _
       & Format(dtpBInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpBCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
       
If cboTransac.Text <> "TODOS" Then
 strSQL = strSQL & " and T.cod_transac = '" & cboTransac.ItemData(cboTransac.ListIndex) & "'"
End If
       
strSQL = strSQL & " order by L.mov_fecha desc"
Call OpenRecordSet(rs, strSQL)


With lswBitacora.ListItems
 .Clear
 Do While Not rs.EOF
  Set itmX = .Add(, , rs!Mov_fecha)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = rs!NOTAS
      itmX.SubItems(3) = rs!Mov_User
      itmX.SubItems(4) = rs!App_Name
      itmX.SubItems(5) = rs!App_Version
      itmX.SubItems(6) = rs!Equipo
  rs.MoveNext
 Loop
 
End With

rs.Close
Me.MousePointer = vbDefault

End Sub


Private Sub Form_Load()

vModulo = 13


 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True


With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "", lswMiembros.Width - 100
End With

With lswRoles.ColumnHeaders
    .Clear
    .Add , , "", lswRoles.Width - 100
End With
        
With lswBitacora.ColumnHeaders
    .Clear
    .Add , , "Fecha", 2100
    .Add , , "Tipo Registro", 2100
    .Add , , "Notas", 4100
    .Add , , "Usuario", 2100, vbCenter
    .Add , , "App.Name", 2100, vbCenter
    .Add , , "App.Version", 2100, vbCenter
    .Add , , "Equipo", 2100
End With
        
cbo2FA.AddItem "MAIL"
cbo2FA.AddItem "SMS"
cbo2FA.AddItem "APP"
cbo2FA.Text = "MAIL"
        

 vEdita = True
 
 Call sbBarra_Accion("Activo")
 
 Call sbLimpiaPantalla
 
 'Llena Combos
 vPaso = True
 
 'Transacciones
 strSQL = "select cod_transac as 'IdX' , rtrim(descripcion) as ItmX from us_transacciones"
 Call sbCbo_Llena_New(cboTransac, strSQL, True, True)
 
 dtpBCorte.Value = fxFechaServidor
 dtpBInicio.Value = DateAdd("m", -1, dtpBCorte.Value)
 
 vPaso = False
 
 
With gAdminAccess
    If Not .Rol_ResetKeys Then
        btnCuenta(0).Enabled = False
        btnCuenta(1).Enabled = False
    End If
    
    If Not .Admin_Portal Then
       tcMain.Item(1).Visible = False
    End If

End With
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
End Sub


Private Sub sbLimpiaPantalla()

vBusca = 1

txtUserID.Text = ""
txtUserName.Text = ""
txtNombre.Text = ""
txtNotas.Text = ""

dtpIngreso.Value = fxFechaServidor
dtpModificacion.Value = dtpIngreso.Value

txtIdentificacion.Text = ""
txtEMail.Text = ""
txtTelCelular.Text = ""
txtTelTrabajo.Text = ""

chkContabiliza.Value = vbChecked

txtEstado.Text = "Inactivo"

If v2FA_Indica = 1 Then
    chk2FA.Enabled = False
    chk2FA.Value = xtpChecked
Else
    chk2FA.Enabled = True
    chk2FA.Value = xtpUnchecked
End If

cbo2FA.Text = v2FA_Metodo



tcMain.Item(0).Selected = True
tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False

End Sub



Private Sub lswBitacora_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswBitacora.SortKey = ColumnHeader.Index - 1
  If lswBitacora.SortOrder = 0 Then lswBitacora.SortOrder = 1 Else lswBitacora.SortOrder = 0
  lswBitacora.Sorted = True
End Sub


Private Sub lswMiembros_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswMiembros.SortKey = ColumnHeader.Index - 1
  If lswMiembros.SortOrder = 0 Then lswMiembros.SortOrder = 1 Else lswMiembros.SortOrder = 0
  lswMiembros.Sorted = True
End Sub

Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
  'Incluirlo
    strSQL = "exec spPGX_Usuario_Cliente_Asigna " & Item.Tag & ",'" & Trim(txtUserName.Text) & "','" & glogon.Usuario & "','I',''"
    Call ConectionExecute(strSQL)

    Call sbSEGCuentaLog("08", "Membresía al Rol: (" & Item.Tag & ") " & Trim(Item.Text), glogon.Usuario, txtUserName.Text)

    'Sincroniza Core
    Call spCore_Usuario_Sincroniza(Item.Tag, Trim(txtUserName.Text), Trim(txtNombre.Text), "A")
Else
  'Excluirlo
    strSQL = "exec spPGX_Usuario_Cliente_Asigna " & Item.Tag & ",'" & Trim(txtUserName.Text) & "','" & glogon.Usuario & "','E',''"
    Call ConectionExecute(strSQL)
    
    Call sbSEGCuentaLog("08", "Exclusión al Rol: (" & Item.Tag & ") " & Trim(Item.Text), glogon.Usuario, txtUserName.Text)
    
    'Sincroniza Core
    Call spCore_Usuario_Sincroniza(Item.Tag, Trim(txtUserName.Text), Trim(txtNombre.Text), "I")

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsultaRoles()

lswRoles.ListItems.Clear

If vPaso Or Not IsNumeric(lblRol.Tag) Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select R.COD_ROL,R.DESCRIPCION,  case when ISNULL( m.REGISTRO_USUARIO  ,'') = '' then 0 else 1 end as 'Asignado'" _
       & "       ,M.REGISTRO_FECHA, M.REGISTRO_USUARIO" _
       & "  from US_ROLES R" _
       & "         left join US_ROL_MIEMBROS M on R.COD_ROL = M.COD_ROL and M.COD_EMPRESA = " & lblRol.Tag _
       & "         and M.USUARIO = '" & Trim(txtUserName.Text) & "'" _
       & "  where R.ACTIVO = 1 and isnull(R.COD_EMPRESA," & lblRol.Tag & ") = " & lblRol.Tag & " order by case when ISNULL( m.REGISTRO_USUARIO  ,'') = '' then 0 else 1 end  desc, R.descripcion"
vPaso = True

With lswRoles.ListItems
    .Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!Descripcion)
          itmX.Tag = rs!cod_rol
          itmX.Checked = rs!Asignado
      rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswMiembros_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub
If lswMiembros.ListItems.Count <= 0 Then Exit Sub

lblRol.Tag = Item.Tag
lblRol.Caption = "Roles para: " & Item.Text
Call sbConsultaRoles
End Sub

Private Sub lswRoles_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswRoles.SortKey = ColumnHeader.Index - 1
  If lswRoles.SortOrder = 0 Then lswRoles.SortOrder = 1 Else lswRoles.SortOrder = 0
  lswRoles.Sorted = True
End Sub

Private Sub lswRoles_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Or lswRoles.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "exec spPGX_Usuario_Rol_Asigna " & lblRol.Tag & ",'" & Trim(txtUserName.Text) & "','" & Item.Tag & "','" & glogon.Usuario & "','I'"
Else
    strSQL = "exec spPGX_Usuario_Rol_Asigna " & lblRol.Tag & ",'" & Trim(txtUserName.Text) & "','" & Item.Tag & "','" & glogon.Usuario & "','E'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

Select Case Item.Index
  Case 1 'Miembro de...
        strSQL = "select C.cod_Empresa,C.Nombre_Corto,isnull(U.Usuario,'') as 'Usuario'" _
               & " from PGX_Clientes C left join PGX_Clientes_Users U on C.cod_Empresa = U.Cod_Empresa" _
               & " and U.Usuario = '" & txtUserName.Text & "'" _
               & " where C.Estado = 'A'" _
               & " order by isnull(U.Usuario,'') desc, C.Nombre_Corto"
        Call OpenRecordSet(rs, strSQL)
        
        vPaso = True
        
        lswRoles.ListItems.Clear
        lblRol.Tag = ""
        lblRol.Caption = "Seleccione a un Cliente!"
        
        With lswMiembros.ListItems
         .Clear
         Do While Not rs.EOF
          Set itmX = .Add(, , rs!Nombre_Corto)
              itmX.Tag = rs!cod_Empresa
              If rs!Usuario <> "" Then
                 itmX.ForeColor = vbBlue
                 itmX.Checked = True
              End If
          rs.MoveNext
         Loop
        End With
        rs.Close
  
        vPaso = False
  
  Case 2 'Bitacora
     Call sbBitacoraLlena

End Select



Exit Sub


vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0


If gEntidad.Tipo = "U" And IsNumeric(gEntidad.UserID) Then
 Call sbConsulta(gEntidad.Usuario)
End If

 'Revisa Credenciales de Administrador de Portal para Abrir cuentas a la libre
If gPortal.Empresa_Id = -1 And Not gAdminAccess.Admin_Portal Then
       MsgBox "Su cuenta de Administración no permite la creación de usuarios en este cliente: " & gPortal.Empresa_Name & ", vbCritical"
       Unload Me
End If

If gPortal.Empresa_Id > 0 And Not gAdminAccess.Rol_LocalUsers Then
       MsgBox "Su cuenta de Administración no permite la creación de usuarios en este cliente: " & gPortal.Empresa_Name & ", vbCritical"
       Unload Me
End If



On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select TFA_IND, TFA_METODO from US_PARAMETROS"
Call OpenRecordSet(rs, strSQL)

v2FA_Indica = rs!TFA_IND
v2FA_Metodo = rs!TFA_METODO

rs.Close

Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Public Sub sbConsulta(vUsuario As String)

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "Select *, isnull(fecha_Mod, registro_fecha) as 'Ultima_Modificacion' " _
       & " from US_Usuarios" _
       & " where Usuario = '" & vUsuario & "'"
       
       
If Not gAdminAccess.Rol_AdminView Then
    strSQL = strSQL & " AND isnull(key_admin,0) = 0"
End If

If Not gAdminAccess.Rol_DirGlobal Then
    strSQL = strSQL & " AND usuario in(select usuario from PGX_CLIENTES_USERS" _
        & " Where cod_Empresa = " & gPortal.Empresa_Id & ")"
End If
       
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
    Call sbBarra_Accion("activo")

    vEdita = True
    
    tcMain.Item(0).Selected = True
    tcMain.Item(1).Enabled = True
    tcMain.Item(2).Enabled = True
    
    
    txtUserName = Trim(rs!Usuario)
    txtUserID = rs!UserID & ""
    txtIdentificacion.Text = rs!Identificacion & ""
    
    txtNombre = IIf(IsNull(rs!Nombre), "", rs!Nombre)
    txtNotas = IIf(IsNull(rs!NOTAS), "", rs!NOTAS)
    
    If rs!ESTADO = "A" Then
       txtEstado.Text = "Activo"
    Else
       txtEstado.Text = "Inactivo"
    End If
    
    dtpIngreso.Value = Format(rs!Registro_Fecha, "dd/mm/yyyy")
    dtpModificacion.Value = Format(rs!Ultima_Modificacion, "dd/mm/yyyy")
    
    txtEMail.Text = rs!EMAIL & ""
    txtTelCelular.Text = rs!Tel_Cell & ""
    txtTelTrabajo.Text = rs!Tel_Trabajo & ""
   
    chkContabiliza.Value = rs!Contabiliza
    
    chk2FA.Value = rs!TFA_IND
    cbo2FA.Text = rs!TFA_METODO
        
Else
   Call sbLimpiaPantalla
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
Dim vMensaje As String

vMensaje = ""

fxValida = True




'Valida que el nombre de usuario esté desocupado
If Not vEdita Then
    strSQL = "select count(*) as 'Existe' from US_USUARIOS" _
           & " where Usuario = '" & txtUserName.Text & "'"
           
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
        vMensaje = vMensaje & vbCrLf & " - Este Usuario se encuentra -Ocupado- debe cambiar el nombre de usuario por otro que se encuentre -Libre-"
    End If
End If


vMensaje = ""

If Trim(txtUserName.Text) = "" Then
   vMensaje = vMensaje & vbCrLf & " - No a indicado el Nombre de Usuario!"
End If

If Trim(txtNombre.Text) = "" Then
   vMensaje = vMensaje & vbCrLf & " - No a indicado el Nombre de la Persona!"
End If

If Trim(txtEMail.Text) = "" Then
   vMensaje = vMensaje & vbCrLf & " - Indicar un Email"
End If

If Trim(txtIdentificacion.Text) = "" Then
   vMensaje = vMensaje & vbCrLf & " - No a indicado la Identificación de la Persona"
End If
  
If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbGuardar()

On Error GoTo vError
 
If Not vEdita Then
  
  strSQL = "select isnull(max(UserID),0) as Ultimo from US_usuarios"
  Call OpenRecordSet(rs, strSQL)
    txtUserID.Text = rs!Ultimo + 1
  rs.Close
  

  strSQL = "Insert US_Usuarios(Usuario,Identificacion,Nombre,Registro_Fecha,Registro_Usuario" _
         & ",Fecha_Mod,Estado,UserID,Notas,email,Tel_Cell,Tel_Trabajo,Contabiliza, PROGRX_THEME, TFA_IND, TFA_METODO)" _
         & " values('" & Trim(txtUserName) & "','" & Trim(txtIdentificacion.Text) & "','" & Trim(txtNombre) _
         & "', getdate(),'" & glogon.Usuario & "','" & Format(dtpModificacion, "yyyy/mm/dd") _
         & "', 'A'," & txtUserID.Text & ",'" & Trim(txtNotas) & "','" & Trim(txtEMail) _
         & "', '" & txtTelCelular.Text & "','" & txtTelTrabajo.Text & "'," & chkContabiliza.Value _
         & ", 'Default', " & chk2FA.Value & ", '" & cbo2FA.Text & "')"
Else
  strSQL = "Update US_Usuarios Set Nombre = '" & Trim(txtNombre) _
         & "', Identificacion = '" & txtIdentificacion.Text _
         & "', Fecha_Mod= getdate(),notas= '" & Trim(txtNotas.Text) _
         & "', Email = '" & Trim(txtEMail) & "',Tel_Cell = '" & Trim(txtTelCelular.Text) _
         & "', Tel_Trabajo = '" & Trim(txtTelTrabajo.Text) & "', Contabiliza  = " & chkContabiliza.Value _
         & ", TFA_IND = " & chk2FA.Value & ", TFA_METODO = '" & cbo2FA.Text _
         & "' Where UserID = " & txtUserID.Text
End If
Call ConectionExecute(strSQL)

If Not vEdita Then
    Call sbSEGCuentaLog("01", , glogon.Usuario, txtUserName.Text)
    
    If Not gAdminAccess.Admin_Portal And gPortal.Empresa_Id > 0 Then
            strSQL = "exec spPGX_Usuario_Cliente_Asigna " & gPortal.Empresa_Id & ",'" & Trim(txtUserName.Text) & "','" & glogon.Usuario & "','I',''"
            Call ConectionExecute(strSQL)
        
            Call sbSEGCuentaLog("08", "Membresía al Rol: (" & gPortal.Empresa_Id & ") " & gPortal.Empresa_Name, glogon.Usuario, txtUserName.Text)
        
            'Sincroniza Core
            Call spCore_Usuario_Sincroniza(gPortal.Empresa_Id, Trim(txtUserName.Text), Trim(txtNombre.Text), "A")
    End If
    
Else
    Call sbSEGCuentaLog("15", , glogon.Usuario, txtUserName.Text)
End If

Call sbBarra_Accion("activo")
Call sbConsulta(txtUserName)

        
vEdita = True
        
MsgBox "Información guardada satisfactoriamente...", vbInformation

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then

'  strSQL = "delete Miembros where Nombre = '" & txtUserName & "'"
'  Call ConectionExecute(strSQL)
'
'  strSQL = "delete permisos where Tipo = 'U' and Nombre = '" & txtUserID & "'"
'  Call ConectionExecute(strSQL)
'
'  strSQL = "delete usuarios where UserID = " & txtUserID
'  Call ConectionExecute(strSQL)
'
'  Call Bitacora("Elimina", "Usuario: " & txtUserName)

  Call sbLimpiaPantalla
  Call sbBarra_Accion("Nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelCelular.SetFocus
End Sub


Private Sub txtLineas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call sbBitacoraLlena
End If
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkContabiliza.SetFocus
End Sub

Private Sub txtTelCelular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelTrabajo.SetFocus
End Sub

Private Sub txtTelTrabajo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEMail.SetFocus
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 
 If vEdita Then Call sbConsulta(txtUserName)
 txtIdentificacion.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Consulta = "select Usuario,Nombre from US_usuarios"
    
    gBusquedas.Filtro = ""
    
    If Not gAdminAccess.Rol_AdminView Then
        gBusquedas.Filtro = " AND isnull(key_admin,0) = 0"
    End If
    
    If Not gAdminAccess.Rol_DirGlobal Then
        gBusquedas.Filtro = gBusquedas.Filtro & " AND usuario in(select usuario from PGX_CLIENTES_USERS" _
            & " Where cod_Empresa = " & gPortal.Empresa_Id & ")"
    End If
    
    frmBusquedas.Show vbModal
    txtUserName = gBusquedas.Resultado
    txtNombre = gBusquedas.Resultado2
    Call sbConsulta(txtUserName)
    txtUserName.SetFocus
End If

End Sub



