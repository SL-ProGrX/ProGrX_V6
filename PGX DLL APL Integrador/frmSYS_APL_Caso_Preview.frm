VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmSYS_APL_Caso_Preview 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "APL: Visualización de Caso"
   ClientHeight    =   7524
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   13224
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7524
   ScaleWidth      =   13224
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3492
      Left            =   240
      TabIndex        =   25
      Top             =   3840
      Width           =   12732
      _Version        =   1310720
      _ExtentX        =   22458
      _ExtentY        =   6159
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
      ItemCount       =   7
      SelectedItem    =   2
      Item(0).Caption =   "Contacto"
      Item(0).ControlCount=   13
      Item(0).Control(0)=   "Label1(15)"
      Item(0).Control(1)=   "Label1(14)"
      Item(0).Control(2)=   "Label1(2)"
      Item(0).Control(3)=   "Label1(1)"
      Item(0).Control(4)=   "Label1(0)"
      Item(0).Control(5)=   "lswFactura"
      Item(0).Control(6)=   "txtInstitucion"
      Item(0).Control(7)=   "txtDepartamento"
      Item(0).Control(8)=   "txtEmail"
      Item(0).Control(9)=   "txtTelefono"
      Item(0).Control(10)=   "txtTelMovil"
      Item(0).Control(11)=   "txtSeccion"
      Item(0).Control(12)=   "Label1(6)"
      Item(1).Caption =   "Financiamiento"
      Item(1).ControlCount=   15
      Item(1).Control(0)=   "Label1(21)"
      Item(1).Control(1)=   "Label1(20)"
      Item(1).Control(2)=   "Label1(19)"
      Item(1).Control(3)=   "Label1(18)"
      Item(1).Control(4)=   "Label1(17)"
      Item(1).Control(5)=   "Label1(4)"
      Item(1).Control(6)=   "Label1(3)"
      Item(1).Control(7)=   "gbResolucion"
      Item(1).Control(8)=   "txtLinea"
      Item(1).Control(9)=   "txtPlan"
      Item(1).Control(10)=   "txtMonto"
      Item(1).Control(11)=   "txtTasa"
      Item(1).Control(12)=   "txtPlazo"
      Item(1).Control(13)=   "txtCuota"
      Item(1).Control(14)=   "txtCoreId"
      Item(2).Caption =   "Referencias"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lswReferencias"
      Item(3).Caption =   "Fiadores"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "lswFiadores"
      Item(4).Caption =   "Adjuntos"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "lswAdjuntos"
      Item(5).Caption =   "Estados"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "lswEstados"
      Item(6).Caption =   "Triangulación"
      Item(6).ControlCount=   27
      Item(6).Control(0)=   "Label1(33)"
      Item(6).Control(1)=   "Label1(34)"
      Item(6).Control(2)=   "Label1(35)"
      Item(6).Control(3)=   "Label1(36)"
      Item(6).Control(4)=   "Label1(37)"
      Item(6).Control(5)=   "Label1(38)"
      Item(6).Control(6)=   "txtFormaliza_Fecha"
      Item(6).Control(7)=   "txtFormaliza_Usuario"
      Item(6).Control(8)=   "txtFormaliza_Documento"
      Item(6).Control(9)=   "Label1(39)"
      Item(6).Control(10)=   "Label1(40)"
      Item(6).Control(11)=   "Label1(41)"
      Item(6).Control(12)=   "txtCobro_Fecha"
      Item(6).Control(13)=   "txtCobro_Usuario"
      Item(6).Control(14)=   "txtCobro_Estado"
      Item(6).Control(15)=   "Label1(42)"
      Item(6).Control(16)=   "Label1(43)"
      Item(6).Control(17)=   "Label1(44)"
      Item(6).Control(18)=   "txtCancela_Fecha"
      Item(6).Control(19)=   "txtCancela_Usuario"
      Item(6).Control(20)=   "txtCancela_Documento"
      Item(6).Control(21)=   "Label1(45)"
      Item(6).Control(22)=   "txtCobro_Remesa"
      Item(6).Control(23)=   "Label1(46)"
      Item(6).Control(24)=   "txtCancela_Tipo"
      Item(6).Control(25)=   "txtFormaliza_Nota"
      Item(6).Control(26)=   "Label1(47)"
      Begin XtremeSuiteControls.ListView lswEstados 
         Height          =   2892
         Left            =   -69880
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   12492
         _Version        =   1310720
         _ExtentX        =   22034
         _ExtentY        =   5101
         _StockProps     =   77
         BackColor       =   -2147483643
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswAdjuntos 
         Height          =   2892
         Left            =   -69880
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   12492
         _Version        =   1310720
         _ExtentX        =   22034
         _ExtentY        =   5101
         _StockProps     =   77
         BackColor       =   -2147483643
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswFiadores 
         Height          =   2892
         Left            =   -69880
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   12492
         _Version        =   1310720
         _ExtentX        =   22034
         _ExtentY        =   5101
         _StockProps     =   77
         BackColor       =   -2147483643
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswReferencias 
         Height          =   2772
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   12252
         _Version        =   1310720
         _ExtentX        =   21611
         _ExtentY        =   4890
         _StockProps     =   77
         BackColor       =   -2147483643
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswFactura 
         Height          =   2892
         Left            =   -63520
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1310720
         _ExtentX        =   10816
         _ExtentY        =   5101
         _StockProps     =   77
         BackColor       =   -2147483643
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtCancela_Tipo 
         Height          =   312
         Left            =   -60520
         TabIndex        =   100
         Top             =   2400
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCobro_Remesa 
         Height          =   312
         Left            =   -64600
         TabIndex        =   98
         Top             =   2400
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCancela_Fecha 
         Height          =   312
         Left            =   -60520
         TabIndex        =   94
         Top             =   960
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCobro_Fecha 
         Height          =   312
         Left            =   -64600
         TabIndex        =   88
         Top             =   960
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.GroupBox gbResolucion 
         Height          =   4212
         Left            =   -63400
         TabIndex        =   53
         Top             =   360
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1310720
         _ExtentX        =   10816
         _ExtentY        =   7429
         _StockProps     =   79
         Caption         =   "Resolución"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtResolucionFecha 
            Height          =   312
            Left            =   2280
            TabIndex        =   54
            Top             =   600
            Width           =   2052
            _Version        =   1310720
            _ExtentX        =   3619
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         End
         Begin XtremeSuiteControls.FlatEdit txtResolucionUsuario 
            Height          =   312
            Left            =   240
            TabIndex        =   55
            Top             =   600
            Width           =   2052
            _Version        =   1310720
            _ExtentX        =   3619
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         End
         Begin XtremeSuiteControls.FlatEdit txtResolucionNota 
            Height          =   792
            Left            =   240
            TabIndex        =   58
            Top             =   1320
            Width           =   5652
            _Version        =   1310720
            _ExtentX        =   9970
            _ExtentY        =   1397
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   28
            Left            =   240
            TabIndex        =   59
            Top             =   1080
            Width           =   1092
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   27
            Left            =   2280
            TabIndex        =   57
            Top             =   360
            Width           =   1092
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   26
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   1572
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtTelefono 
         Height          =   312
         Left            =   -69760
         TabIndex        =   29
         Top             =   3000
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtTelMovil 
         Height          =   312
         Left            =   -67720
         TabIndex        =   30
         Top             =   3000
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtInstitucion 
         Height          =   312
         Left            =   -69760
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1310720
         _ExtentX        =   10816
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDepartamento 
         Height          =   312
         Left            =   -69760
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1310720
         _ExtentX        =   10816
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSeccion 
         Height          =   312
         Left            =   -69760
         TabIndex        =   65
         Top             =   1800
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1310720
         _ExtentX        =   10816
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   312
         Left            =   -69760
         TabIndex        =   28
         Top             =   2400
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1310720
         _ExtentX        =   10816
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPlan 
         Height          =   312
         Left            =   -69760
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1310720
         _ExtentX        =   10816
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLinea 
         Height          =   312
         Left            =   -69760
         TabIndex        =   36
         Top             =   720
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1310720
         _ExtentX        =   10816
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   -69760
         TabIndex        =   38
         Top             =   2040
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1310720
         _ExtentX        =   2773
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   312
         Left            =   -68200
         TabIndex        =   39
         Top             =   2040
         Visible         =   0   'False
         Width           =   852
         _Version        =   1310720
         _ExtentX        =   1503
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   312
         Left            =   -67360
         TabIndex        =   40
         Top             =   2040
         Visible         =   0   'False
         Width           =   852
         _Version        =   1310720
         _ExtentX        =   1503
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuota 
         Height          =   312
         Left            =   -66520
         TabIndex        =   41
         Top             =   2040
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCoreId 
         Height          =   312
         Left            =   -65080
         TabIndex        =   75
         Top             =   2040
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1310720
         _ExtentX        =   2561
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFormaliza_Fecha 
         Height          =   312
         Left            =   -68680
         TabIndex        =   82
         Top             =   960
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtFormaliza_Usuario 
         Height          =   312
         Left            =   -68680
         TabIndex        =   83
         Top             =   1440
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtFormaliza_Documento 
         Height          =   312
         Left            =   -68680
         TabIndex        =   84
         Top             =   1920
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCobro_Usuario 
         Height          =   312
         Left            =   -64600
         TabIndex        =   89
         Top             =   1440
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCobro_Estado 
         Height          =   312
         Left            =   -64600
         TabIndex        =   90
         Top             =   1920
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCancela_Usuario 
         Height          =   312
         Left            =   -60520
         TabIndex        =   95
         Top             =   1440
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCancela_Documento 
         Height          =   312
         Left            =   -60520
         TabIndex        =   96
         Top             =   1920
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtFormaliza_Nota 
         Height          =   912
         Left            =   -68680
         TabIndex        =   101
         Top             =   2400
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   1609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   47
         Left            =   -69880
         TabIndex        =   102
         Top             =   2400
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   46
         Left            =   -61720
         TabIndex        =   99
         Top             =   2400
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remesa:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   45
         Left            =   -65800
         TabIndex        =   97
         Top             =   2400
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   44
         Left            =   -61720
         TabIndex        =   93
         Top             =   1920
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   43
         Left            =   -61720
         TabIndex        =   92
         Top             =   1440
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   42
         Left            =   -61720
         TabIndex        =   91
         Top             =   960
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   41
         Left            =   -65800
         TabIndex        =   87
         Top             =   1920
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   40
         Left            =   -65800
         TabIndex        =   86
         Top             =   1440
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   39
         Left            =   -65800
         TabIndex        =   85
         Top             =   960
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   38
         Left            =   -69880
         TabIndex        =   81
         Top             =   1920
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   37
         Left            =   -69880
         TabIndex        =   80
         Top             =   1440
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   36
         Left            =   -69880
         TabIndex        =   79
         Top             =   960
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelación:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   35
         Left            =   -61720
         TabIndex        =   78
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cobros:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   34
         Left            =   -65800
         TabIndex        =   77
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Formalización:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   33
         Left            =   -69880
         TabIndex        =   76
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sección:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   6
         Left            =   -69760
         TabIndex        =   66
         Top             =   1560
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pyme:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   -69760
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan de Financiamiento:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   -69760
         TabIndex        =   47
         Top             =   1080
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   17
         Left            =   -69760
         TabIndex        =   46
         Top             =   1800
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasa:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   18
         Left            =   -68200
         TabIndex        =   45
         Top             =   1800
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   19
         Left            =   -67360
         TabIndex        =   44
         Top             =   1800
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuota:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   20
         Left            =   -66520
         TabIndex        =   43
         Top             =   1800
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Operación:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   21
         Left            =   -65080
         TabIndex        =   42
         Top             =   1800
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Institución:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   -69760
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   -69760
         TabIndex        =   34
         Top             =   960
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   -69760
         TabIndex        =   33
         Top             =   2160
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   14
         Left            =   -69760
         TabIndex        =   32
         Top             =   2760
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. Móvil:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   15
         Left            =   -67720
         TabIndex        =   31
         Top             =   2760
         Visible         =   0   'False
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.GroupBox gbCliente 
      Height          =   2532
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   6372
      _Version        =   1310720
      _ExtentX        =   11239
      _ExtentY        =   4466
      _StockProps     =   79
      Caption         =   "Datos del Cliente"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtGenero 
         Height          =   312
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtProfesion 
         Height          =   312
         Left            =   4200
         TabIndex        =   10
         Top             =   1320
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   6252
         _Version        =   1310720
         _ExtentX        =   11028
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEstadoCivil 
         Height          =   312
         Left            =   2160
         TabIndex        =   9
         Top             =   1320
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtProvincia 
         Height          =   312
         Left            =   120
         TabIndex        =   69
         Top             =   2040
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCanton 
         Height          =   312
         Left            =   2160
         TabIndex        =   70
         Top             =   2040
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtDistrito 
         Height          =   312
         Left            =   4200
         TabIndex        =   71
         Top             =   2040
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   32
         Left            =   120
         TabIndex        =   74
         Top             =   1800
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantón:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   31
         Left            =   2160
         TabIndex        =   73
         Top             =   1800
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   7
         Left            =   4200
         TabIndex        =   72
         Top             =   1800
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Profesión:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   12
         Left            =   4200
         TabIndex        =   13
         Top             =   1080
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   11
         Left            =   2160
         TabIndex        =   12
         Top             =   1080
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Genero:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   552
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   4692
      _Version        =   1310720
      _ExtentX        =   8276
      _ExtentY        =   974
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   552
      Left            =   6480
      TabIndex        =   4
      Top             =   120
      Width           =   3252
      _Version        =   1310720
      _ExtentX        =   5736
      _ExtentY        =   974
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbVenta 
      Height          =   2652
      Left            =   6720
      TabIndex        =   16
      Top             =   1200
      Width           =   6492
      _Version        =   1310720
      _ExtentX        =   11451
      _ExtentY        =   4678
      _StockProps     =   79
      Caption         =   "Venta"
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtFacturaNotas 
         Height          =   672
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   6132
         _Version        =   1310720
         _ExtentX        =   10816
         _ExtentY        =   1185
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtFacturaNo 
         Height          =   312
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtFacturaFecha 
         Height          =   312
         Left            =   2160
         TabIndex        =   20
         Top             =   480
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin XtremeSuiteControls.FlatEdit txtFacturaMonto 
         Height          =   312
         Left            =   4200
         TabIndex        =   21
         Top             =   480
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLugarFirma 
         Height          =   312
         Left            =   120
         TabIndex        =   67
         Top             =   2160
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lugar de Firma:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   13
         Left            =   120
         TabIndex        =   68
         Top             =   1920
         Width           =   1572
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   25
         Left            =   4200
         TabIndex        =   24
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   24
         Left            =   2160
         TabIndex        =   23
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Factura:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   23
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   22
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtRegistroUsuario 
      Height          =   312
      Left            =   10920
      TabIndex        =   61
      Top             =   360
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   End
   Begin XtremeSuiteControls.FlatEdit txtRegistroFecha 
      Height          =   312
      Left            =   10920
      TabIndex        =   62
      Top             =   720
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   312
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   2412
      _Version        =   1310720
      _ExtentX        =   4254
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDominio 
      Height          =   312
      Left            =   5280
      TabIndex        =   14
      Top             =   720
      Width           =   4452
      _Version        =   1310720
      _ExtentX        =   7853
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registra:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   30
      Left            =   9840
      TabIndex        =   64
      Top             =   360
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   29
      Left            =   9840
      TabIndex        =   63
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dominio:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   16
      Left            =   4200
      TabIndex        =   15
      Top             =   720
      Width           =   2532
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   9
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No. Tramite:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   5
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmSYS_APL_Caso_Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub sbConsulta_Externa(pDominio As String, pOperacion As Long)
    Call sbConsulta(pDominio, pOperacion)
End Sub


Private Sub sbConsulta(Optional pDominio As String = "", Optional pOperacion As Long = 0)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass
    
tcMain.Item(0).Selected = True
    
    
If pDominio = "" Then
    strSQL = "select * from vAPL_Analisis_Main" _
           & " where APL_OPERACION = " & pOperacion

Else
    strSQL = "select * from vAPL_Analisis_Main" _
           & " where COD_DOMINIO = '" & pDominio & "'" _
           & " and APL_OPERACION = " & pOperacion
End If
    
txtIdentificacion.SetFocus
    
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
       
   pDominio = rs!Cod_Dominio
       
   txtOperacion.Text = CStr(rs!APL_OPERACION & "")
   txtEstado.Text = rs!Estado_Desc & ""
   txtRegistroFecha.Text = rs!registro_Fecha & ""
   txtRegistroUsuario.Text = rs!registro_usuario & ""
   txtIdentificacion.Text = rs!Cedula & ""
   txtDominio.Text = rs!Cod_Dominio
   txtNombre.Text = rs!Nombre & ""
   txtGenero.Text = "No Indica"
   txtEstadoCivil.Text = rs!ESTADO_CIVIL & ""
   txtProfesion.Text = rs!PROFESION & ""
   txtProvincia.Text = rs!Provincia_Desc
   txtCanton.Text = rs!Canton_Desc
   txtDistrito.Text = rs!Distrito_Desc
   txtFacturaNo.Text = rs!FACTURA_NUMERO & ""
   txtFacturaMonto.Text = Format(rs!Factura_Monto, "Standard")
   txtFacturaFecha.Text = rs!Factura_Fecha & ""
   txtFacturaNotas.Text = rs!Factura_Detalle & ""
   txtLugarFirma.Text = rs!Lugar_Firma & ""
   txtInstitucion.Text = rs!Institucion_desc
   txtDepartamento.Text = rs!Departamento_desc
   txtSeccion.Text = rs!Seccion_Desc
   txtEmail.Text = rs!CLIENTE_EMAIL & ""
   txtTelefono.Text = rs!CLIENTE_TELEFONO & ""
   txtTelMovil.Text = rs!CLIENTE_CELULAR & ""
   txtLinea.Text = rs!Linea_Desc
   txtPlan.Text = rs!Plan_desc
   txtMonto.Text = Format(rs!Factura_Monto, "Standard")
   txtTasa.Text = Format(rs!Tasa, "Standard")
   txtPlazo.Text = CStr(rs!Plazo)
   txtCuota.Text = Format(rs!Cuota, "Standard")
   txtCoreId.Text = CStr(rs!Operacion & "")

   txtResolucionFecha.Text = rs!Resolucion_Fecha & ""
   txtResolucionUsuario.Text = rs!Resolucion_Usuario & ""
   txtResolucionNota.Text = rs!notas & vbCrLf & rs!Resolucion_Notas & ""
   
   txtFormaliza_Documento.Text = rs!Formaliza_Documento & ""
   txtFormaliza_Fecha.Text = rs!Formaliza_Fecha & ""
   txtFormaliza_Usuario.Text = rs!Formaliza_Usuario & ""
   txtFormaliza_Nota.Text = rs!Formaliza_Nota & ""
   
   txtCobro_Estado.Text = rs!Cobro_Estado & ""
   txtCobro_Fecha.Text = rs!Cobro_Fecha & ""
   txtCobro_Remesa.Text = rs!Cobro_Remesa & ""
   txtCobro_Usuario.Text = rs!Cobro_Usuario & ""
       
   txtCancela_Documento.Text = rs!Cancela_Documento & ""
   txtCancela_Fecha.Text = rs!Cancela_Fecha & ""
   txtCancela_Tipo.Text = rs!Cancela_Tipo & ""
   txtCancela_Usuario.Text = rs!Cancela_Usuario & ""
   
      
   With lswFiadores.ListItems
    .Clear
    Set itmX = .Add(, , "No.1")
        itmX.SubItems(1) = rs!Fiador1_Cedula & ""
        itmX.SubItems(2) = rs!Fiador1_Nombre & ""
        itmX.SubItems(3) = rs!Fiador1_Telefono & ""
    
   End With 'Fiadores
       
   
Else
   lswFiadores.ListItems.Clear
   txtOperacion.Text = ""
   txtEstado.Text = ""
   txtRegistroFecha.Text = ""
   txtRegistroUsuario.Text = ""
   txtIdentificacion.Text = ""
   txtDominio.Text = ""
   txtNombre.Text = ""
   txtGenero.Text = ""
   txtEstadoCivil.Text = ""
   txtProfesion.Text = ""
   txtProvincia.Text = ""
   txtCanton.Text = ""
   txtDistrito.Text = ""
   txtFacturaNo.Text = ""
   txtFacturaMonto.Text = "0"
   txtFacturaFecha.Text = ""
   txtFacturaNotas.Text = ""
   txtLugarFirma.Text = ""
   txtInstitucion.Text = ""
   txtDepartamento.Text = ""
   txtSeccion.Text = ""
   txtEmail.Text = ""
   txtTelefono.Text = ""
   txtTelMovil.Text = ""
   txtLinea.Text = ""
   txtPlan.Text = ""
   txtMonto.Text = "0"
   txtTasa.Text = "0"
   txtPlazo.Text = "1"
   txtCuota.Text = "0"
   txtCoreId.Text = ""

   txtResolucionFecha.Text = ""
   txtResolucionUsuario.Text = ""
   txtResolucionNota.Text = ""
   
   txtFormaliza_Documento.Text = ""
   txtFormaliza_Fecha.Text = ""
   txtFormaliza_Usuario.Text = ""
   txtFormaliza_Nota.Text = ""
   
   txtCobro_Estado.Text = ""
   txtCobro_Fecha.Text = ""
   txtCobro_Remesa.Text = ""
   txtCobro_Usuario.Text = ""
       
   txtCancela_Documento.Text = ""
   txtCancela_Fecha.Text = ""
   txtCancela_Tipo.Text = ""
   txtCancela_Usuario.Text = ""

End If
rs.Close

'Detalle de la Factura
With lswFactura.ListItems
    .Clear
    
    strSQL = "select L.*, C.DESCRIPCION as 'CONCEPTO_DESC'" _
           & " from APL_OPERACIONES_CONCEPTOS L INNER JOIN APL_CONCEPTOS C on L.COD_DOMINIO = C.COD_DOMINIO" _
           & " and L.COD_CONCEPTO = C.COD_CONCEPTO" _
           & " where l.COD_DOMINIO = '" & pDominio & "' AND L.APL_OPERACION = " & pOperacion
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = .Add(, , rs!Num_Linea)
           itmX.SubItems(1) = rs!cod_Concepto
           itmX.SubItems(2) = rs!Concepto_Desc
           itmX.SubItems(3) = rs!Cantidad
           itmX.SubItems(4) = Format(rs!Monto, "Standard")
       rs.MoveNext
    Loop
    rs.Close

End With


'Referencias
With lswReferencias.ListItems
    .Clear
    
    strSQL = " select *" _
           & " From APL_OPERACIONES_REFERENCIAS" _
           & " Where COD_DOMINIO = '" & pDominio & "' and APL_OPERACION = " & pOperacion
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = .Add(, , rs!cod_referencia)
           itmX.SubItems(1) = rs!Cedula & ""
           itmX.SubItems(2) = rs!Nombre & ""
           itmX.SubItems(3) = rs!Parentesco & ""
           itmX.SubItems(4) = rs!Telefono & ""
       rs.MoveNext
    Loop
    rs.Close
End With



'Adjuntos
With lswAdjuntos.ListItems
    .Clear
    
    strSQL = " select TIPO,ARCHIVO_NOMBRE, ARCHIVO_TIPO, REGISTRO_FECHA, REGISTRO_USUARIO" _
           & " From APL_OPERACIONES_ADJUNTOS" _
           & " Where COD_DOMINIO = '" & pDominio & "' and APL_OPERACION = " & pOperacion
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = .Add(, , rs!Tipo)
           itmX.SubItems(1) = rs!Archivo_Tipo & ""
           itmX.SubItems(2) = rs!Archivo_Nombre & ""
           itmX.SubItems(3) = rs!registro_usuario & ""
           itmX.SubItems(4) = rs!registro_Fecha & ""
       rs.MoveNext
    Loop
    rs.Close
End With


'Estados
With lswEstados.ListItems
    .Clear
    
    strSQL = " select * " _
           & " , case when ESTADO = 'D' then 'Denegado' when ESTADO = 'R' then 'Recibida'" _
           & " when ESTADO = 'P' then 'Pendiente' when ESTADO = 'A' then 'Autorizada'" _
           & " when ESTADO = 'F' then 'Formalizada' when ESTADO = 'N' then 'Anulada' else 'Denegadas' end as 'ESTADO_DESC'" _
           & " From APL_OPERACIONES_ESTADOS" _
           & " Where COD_DOMINIO = '" & pDominio & "' and APL_OPERACION = " & pOperacion

    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = .Add(, , rs!Num_Linea)
           itmX.SubItems(1) = rs!Estado_Desc
           itmX.SubItems(2) = rs!notas & ""
           itmX.SubItems(3) = rs!registro_Fecha & ""
           itmX.SubItems(4) = rs!registro_usuario & ""
       rs.MoveNext
    Loop
    rs.Close
End With




Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
vModulo = 38

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lswFactura.ColumnHeaders
    .Clear
    .Add , , "Ln", 300
    .Add , , "Cod.", 500
    .Add , , "Descripción", 2500
    .Add , , "Qty", 1000, vbCenter
    .Add , , "Monto", 1750, vbRightJustify
End With


With lswReferencias.ColumnHeaders
    .Clear
    .Add , , "Ln", 300
    .Add , , "Identificación", 2100
    .Add , , "Nombre", 4500
    .Add , , "Parentesco", 2000, vbCenter
    .Add , , "Teléfono", 1750, vbCenter
End With

With lswAdjuntos.ColumnHeaders
    .Clear
    .Add , , "Tipo", 1400
    .Add , , "Archivo Tipo:", 1600
    .Add , , "Archivo Nombre:", 5000
    .Add , , "Usuario", 2000, vbCenter
    .Add , , "Fecha", 2100, vbCenter
End With

With lswFiadores.ColumnHeaders
    .Clear
    .Add , , "Ln", 300
    .Add , , "Identificación", 2100
    .Add , , "Nombre", 4500
    .Add , , "Teléfono", 1750, vbCenter
End With

With lswEstados.ColumnHeaders
    .Clear
    .Add , , "Ln", 300
    .Add , , "Estado", 2100
    .Add , , "Notas", 4500
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuairo", 2100, vbCenter
End With


tcMain.Item(0).Selected = True

End Sub


Private Sub txtOperacion_LostFocus()
If IsNumeric(txtOperacion.Text) Then
  Call sbConsulta("", txtOperacion.Text)
End If
End Sub
