VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCO_Reportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros: Informes de Antiguedad/Mora (Días Reales)"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   10305
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4095
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   3735
      _Version        =   1441793
      _ExtentX        =   6588
      _ExtentY        =   7223
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
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.GroupBox gbInforme 
      Height          =   972
      Left            =   120
      TabIndex        =   21
      Top             =   5880
      Width           =   10212
      _Version        =   1441793
      _ExtentX        =   18013
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   6720
         TabIndex        =   22
         Top             =   240
         Width           =   2532
         _Version        =   1441793
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin MSComctlLib.Toolbar tlb 
         Height          =   312
         Left            =   6720
         TabIndex        =   23
         Top             =   600
         Width           =   2532
         _ExtentX        =   4471
         _ExtentY        =   556
         ButtonWidth     =   1799
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgArbol"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "Reporte"
               Object.ToolTipText     =   "Informe (Reporte Seleccionado)"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cubo"
               Key             =   "Cubos"
               Object.ToolTipText     =   "Actualiza Cubo de Información en Analisis"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   5172
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4092
      Left            =   3840
      TabIndex        =   1
      Top             =   1680
      Width           =   6372
      _Version        =   1441793
      _ExtentX        =   11239
      _ExtentY        =   7218
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
      ItemCount       =   5
      Item(0).Caption =   "General"
      Item(0).ControlCount=   18
      Item(0).Control(0)=   "cboDestino"
      Item(0).Control(1)=   "cboRecurso"
      Item(0).Control(2)=   "cboInstitucion"
      Item(0).Control(3)=   "cboDeductora"
      Item(0).Control(4)=   "chkLineas"
      Item(0).Control(5)=   "txtCodigo"
      Item(0).Control(6)=   "txtDescripcion"
      Item(0).Control(7)=   "Label1(18)"
      Item(0).Control(8)=   "Label1(15)"
      Item(0).Control(9)=   "Label1(13)"
      Item(0).Control(10)=   "Label1(37)"
      Item(0).Control(11)=   "cboComite"
      Item(0).Control(12)=   "Label1(1)"
      Item(0).Control(13)=   "Label1(2)"
      Item(0).Control(14)=   "Label1(0)"
      Item(0).Control(15)=   "cboEstadoLaboral"
      Item(0).Control(16)=   "Label1(3)"
      Item(0).Control(17)=   "cboDivisa"
      Item(1).Caption =   "Antiguedad"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "chkAntiguedad"
      Item(1).Control(1)=   "lswAntiguedad"
      Item(2).Caption =   "Garantía"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "chkGarantias"
      Item(2).Control(1)=   "lswGarantias"
      Item(3).Caption =   "Estados"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "chkPersona"
      Item(3).Control(1)=   "lswEstados"
      Item(4).Caption =   "Carteras"
      Item(4).ControlCount=   2
      Item(4).Control(0)=   "chkCarteras"
      Item(4).Control(1)=   "lswCarteras"
      Begin XtremeSuiteControls.ListView lswAntiguedad 
         Height          =   3252
         Left            =   -69880
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1441793
         _ExtentX        =   10816
         _ExtentY        =   5736
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
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboDestino 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Appearance      =   7
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboRecurso 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1200
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Appearance      =   7
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboInstitucion 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   2040
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Appearance      =   7
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboDeductora 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   2400
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Appearance      =   7
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkLineas 
         Height          =   255
         Left            =   5880
         TabIndex        =   6
         Top             =   480
         Width           =   255
         _Version        =   1441793
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   480
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   480
         Width           =   3855
         _Version        =   1441793
         _ExtentX        =   6800
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboComite 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   1560
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Appearance      =   7
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   3120
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Appearance      =   7
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoLaboral 
         Height          =   315
         Left            =   1080
         TabIndex        =   18
         Top             =   3480
         Width           =   4575
         _Version        =   1441793
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Appearance      =   7
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkAntiguedad 
         Height          =   252
         Left            =   -69880
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkGarantias 
         Height          =   252
         Left            =   -69880
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.ListView lswGarantias 
         Height          =   3252
         Left            =   -69880
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1441793
         _ExtentX        =   10816
         _ExtentY        =   5736
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
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkPersona 
         Height          =   252
         Left            =   -69880
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.ListView lswEstados 
         Height          =   3252
         Left            =   -69880
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   6132
         _Version        =   1441793
         _ExtentX        =   10816
         _ExtentY        =   5736
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
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkCarteras 
         Height          =   252
         Left            =   -69880
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.ListView lswCarteras 
         Height          =   3255
         Left            =   -69880
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   6135
         _Version        =   1441793
         _ExtentX        =   10816
         _ExtentY        =   5736
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
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "L.Laboral"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Divisa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Línea"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Comité"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Deductora"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   37
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   13
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Recurso"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   15
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Institución"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   18
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   975
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9120
      Top             =   120
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   9600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Reportes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Reportes.frx":6862
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Left            =   3840
      TabIndex        =   34
      Top             =   1320
      Width           =   6372
      _Version        =   1441793
      _ExtentX        =   11239
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Filtros:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption lblReporte 
      Height          =   372
      Left            =   120
      TabIndex        =   33
      Top             =   1320
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6583
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de Cobros"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   2160
      TabIndex        =   0
      Top             =   300
      Width           =   5172
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10332
   End
End
Attribute VB_Name = "frmCO_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbReporteListados()
Dim strSQL As String, vSubTitulo As String
Dim i As Byte, iCantidad As Integer, vCadena As String
Dim vReporte As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vSubTitulo = ""

strSQL = "{vCbrListadoMoraDiasReal.Estado} = 'A'"

If cboDivisa.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{vCbrListadoMoraDiasReal.COD_DIVISA} = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
  vSubTitulo = vSubTitulo & "Divisa: " & cboDivisa.ItemData(cboDivisa.ListIndex) & " ¦"
End If

If chkLineas.Value = vbUnchecked Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{vCbrListadoMoraDiasReal.CODIGO} = '" & txtCodigo.Text & "'"
  vSubTitulo = vSubTitulo & "Línea : " & txtCodigo.Text & " ¦"
End If


If cboDestino.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{vCbrListadoMoraDiasReal.COD_DESTINO} = '" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
  vSubTitulo = vSubTitulo & "Destino: " & cboDestino.ItemData(cboDestino.ListIndex) & " ¦"
End If

If cboRecurso.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{vCbrListadoMoraDiasReal.COD_GRUPO} = '" & cboRecurso.ItemData(cboRecurso.ListIndex) & "'"
  vSubTitulo = vSubTitulo & "Recurso: " & cboRecurso.ItemData(cboRecurso.ListIndex) & " ¦"
End If


If cboComite.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{vCbrListadoMoraDiasReal.ID_COMITE} = " & cboComite.ItemData(cboComite.ListIndex)
  vSubTitulo = vSubTitulo & "Comité Id: " & cboComite.ItemData(cboComite.ListIndex) & " ¦"
End If

If cboEstadoLaboral.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{vCbrListadoMoraDiasReal.ESTADO_LABORAL} = '" & cboEstadoLaboral.ItemData(cboEstadoLaboral.ListIndex) & "'"
  vSubTitulo = vSubTitulo & "E.Laboral: " & cboEstadoLaboral.Text & " ¦"
End If


If cboInstitucion.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{vCbrListadoMoraDiasReal.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
  vSubTitulo = vSubTitulo & "Inst: " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & " ¦"
End If

If cboDeductora.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{vCbrListadoMoraDiasReal.COD_DEDUCTORA} = " & cboDeductora.ItemData(cboDeductora.ListIndex)
  vSubTitulo = vSubTitulo & "Deductora: " & cboDestino.ItemData(cboDeductora.ListIndex) & " ¦"
End If



'Lista de Garantias
iCantidad = 0
If chkGarantias.Value = vbUnchecked Then
  vSubTitulo = vSubTitulo & "Garantías: Varias ¦"

    vCadena = " AND ( {vCbrListadoMoraDiasReal.Garantia} in ['"
    For i = 1 To lswGarantias.ListItems.Count
      If lswGarantias.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswGarantias.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i

    If i > 0 Then
        strSQL = strSQL & vCadena & "'])"
    End If
Else
    vSubTitulo = vSubTitulo & "Garantías: Todas ¦"
End If

'Lista de Tipos de Antiguedad
iCantidad = 0
If Mid(lblReporte.Tag, 1, 2) = "08" Or Mid(lblReporte.Tag, 1, 2) = "09" Then 'Antiguedad de Saldos Evalua Todas Siempre
            vSubTitulo = vSubTitulo & "Antiguedad: Todas ¦"
Else
        If chkAntiguedad.Value = vbUnchecked Then
            vSubTitulo = vSubTitulo & "Antiguedad: Varias ¦"
            
            vCadena = " AND ( {vCbrListadoMoraDiasReal.COD_ANTIGUEDAD} in ['"
            For i = 1 To lswAntiguedad.ListItems.Count
              If lswAntiguedad.ListItems.Item(i).Checked Then
                vCadena = vCadena & "','" & lswAntiguedad.ListItems.Item(i).Tag
                iCantidad = iCantidad + 1
              End If
            Next i
        
            If i > 0 Then
                strSQL = strSQL & vCadena & "'])"
            End If
        Else
            vSubTitulo = vSubTitulo & "Antiguedad: Todas ¦"
        End If
End If

'Lista de Tipos de Estado de la Persona
iCantidad = 0
If chkPersona.Value = vbUnchecked Then
    vSubTitulo = vSubTitulo & "Est.Persona: Varias ¦"
    vCadena = " AND ( {vCbrListadoMoraDiasReal.ESTADOACTUAL} in ['"
    For i = 1 To lswEstados.ListItems.Count
      If lswEstados.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswEstados.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i

    If i > 0 Then
        strSQL = strSQL & vCadena & "'])"
    End If
Else
    vSubTitulo = vSubTitulo & "Est.Persona: Todas ¦"
End If


'Lista de Carteras de Cobros
iCantidad = 0
If chkCarteras.Value = vbUnchecked Then
    vSubTitulo = "Carteras: Varias ¦ " & vSubTitulo
    vCadena = " AND ( {CBR_CLASIFICACION_CARTERA.COD_CLASIFICACION} in ['"
    For i = 1 To lswCarteras.ListItems.Count
      If lswCarteras.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswCarteras.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i

    If i > 0 Then
        strSQL = strSQL & vCadena & "'])"
    End If
Else
    vSubTitulo = "Carteras: Todas ¦" & vSubTitulo
End If


With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobro"

    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "fxSubTitulo='" & Mid(vSubTitulo, 1, 250) & "'"
    .Formulas(2) = "fxFecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "fxUsuario='" & glogon.Usuario & "'"
        
  Select Case lblReporte.Tag
   Case "01" 'Listado General
        vReporte = "Cobro_MoraReal_ListadoGeneral"
   
   Case "02" 'Listado por Garantía
        vReporte = "Cobro_MoraReal_ListadoGarantia"
   
   Case "03" 'Listado por Líneas
        vReporte = "Cobro_MoraReal_ListadoLinea"
   
   Case "04" 'Listado por Estado Persona
        vReporte = "Cobro_MoraReal_ListadoPersona"
   
   Case "05" 'Listado por Institución
        vReporte = "Cobro_MoraReal_ListadoInst"
   
   Case "05.1" 'Listado por Deductora
        vReporte = "Cobro_MoraReal_ListadoDeductora"
        
   Case "06" 'Listado por Comité Evaluador
        vReporte = "Cobro_MoraReal_ListadoComite"
        
   Case "07" 'Listado por Provincia
        vReporte = "Cobro_MoraReal_ListadoProvincia"

   Case "08" 'Antiguedad de Saldos
        vReporte = "Cobro_MoraReal_AntiguedadSaldos"

   Case "08.1" 'Antiguedad de Saldos x Garantia
        vReporte = "Cobro_MoraReal_AntiguedadSaldosGarantia"
   
   Case "08.2" 'Antiguedad de Saldos x Comité
        vReporte = "Cobro_MoraReal_AntiguedadSaldosComite"

   Case "09" 'Antiguedad Mora Legal
        vReporte = "Cobro_MoraReal_AntiguedadLegal"
   
   Case "09.1" 'Antiguedad Mora Legal x Garantía
        vReporte = "Cobro_MoraReal_AntiguedadLegalGarantia"
   
   Case "09.2" 'Antiguedad Mora Legal x Comité
        vReporte = "Cobro_MoraReal_AntiguedadLegalComite"

 End Select 'lblReporte.Tag
 
 Select Case cboTipo.Text
   Case "Detalle"
        vReporte = vReporte & ".rpt"
   Case "Resumen"
        vReporte = vReporte & "Rsm.rpt"
   Case "Comparativo"
       If Mid(lblReporte.Tag, 1, 2) = "08" Or Mid(lblReporte.Tag, 1, 2) = "09" Then
            vReporte = vReporte & "Rsm.rpt"
       Else
            vReporte = vReporte & "CPC.rpt"
       End If
 End Select 'cboTipo.Text
 
    .ReportFileName = SIFGlobal.fxPathReportes(vReporte)
    .SelectionFormula = strSQL
    .Action = 1
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub



Private Sub cboInstitucion_Click()
Dim strSQL As String

If vPaso Then Exit Sub

cboDeductora.Clear

If cboInstitucion.Text = "TODOS" Then
    strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
           & " from instituciones order by descripcion"
    Call sbCbo_Llena_New(cboDeductora, strSQL, True, True)
Else
    strSQL = "exec spAFI_Institucion_Vinculadas " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ",3"
    Call sbCbo_Llena_New(cboDeductora, strSQL, True, True)
End If
End Sub

Private Sub cboTipo_Click()
If cboTipo.ListCount <= 0 Then Exit Sub

If cboTipo.Text = "Comparativo" Then
   chkAntiguedad.Value = vbChecked
   Call chkAntiguedad_Click
End If

End Sub

Private Sub chkAntiguedad_Click()
If chkAntiguedad.Value = vbChecked Then
  lswAntiguedad.Enabled = False
Else
  lswAntiguedad.Enabled = True
End If

End Sub

Private Sub chkCarteras_Click()
If chkCarteras.Value = vbChecked Then
  lswCarteras.Enabled = False
Else
  lswCarteras.Enabled = True
End If
End Sub

Private Sub chkGarantias_Click()
If chkGarantias.Value = vbChecked Then
  lswGarantias.Enabled = False
Else
  lswGarantias.Enabled = True
End If
End Sub


Private Sub chkLineas_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If chkLineas.Value = vbChecked Then
  
  txtCodigo.Enabled = False
  
  strSQL = "select rtrim(cod_grupo) as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  catalogo_grupos order by descripcion"
  Call sbCbo_Llena_New(cboRecurso, strSQL, True, True)
  
  strSQL = "select rtrim(cod_destino) as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  catalogo_destinos order by descripcion"
  Call sbCbo_Llena_New(cboDestino, strSQL, True, True)
  
Else
  txtCodigo.Enabled = True

  strSQL = "select (R.cod_grupo) as 'IdX', rtrim(R.descripcion) as 'ItmX'" _
         & " from catalogo_grupos R inner join catalogo_AsignaGrp A on R.cod_grupo = A.cod_grupo" _
         & " where A.codigo = '" & txtCodigo & "' order by R.descripcion"
  Call sbCbo_Llena_New(cboRecurso, strSQL, True, True)
  
  strSQL = "select (R.cod_destino) as 'IdX', rtrim(R.descripcion) as 'ItmX'" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtCodigo & "' order by R.Descripcion"
  Call sbCbo_Llena_New(cboDestino, strSQL, True, True)

End If
End Sub

Private Sub chkPersona_Click()
If chkPersona.Value = vbChecked Then
  lswEstados.Enabled = False
Else
  lswEstados.Enabled = True
End If
End Sub

Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Reporte", 3240
End With

With lswAntiguedad.ColumnHeaders
    .Clear
    .Add , , "Descripción", 6000
End With

With lswGarantias.ColumnHeaders
    .Clear
    .Add , , "Descripción", 6000
End With

With lswEstados.ColumnHeaders
    .Clear
    .Add , , "Descripción", 6000
End With

With lswCarteras.ColumnHeaders
    .Clear
    .Add , , "Descripción", 6000
End With

tcMain.Item(0).Selected = True


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbCubo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim vMensaje As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lblStatus.Caption = "Procesando Información Espere!....Este proceso puede durar varios minutos."
lblStatus.Refresh

vMensaje = "AntiguedadDiasReal"

strSQL = "exec spCbrAntiguedadDiasRealAnalisisCubo "
Call ConectionExecute(strSQL)

lblStatus.Caption = "Proceso Concluido con éxito, la información puede ser utilizada desde la base de datos de análisis, cubo: " & vMensaje

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lsw.ListItems.Count = 0 Then Exit Sub

lblReporte.Caption = Item.Text
lblReporte.Tag = Item.Tag

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
 Case "Reporte"
    lblStatus.Visible = False
    Call sbReporteListados

 Case "Cubos"
    lblStatus.Visible = True
    Call sbCubo
End Select
 
End Sub


Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

TimerX.Interval = 0

lblReporte.Tag = ""
lblReporte.Caption = ">>> Seleccione Un Reporte <<<"

cboTipo.Clear
cboTipo.AddItem "Detalle"
cboTipo.AddItem "Resumen"
cboTipo.AddItem "Comparativo"
cboTipo.Text = "Detalle"

cboEstadoLaboral.Clear
cboEstadoLaboral.AddItem "TODOS"
cboEstadoLaboral.AddItem "Propiedad"
cboEstadoLaboral.AddItem "Interino"
cboEstadoLaboral.Text = "TODOS"

strSQL = " select rtrim(COD_DIVISA) as 'IdX', DESCRIPCION AS 'ItmX'" _
       & " From CNTX_DIVISAS" _
       & " Where COD_CONTABILIDAD = " & GLOBALES.gEnlace
Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)

strSQL = "select rtrim(descripcion) as 'Itmx', id_comite as 'Idx'" _
       & " from comites order by descripcion"
Call sbCbo_Llena_New(cboComite, strSQL, True, True)

strSQL = "select Estado_Laboral as 'IdX', Descripcion as 'ItmX'" _
       & " from AFI_ESTADO_LABORAL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoLaboral, strSQL, True, True)

'Instituciones
vPaso = True
    strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
           & " from instituciones order by descripcion"
    Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)
vPaso = False



With lsw.ListItems
  .Clear
  Set itmX = .Add(, , "Listado General")
      itmX.Tag = "01"
  Set itmX = .Add(, , "Listado por Garantía")
      itmX.Tag = "02"
  Set itmX = .Add(, , "Listado por Líneas")
      itmX.Tag = "03"
  Set itmX = .Add(, , "Listado por Estado Persona")
      itmX.Tag = "04"
  Set itmX = .Add(, , "Listado por Institución")
      itmX.Tag = "05"
  Set itmX = .Add(, , "Listado por Deductora")
      itmX.Tag = "05.1"
  Set itmX = .Add(, , "Listado por Comité Evaluador")
      itmX.Tag = "06"
  Set itmX = .Add(, , "Listado por Provincia")
      itmX.Tag = "07"
  Set itmX = .Add(, , "Antiguedad de Saldos")
      itmX.Tag = "08"
  Set itmX = .Add(, , "Antiguedad de Saldos - Garantía")
      itmX.Tag = "08.1"
  Set itmX = .Add(, , "Antiguedad de Saldos - Comité")
      itmX.Tag = "08.2"
  Set itmX = .Add(, , "Antiguedad Mora Legal")
      itmX.Tag = "09"
  Set itmX = .Add(, , "Antiguedad Mora Legal - Garantía")
      itmX.Tag = "09.1"
  Set itmX = .Add(, , "Antiguedad Mora Legal - Comité")
      itmX.Tag = "09.2"
End With


strSQL = "select garantia as 'Idx',rtrim(descripcion) as 'Descripcion'" _
       & " from crd_garantia_tipos order by descripcion"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswGarantias.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!IdX
  rs.MoveNext
Loop
rs.Close


strSQL = "select cod_Antiguedad as 'Idx',rtrim(descripcion) as 'Descripcion'" _
       & " from CBR_ANTIGUEDAD_TIPOS order by cod_Antiguedad"
rs.Open strSQL, glogon.Conection
Do While Not rs.EOF
  Set itmX = lswAntiguedad.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!IdX
  rs.MoveNext
Loop
rs.Close

'Agrega la de Cobro Judicial
Set itmX = lswAntiguedad.ListItems.Add(, , "9. Cbr.Judicial")
    itmX.Tag = "CBJ"

strSQL = "select cod_clasificacion as 'Idx',rtrim(descripcion) as 'Descripcion'" _
       & " from CBR_CLASIFICACION_CARTERA Where Estado = 1 order by cod_clasificacion"
rs.Open strSQL, glogon.Conection
Do While Not rs.EOF
  Set itmX = lswCarteras.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!IdX
  rs.MoveNext
Loop
rs.Close


strSQL = "select cod_estado as 'Idx',rtrim(descripcion) as 'Descripcion'" _
       & " from AFI_ESTADOS_PERSONA Where ACTIVO = 1 order by cod_estado"
rs.Open strSQL, glogon.Conection
Do While Not rs.EOF
  Set itmX = lswEstados.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!IdX
  rs.MoveNext
Loop
rs.Close


Call chkAntiguedad_Click
Call chkGarantias_Click
Call chkCarteras_Click
Call chkPersona_Click

Call cboInstitucion_Click
Call chkLineas_Click

End Sub




Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then cboDestino.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCodigo_LostFocus()
 If Len(Trim(txtCodigo)) > 0 Then txtDescripcion.Text = fxDescribeCodigo(Trim(txtCodigo))
 Call chkLineas_Click
End Sub

