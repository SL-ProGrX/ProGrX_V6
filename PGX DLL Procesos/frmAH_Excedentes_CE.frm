VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmAH_Excedentes_CE 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cargas Masivas"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   600
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   12495
      _Version        =   1572864
      _ExtentX        =   22040
      _ExtentY        =   10821
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
      ItemCount       =   4
      Item(0).Caption =   "CE Solicitudes"
      Item(0).ControlCount=   23
      Item(0).Control(0)=   "scTitulo(0)"
      Item(0).Control(1)=   "btnBarra(0)"
      Item(0).Control(2)=   "btnBarra(1)"
      Item(0).Control(3)=   "btnBarra(2)"
      Item(0).Control(4)=   "btnBarra(3)"
      Item(0).Control(5)=   "btnBarra(4)"
      Item(0).Control(6)=   "btnBarra(5)"
      Item(0).Control(7)=   "lsw"
      Item(0).Control(8)=   "txtFiltro"
      Item(0).Control(9)=   "btnExport(0)"
      Item(0).Control(10)=   "txtCedula"
      Item(0).Control(11)=   "txtNombre"
      Item(0).Control(12)=   "txtId"
      Item(0).Control(13)=   "FlatEdit2"
      Item(0).Control(14)=   "txtPorcentaje"
      Item(0).Control(15)=   "Label2(1)"
      Item(0).Control(16)=   "Label2(2)"
      Item(0).Control(17)=   "Label2(3)"
      Item(0).Control(18)=   "Label2(4)"
      Item(0).Control(19)=   "txtSalida"
      Item(0).Control(20)=   "Label2(5)"
      Item(0).Control(21)=   "txtDetalle"
      Item(0).Control(22)=   "btnAdjunto"
      Item(1).Caption =   "Carga Masiva"
      Item(1).ControlCount=   17
      Item(1).Control(0)=   "scTitulo(1)"
      Item(1).Control(1)=   "btnExport(1)"
      Item(1).Control(2)=   "btnBuscar"
      Item(1).Control(3)=   "btnCargar"
      Item(1).Control(4)=   "btnInfo"
      Item(1).Control(5)=   "txtArchivo"
      Item(1).Control(6)=   "Label1(2)"
      Item(1).Control(7)=   "btnAplicar"
      Item(1).Control(8)=   "lswCarga"
      Item(1).Control(9)=   "rbCarga(0)"
      Item(1).Control(10)=   "rbCarga(1)"
      Item(1).Control(11)=   "Label3(0)"
      Item(1).Control(12)=   "txtCasos"
      Item(1).Control(13)=   "Label3(1)"
      Item(1).Control(14)=   "txtCasosInco"
      Item(1).Control(15)=   "txtCasosApl"
      Item(1).Control(16)=   "Label3(2)"
      Item(2).Caption =   "Cambio de Salidas"
      Item(2).ControlCount=   22
      Item(2).Control(0)=   "Label2(7)"
      Item(2).Control(1)=   "Label2(8)"
      Item(2).Control(2)=   "Label2(9)"
      Item(2).Control(3)=   "Label2(10)"
      Item(2).Control(4)=   "txtCS_Nombre"
      Item(2).Control(5)=   "txtCS_ID"
      Item(2).Control(6)=   "txtCS_Salida"
      Item(2).Control(7)=   "txtCS_Detalle"
      Item(2).Control(8)=   "lswCS"
      Item(2).Control(9)=   "btnExport(2)"
      Item(2).Control(10)=   "scTitulo(2)"
      Item(2).Control(11)=   "Label2(11)"
      Item(2).Control(12)=   "btnCS(0)"
      Item(2).Control(13)=   "btnCS(1)"
      Item(2).Control(14)=   "btnCS(2)"
      Item(2).Control(15)=   "btnCS(3)"
      Item(2).Control(16)=   "txtCS_Filtro"
      Item(2).Control(17)=   "rbCS_Autorizado(0)"
      Item(2).Control(18)=   "rbCS_Autorizado(1)"
      Item(2).Control(19)=   "rbCS_Autorizado(2)"
      Item(2).Control(20)=   "txtCS_Autoriza"
      Item(2).Control(21)=   "txtCS_Identificacion"
      Item(3).Caption =   "CE Consultas"
      Item(3).ControlCount=   12
      Item(3).Control(0)=   "Label2(12)"
      Item(3).Control(1)=   "Label2(13)"
      Item(3).Control(2)=   "Label2(14)"
      Item(3).Control(3)=   "Label2(15)"
      Item(3).Control(4)=   "txtCE_Cedula"
      Item(3).Control(5)=   "txtCE_Nombre"
      Item(3).Control(6)=   "txtCE_Detalle"
      Item(3).Control(7)=   "txtCE_Salida"
      Item(3).Control(8)=   "btnCE(0)"
      Item(3).Control(9)=   "lswCE"
      Item(3).Control(10)=   "btnExport(3)"
      Item(3).Control(11)=   "scTitulo(3)"
      Begin XtremeSuiteControls.ListView lswCE 
         Height          =   4335
         Left            =   -69880
         TabIndex        =   72
         Top             =   1680
         Visible         =   0   'False
         Width           =   12255
         _Version        =   1572864
         _ExtentX        =   21616
         _ExtentY        =   7646
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
         Appearance      =   20
      End
      Begin XtremeSuiteControls.ListView lswCarga 
         Height          =   3375
         Left            =   -69880
         TabIndex        =   31
         Top             =   1920
         Visible         =   0   'False
         Width           =   12255
         _Version        =   1572864
         _ExtentX        =   21616
         _ExtentY        =   5953
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
         Appearance      =   20
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3135
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   12255
         _Version        =   1572864
         _ExtentX        =   21616
         _ExtentY        =   5530
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
         Appearance      =   20
      End
      Begin XtremeSuiteControls.ListView lswCS 
         Height          =   3015
         Left            =   -69880
         TabIndex        =   48
         Top             =   3000
         Visible         =   0   'False
         Width           =   12255
         _Version        =   1572864
         _ExtentX        =   21616
         _ExtentY        =   5318
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
         Appearance      =   20
      End
      Begin XtremeSuiteControls.FlatEdit txtCasosInco 
         Height          =   330
         Left            =   -62560
         TabIndex        =   79
         Top             =   5520
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCE_Nombre 
         Height          =   315
         Left            =   -67960
         TabIndex        =   64
         Top             =   840
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
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
      Begin XtremeSuiteControls.RadioButton rbCarga 
         Height          =   375
         Index           =   0
         Left            =   -67840
         TabIndex        =   61
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Carga Casos Especiales"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbCS_Autorizado 
         Height          =   255
         Index           =   0
         Left            =   -65920
         TabIndex        =   58
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todos "
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
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         ToolTipText     =   "Nuevo"
         Top             =   360
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Nuevo"
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
         Picture         =   "frmAH_Excedentes_CE.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   1
         Left            =   2760
         TabIndex        =   7
         ToolTipText     =   "Editar"
         Top             =   360
         Width           =   375
         _Version        =   1572864
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":0632
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         ToolTipText     =   "Eliminar"
         Top             =   360
         Width           =   375
         _Version        =   1572864
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":0C2D
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   3
         Left            =   3720
         TabIndex        =   9
         ToolTipText     =   "Guardar"
         Top             =   360
         Width           =   375
         _Version        =   1572864
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":11D1
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   4
         Left            =   4080
         TabIndex        =   10
         ToolTipText     =   "Deshacer"
         Top             =   360
         Width           =   375
         _Version        =   1572864
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":1902
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   5
         Left            =   4560
         TabIndex        =   11
         ToolTipText     =   "Reporte"
         Top             =   360
         Width           =   375
         _Version        =   1572864
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":2002
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Left            =   4080
         TabIndex        =   13
         Top             =   2560
         Width           =   6615
         _Version        =   1572864
         _ExtentX        =   11668
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
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   255
         Index           =   0
         Left            =   11760
         TabIndex        =   14
         ToolTipText     =   "Exportar Listado a Excel"
         Top             =   2560
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmAH_Excedentes_CE.frx":2709
      End
      Begin XtremeSuiteControls.PushButton btnAdjunto 
         Height          =   330
         Left            =   5280
         TabIndex        =   16
         ToolTipText     =   "Adjuntar documentos"
         Top             =   360
         Width           =   375
         _Version        =   1572864
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":2873
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   1080
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Left            =   3480
         TabIndex        =   18
         Top             =   1080
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
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
      Begin XtremeSuiteControls.FlatEdit txtId 
         Height          =   555
         Left            =   9000
         TabIndex        =   19
         Top             =   480
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   979
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
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
      Begin XtremeSuiteControls.FlatEdit FlatEdit2 
         Height          =   315
         Left            =   3000
         TabIndex        =   20
         Top             =   2040
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
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
         Text            =   "%"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   2040
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSalida 
         Height          =   315
         Left            =   4440
         TabIndex        =   27
         ToolTipText     =   "Presione F4 para Consultar Salidas Disponibles"
         Top             =   2040
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   555
         Left            =   1680
         TabIndex        =   29
         Top             =   1440
         Width           =   9135
         _Version        =   1572864
         _ExtentX        =   16113
         _ExtentY        =   979
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
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   255
         Index           =   1
         Left            =   -58240
         TabIndex        =   32
         ToolTipText     =   "Exportar Listado a Excel"
         Top             =   1605
         Visible         =   0   'False
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmAH_Excedentes_CE.frx":28FC
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   375
         Left            =   -60880
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":2A66
      End
      Begin XtremeSuiteControls.PushButton btnCargar 
         Height          =   375
         Left            =   -60400
         TabIndex        =   34
         Top             =   960
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":3166
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   375
         Left            =   -59920
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":387F
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   495
         Left            =   -67840
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   6855
         _Version        =   1572864
         _ExtentX        =   12091
         _ExtentY        =   873
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   495
         Left            =   -58960
         TabIndex        =   5
         Top             =   5400
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":3F98
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtCS_Identificacion 
         Height          =   315
         Left            =   -68320
         TabIndex        =   39
         ToolTipText     =   "Presione F4"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCS_Nombre 
         Height          =   315
         Left            =   -66520
         TabIndex        =   40
         Top             =   1080
         Visible         =   0   'False
         Width           =   7335
         _Version        =   1572864
         _ExtentX        =   12938
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
      Begin XtremeSuiteControls.FlatEdit txtCS_ID 
         Height          =   555
         Left            =   -68320
         TabIndex        =   41
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   979
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
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
      Begin XtremeSuiteControls.FlatEdit txtCS_Salida 
         Height          =   315
         Left            =   -68320
         TabIndex        =   42
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCS_Detalle 
         Height          =   555
         Left            =   -68320
         TabIndex        =   43
         Top             =   1440
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1572864
         _ExtentX        =   16113
         _ExtentY        =   979
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
      Begin XtremeSuiteControls.FlatEdit txtCS_Filtro 
         Height          =   315
         Left            =   -65920
         TabIndex        =   49
         Top             =   2685
         Visible         =   0   'False
         Width           =   6615
         _Version        =   1572864
         _ExtentX        =   11668
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
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   255
         Index           =   2
         Left            =   -58240
         TabIndex        =   50
         ToolTipText     =   "Exportar Listado a Excel"
         Top             =   2685
         Visible         =   0   'False
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmAH_Excedentes_CE.frx":46BF
      End
      Begin XtremeSuiteControls.FlatEdit txtCS_Autoriza 
         Height          =   315
         Left            =   -61000
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCS 
         Height          =   330
         Index           =   0
         Left            =   -66400
         TabIndex        =   54
         ToolTipText     =   "Nuevo"
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Nuevo"
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
         Picture         =   "frmAH_Excedentes_CE.frx":4829
      End
      Begin XtremeSuiteControls.PushButton btnCS 
         Height          =   330
         Index           =   1
         Left            =   -65320
         TabIndex        =   55
         ToolTipText     =   "Eliminar"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
         _Version        =   1572864
         _ExtentX        =   661
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":4E5B
      End
      Begin XtremeSuiteControls.PushButton btnCS 
         Height          =   330
         Index           =   2
         Left            =   -64840
         TabIndex        =   56
         ToolTipText     =   "Registrar"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
         _Version        =   1572864
         _ExtentX        =   661
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":53FF
      End
      Begin XtremeSuiteControls.PushButton btnCS 
         Height          =   330
         Index           =   3
         Left            =   -59080
         TabIndex        =   57
         ToolTipText     =   "Autorizar"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
         _Version        =   1572864
         _ExtentX        =   661
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAH_Excedentes_CE.frx":5B30
      End
      Begin XtremeSuiteControls.RadioButton rbCS_Autorizado 
         Height          =   255
         Index           =   1
         Left            =   -64480
         TabIndex        =   59
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Pendientes"
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
      Begin XtremeSuiteControls.RadioButton rbCS_Autorizado 
         Height          =   255
         Index           =   2
         Left            =   -62920
         TabIndex        =   60
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Autorizados"
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
      Begin XtremeSuiteControls.RadioButton rbCarga 
         Height          =   375
         Index           =   1
         Left            =   -65080
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Carga Cambios de Salida"
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
      Begin XtremeSuiteControls.FlatEdit txtCE_Cedula 
         Height          =   315
         Left            =   -69760
         TabIndex        =   63
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCE_Detalle 
         Height          =   315
         Left            =   -64840
         TabIndex        =   67
         Top             =   840
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
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
      Begin XtremeSuiteControls.FlatEdit txtCE_Salida 
         Height          =   315
         Left            =   -61720
         TabIndex        =   69
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCE 
         Height          =   330
         Index           =   0
         Left            =   -60040
         TabIndex        =   71
         ToolTipText     =   "Autorizar"
         Top             =   840
         Visible         =   0   'False
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Picture         =   "frmAH_Excedentes_CE.frx":6257
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   255
         Index           =   3
         Left            =   -58240
         TabIndex        =   73
         ToolTipText     =   "Exportar Listado a Excel"
         Top             =   1365
         Visible         =   0   'False
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmAH_Excedentes_CE.frx":6957
      End
      Begin XtremeSuiteControls.FlatEdit txtCasos 
         Height          =   330
         Left            =   -68920
         TabIndex        =   77
         Top             =   5520
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCasosApl 
         Height          =   330
         Left            =   -65680
         TabIndex        =   80
         Top             =   5520
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   2
         Left            =   -67600
         TabIndex        =   81
         Top             =   5520
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Casos Correctos: "
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   1
         Left            =   -64120
         TabIndex        =   78
         Top             =   5520
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Inconsistencias: "
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   0
         Left            =   -69760
         TabIndex        =   76
         Top             =   5520
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Casos: "
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
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Index           =   3
         Left            =   -69880
         TabIndex        =   74
         Top             =   1320
         Visible         =   0   'False
         Width           =   12255
         _Version        =   1572864
         _ExtentX        =   21616
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Lista de Casos Especiales y su aplicación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   15
         Left            =   -61720
         TabIndex        =   70
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Salida"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   14
         Left            =   -64840
         TabIndex        =   68
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detalle"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   13
         Left            =   -67960
         TabIndex        =   66
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   12
         Left            =   -69760
         TabIndex        =   65
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificacion"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   11
         Left            =   -62320
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Autorizado ?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Index           =   2
         Left            =   -69880
         TabIndex        =   51
         Top             =   2640
         Visible         =   0   'False
         Width           =   12255
         _Version        =   1572864
         _ExtentX        =   21616
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Listado con Cambios de Salida                 Filtros:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   10
         Left            =   -69640
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Id"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   -69640
         TabIndex        =   46
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificacion"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   -69640
         TabIndex        =   45
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nueva Salida"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   7
         Left            =   -69640
         TabIndex        =   44
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detalle"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   -69400
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Index           =   1
         Left            =   -69880
         TabIndex        =   30
         Top             =   1560
         Visible         =   0   'False
         Width           =   12255
         _Version        =   1572864
         _ExtentX        =   21616
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Casos Cargados :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   28
         Top             =   1560
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detalle"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   3720
         TabIndex        =   26
         Top             =   2040
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Salida"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   25
         Top             =   2040
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Porcentaje"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificacion"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   7920
         TabIndex        =   23
         Top             =   600
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Id"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   12255
         _Version        =   1572864
         _ExtentX        =   21616
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Listado de Casos Especiales                 Filtros:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   3735
      _Version        =   1572864
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
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   12495
      _Version        =   1572864
      _ExtentX        =   22040
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   315
      Left            =   11760
      TabIndex        =   37
      ToolTipText     =   "Cantidad de Casos a Mostrar"
      Top             =   8040
      Width           =   855
      _Version        =   1572864
      _ExtentX        =   1508
      _ExtentY        =   556
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
      Text            =   "100"
      BackColor       =   16777215
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label lblPermiteCambios 
      Height          =   255
      Left            =   5280
      TabIndex        =   75
      Top             =   1320
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Permite Cambios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   6
      Left            =   11040
      TabIndex        =   38
      Top             =   8040
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Lineas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Periodo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Excedentes: Gestión de Casos Especiales"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   11
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   9252
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmAH_Excedentes_CE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean
Dim itmX As ListViewItem
Dim vEdita As Boolean, vFecha As Date


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


Private Sub btnAdjunto_Click()
 
If txtId.Text <> "" Then
 gGA.Modulo = "EX_01"
 gGA.Llave_01 = txtCedula.Text
 gGA.Llave_02 = cboPeriodo.ItemData(cboPeriodo.ListIndex)
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)

Else
 MsgBox "Ingrese el Caso Primero y luego suba los adjuntos requeridos!", vbInformation
End If

End Sub

Private Sub btnAplicar_Click()
Dim i As Integer, pPeriodoId As Long

On Error GoTo vError

If CLng(txtCasosInco.Text) > 0 Then
    i = MsgBox("Se encuentra Casos Inconsistentes, de continuar el proceso solo se aplicarán los casos correctos, desea continuar?", vbYesNo)
    If i = vbNo Then
        Exit Sub
    End If
End If

If CLng(txtCasos.Text) = 0 Or CLng(txtCasosApl.Text) = 0 Then
    MsgBox "No Existen Casos a Procesar!", vbExclamation
    Exit Sub
End If

pPeriodoId = cboPeriodo.ItemData(cboPeriodo.ListIndex)

i = MsgBox("Está Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then

    Me.MousePointer = vbHourglass


    Select Case True
     Case rbCarga(0).Value 'Casos Especiales
         strSQL = "exec spEXC_Mass_CE_Procesa " & pPeriodoId & ", 'CE'"
    
     Case rbCarga(1).Value 'Cambios de Salidas
         strSQL = "exec spEXC_Mass_CS_Procesa " & pPeriodoId & ", 'CS'"
    
    End Select

    Call ConectionExecute(strSQL)

End If

MsgBox "Casos Procesados Satisfactoriamente!", vbInformation
Call sbMass_Limpia

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBarra_Click(Index As Integer)

Select Case Index
    Case 0 'NUEVO
        vEdita = False
        Call sbLimpia
        txtCedula.SetFocus

        Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
      If txtId.Text = "" Then
        MsgBox "Seleccione una caso de la lista del activo para modificacion...", vbInformation
      Else
        vEdita = True
        txtPorcentaje.SetFocus
        Call sbBarra_Accion("Editar")
      End If
      
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
      Call sbBarra_Accion("Editar")
      If txtId.Text = "" Then
        Call sbLimpia
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      End If
    
    Case 5 'REPORTES
   
End Select

End Sub

Private Function fxValida() As Boolean

Dim vMensaje As String

vMensaje = ""
fxValida = True

If lblPermiteCambios.Tag = "0" Then
  vMensaje = vMensaje & "- El Periodo Consultado, no permite cambios!" & vbCrLf
End If

strSQL = "select Estado from Exc_Periodos where Id_Periodo = " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
Call OpenRecordSet(rs, strSQL)
If rs!ESTADO = "C" Then
  vMensaje = vMensaje & "- No se pueden registrar casos especiales en un Periodo Cerrado!" & vbCrLf
End If
rs.Close

If Trim(txtCedula) = "" Then
  vMensaje = vMensaje & "- El No. de Cédula no es válido!" & vbCrLf
End If

If Trim(txtSalida) = "" Then
  vMensaje = vMensaje & "- No se ha indicado una salida válida!" & vbCrLf
End If

If Trim(txtDetalle) = "" Then
  vMensaje = vMensaje & "- El Detalle no es válido!" & vbCrLf
End If

If Not IsNumeric(txtPorcentaje.Text) Then
  vMensaje = vMensaje & "- El Porcentaje no es valido!" & vbCrLf
End If

If IsNumeric(txtPorcentaje.Text) Then
    If CCur(txtPorcentaje.Text) < 0 Or CCur(txtPorcentaje.Text) > 100 Then
      vMensaje = vMensaje & "- El Porcentaje no es valido!" & vbCrLf
    End If
End If


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Function fxCS_Valida() As Boolean

Dim vMensaje As String

vMensaje = ""
fxCS_Valida = True

If lblPermiteCambios.Tag = "0" Then
  vMensaje = vMensaje & "- El Periodo Consultado, no permite cambios!" & vbCrLf
End If

strSQL = "select Estado from Exc_Periodos where Id_Periodo = " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
Call OpenRecordSet(rs, strSQL)
If rs!ESTADO <> "C" Then
  vMensaje = vMensaje & "- No se pueden registrar Cambios de Salida en un Periodo Abierto!" & vbCrLf
End If
rs.Close


If Trim(txtCS_Identificacion.Text) = "" Then
  vMensaje = vMensaje & "- El No. de Cédula no es válido!" & vbCrLf
End If

If Trim(txtCS_Salida.Text) = "" Then
  vMensaje = vMensaje & "- No se ha indicado una salida válida!" & vbCrLf
End If

If Trim(txtCS_Detalle.Text) = "" Then
  vMensaje = vMensaje & "- El Detalle no es válido!" & vbCrLf
End If


If Len(vMensaje) > 0 Then
  fxCS_Valida = False
  MsgBox vMensaje, vbCritical
End If

End Function



Private Sub sbBorrar()
Dim pId As Long

Dim i As Integer

On Error GoTo vError

If Trim(txtId.Text) = "" Then Exit Sub

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
       
    pId = txtId.Text
    txtCedula.Text = fxSysCleanTxtInject(txtCedula.Text)
 
    strSQL = "exec spExc_Caso_Especial_Delete " & pId & ", " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & ", '" & txtCedula.Text & "', '" & glogon.Usuario & "'"
    
     Call OpenRecordSet(rs, strSQL)
     If rs!Pass = 1 Then
         Call Bitacora("Borra", "Excedente Casos Especial Id:  " & txtId.Text & ", Cedula: " & txtCedula.Text & ", Porcentaje: " & txtPorcentaje.Text & ", Periodo: " & cboPeriodo.Text)
         MsgBox "Caso Eliminado Satisfactoriamente!", vbInformation
     
        Call sbBarra_Accion("NUEVO")
        Call sbListado_Load
        Call RefrescaTags(Me)
     
     Else
         MsgBox "No se puede Eliminar el Caso por que ya fué procesado o no Existe!", vbExclamation
     End If
     
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbLimpia()

txtId.Text = ""
txtCedula.Text = ""
txtNombre.Text = ""
txtPorcentaje.Text = "0"
     
txtSalida.Text = ""
txtDetalle.Text = ""
     
End Sub

Private Sub sbListado_Load()

On Error GoTo vError

Dim pFiltro As String

pFiltro = fxSysCleanTxtInject(txtFiltro.Text)
txtLineas.Text = fxSysCleanTxtInject(txtLineas.Text)

If Not IsNumeric(txtLineas.Text) Then
    txtLineas.Text = "100"
End If

strSQL = "Select Top " & txtLineas.Text & " A.*, S.Nombre" _
       & " from EXC_CASOS_ESPECIALES A inner join Socios S on A.cedula = S.cedula" _
       & " where A.ID_Periodo = " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
       & "   and (S.cedula like '%" & pFiltro & "%'" _
       & "     or S.Nombre like '%" & pFiltro & "%')"
lsw.ListItems.Clear

vPaso = True

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!consec)
       itmX.SubItems(1) = Trim(rs!Cedula)
       itmX.SubItems(2) = rs!Nombre
       itmX.SubItems(3) = Format(rs!Porcentaje, "Standard")
       itmX.SubItems(4) = Trim(rs!Salida)
       itmX.SubItems(5) = rs!Detalle & ""
   rs.MoveNext
Loop
rs.Close

vPaso = False

'txtCasos.Text = Format(lsw.ListItems.Count, "###,##0")

Call sbLimpia

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub sbCedula_Load()

strSQL = "Select Top 1 A.Cedula, S.Nombre, A.Salida, A.detalle, case when A.Doc_Ajunto is null then 0 else 1 end as 'Adjunto'" _
       & " from EXC_CASOS_ESPECIALES A inner join Socios S on A.cedula = S.cedula" _
       & " where A.ID_PERIODO = " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & " and A.cedula = '" & Trim(txtCedula.Text) _
       & "' order by A.CONSEC desc"

Call OpenRecordSet(rs, strSQL)


If Not rs.EOF And Not rs.BOF Then
   vEdita = True

   txtId.Text = rs!consec

   txtPorcentaje.Text = Format(rs!Porcentaje, "Standard")
   txtSalida.Text = rs!Salida & ""
   txtDetalle.Text = rs!Detalle & ""
   
   Call sbBarra_Accion("EDITAR")

Else
   txtId.Text = ""
   txtNombre.Text = fxNombre(txtCedula.Text)
End If

rs.Close

End Sub

Private Sub sbGuardar()
Dim pId As Long

On Error GoTo vError


If txtId.Text = "" Then
   pId = 0
Else
   pId = txtId.Text
End If
txtCedula.Text = fxSysCleanTxtInject(txtCedula.Text)
txtDetalle.Text = fxSysCleanTxtInject(txtDetalle.Text)

strSQL = "exec spExc_Caso_Especial_Add " & pId & ", " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & ", '" & txtCedula.Text & "', '" & txtDetalle.Text _
       & "', " & CCur(txtPorcentaje.Text) & ", '" & txtSalida.Text & "', '" & glogon.Usuario & "'"

 Call OpenRecordSet(rs, strSQL)
 If rs!Pass = 1 Then
     Call Bitacora("Registra", "Excedente Casos Especial Id:  " & txtId.Text & ", Cedula: " & txtCedula.Text & ", Porcentaje: " & txtPorcentaje.Text & ", Periodo: " & cboPeriodo.Text)
     MsgBox "Caso Registrado Satisfactoriamente!", vbInformation
 
    Call sbBarra_Accion("NUEVO")
    Call sbListado_Load
 
 Else
     MsgBox "No se puede registrar el Caso por que ya fué procesado!", vbExclamation
 End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCE_Lista()

On Error GoTo vError

Me.MousePointer = vbHourglass


lswCE.ListItems.Clear
With lswCE.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Periodo Id", 1200, vbCenter
    .Add , , "Cedula", 1400, vbCenter
    .Add , , "Nombre", 3500
    .Add , , "Salida", 1100, vbCenter
    .Add , , "Detalle", 3000
    .Add , , "Adjunto ?", 1400, vbCenter
    .Add , , "Aplica Id", 1800, vbCenter
    .Add , , "Porcentaje", 1100, vbRightJustify
    .Add , , "Reg.Fecha", 1800
    .Add , , "Reg.Usuario", 1500, vbCenter
    .Add , , "Mod.Fecha", 1800
    .Add , , "Mod.Usuario", 1500, vbCenter
End With

strSQL = "exec spExc_Casos_Especiales_Aplicados " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & ", '" & txtCE_Salida.Text & "', '" & txtCE_Cedula.Text _
       & "', '" & txtCE_Nombre.Text & "', '" & txtCE_Detalle.Text & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    Set itmX = lswCE.ListItems.Add(, , rs!consec)
        itmX.SubItems(1) = rs!ID_PERIODO
        itmX.SubItems(2) = rs!Cedula
        itmX.SubItems(3) = rs!Nombre
        itmX.SubItems(4) = rs!Salida
        itmX.SubItems(5) = rs!Detalle
        itmX.SubItems(6) = rs!DOC_ADJUNTO
        itmX.SubItems(7) = rs!Consec_Apl
        itmX.SubItems(8) = rs!Porcentaje & ""
        itmX.SubItems(9) = rs!Registro_Fecha & ""
        itmX.SubItems(10) = rs!Registro_Usuario & ""
        itmX.SubItems(11) = rs!Modifica_Fecha & ""
        itmX.SubItems(12) = rs!Modifica_Usuario & ""
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBuscar_Click()

txtArchivo.Text = ""

With frmContenedor.CD
        
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]..."
    .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
    .ShowOpen

    If .FileName = "" Then
        MsgBox "Archivo no válido...", vbExclamation
        Exit Sub
    End If

    If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
        'Ok
    Else
        MsgBox "La Extensión del Archivo no es válido...", vbExclamation
        Exit Sub
    End If
    
    txtArchivo.Text = .FileName

End With

End Sub

Private Sub sbCE_Mass_Carga()
Dim pCedula As String, pNombre As String, pPorcentaje As Currency
Dim pPeriodoId As Long, pSalida As String, pDetalle As String

Dim pLinea As Long, pArchivo As String

Dim i As Integer, vCampos As Boolean

On Error GoTo vError

lswCarga.ListItems.Clear

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If


pPeriodoId = cboPeriodo.ItemData(cboPeriodo.ListIndex)

txtCasosApl.Text = 0
txtCasosInco.Text = 0
txtCasos.Text = 0


pArchivo = Dir(txtArchivo.Text, vbArchive)

Set rs = Excel_Load(txtArchivo.Text, "IMPORT")
    
'Validaciónn del Archivo
vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "CEDULA" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, PORCENTAJE, DETALLE ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "NOMBRE" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, PORCENTAJE, DETALLE ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "SALIDA" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, PORCENTAJE, DETALLE¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If

vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "PORCENTAJE" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, PORCENTAJE, DETALLE¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If

vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "DETALLE" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, PORCENTAJE, DETALLE  ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If
'FIN: Validación del Archivo


ProgressBarX.Visible = True
ProgressBarX.Caption = "Subiendo Archivo, Espere!"
DoEvents

Me.MousePointer = vbHourglass


'Sube, Revisa y Carga
With lswCarga.ListItems
    .Clear
    
    pLinea = 0
    strSQL = ""
    
    Do While Not rs.EOF
      If Trim(rs!Cedula) <> "" Then
        pCedula = rs!Cedula
        pNombre = rs!Nombre & ""
        pPorcentaje = rs!Porcentaje
        pSalida = rs!Salida & ""
        pDetalle = rs!Detalle & ""
        
        pLinea = pLinea + 1
        
        If pLinea = 1 Then
            strSQL = strSQL & Space(10) & "exec spEXC_Mass_CE_Sube " & pPeriodoId & ", 'CE', '" & pCedula & "', '" & pNombre _
                   & "', '" & pSalida & "', " & pPorcentaje & ",  '" & pDetalle & "', '" & glogon.Usuario & "', 1"
            Call ConectionExecute(strSQL)
            strSQL = ""
        Else
            strSQL = strSQL & Space(10) & "exec spEXC_Mass_CE_Sube " & pPeriodoId & ", 'CE', '" & pCedula & "', '" & pNombre _
                   & "', '" & pSalida & "', " & pPorcentaje & ",  '" & pDetalle & "', '" & glogon.Usuario & "', 0"
        End If
        
        If Len(strSQL) > 40000 Then
           Call ConectionExecute(strSQL)
           If glogon.error Then
              Exit Sub
           End If
           
           ProgressBarX.Caption = "Subiendo Archivo, Registros Procesados:  " & pLinea & ", Espere!"
           DoEvents
           strSQL = ""
        End If
        
      End If
      rs.MoveNext
    Loop
    rs.Close

'Procesa Ultimo Bloque

If Len(strSQL) > 0 Then
   ProgressBarX.Caption = "Subiendo Archivo, Registros Procesados:  " & pLinea & ", Espere!"
   DoEvents
   Call ConectionExecute(strSQL)
   If glogon.error Then
      Exit Sub
   End If
   strSQL = ""
End If



ProgressBarX.Caption = "Revisando Registros e Inconsistencias"
DoEvents

Me.MousePointer = vbHourglass

'Revisa Lote y lo Carga
strSQL = "exec spEXC_Mass_CE_Valida " & pPeriodoId & ", 'CE'"
Call OpenRecordSet(rs, strSQL)

    txtCasos.Text = Format(rs!Total, "###,##0")
    txtCasosApl.Text = Format(rs!Aplica, "###,##0")
    txtCasosInco.Text = Format(rs!Inco, "###,##0")
rs.Close
                   
strSQL = "exec spEXC_Mass_CE_Consulta " & pPeriodoId & ", 'CE'"
Call OpenRecordSet(rs, strSQL)
If glogon.error Then
   Exit Sub
End If

    Do While Not rs.EOF
       Set itmX = .Add(, , rs!Cedula)
           itmX.SubItems(1) = rs!Nombre
           itmX.SubItems(2) = rs!Porcentaje
           itmX.SubItems(3) = rs!Salida
           itmX.SubItems(4) = rs!Detalle
           itmX.SubItems(5) = rs!Inconsistencia
           
           If rs!Aplica = 0 Then
            itmX.TextBackColor = RGB(250, 219, 216)
           End If

      rs.MoveNext
    Loop
    rs.Close


End With 'LswCarga


Me.MousePointer = vbDefault
ProgressBarX.Visible = False

MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    ProgressBarX.Visible = False
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbMass_Limpia
End Sub



Private Sub sbCS_Mass_Carga()
Dim pCedula As String, pNombre As String
Dim pPeriodoId As Long, pSalida As String, pDetalle As String
Dim pAutorizaInd As Integer, pAutorizaUsuario As String

Dim pLinea As Long, pArchivo As String

Dim i As Integer, vCampos As Boolean


On Error GoTo vError

lswCarga.ListItems.Clear

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If


pPeriodoId = cboPeriodo.ItemData(cboPeriodo.ListIndex)

txtCasosApl.Text = 0
txtCasosInco.Text = 0
txtCasos.Text = 0


pArchivo = Dir(txtArchivo.Text, vbArchive)

Set rs = Excel_Load(txtArchivo.Text, "IMPORT")
    
'Validaciónn del Archivo
vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "CEDULA" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, DETALLE, AUTORIZA_IND, AUTORIZA_USUARIO¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "NOMBRE" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, DETALLE, AUTORIZA_IND, AUTORIZA_USUARIO ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "SALIDA" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, DETALLE, AUTORIZA_IND, AUTORIZA_USUARIO ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If

vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "AUTORIZA_IND" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, DETALLE, AUTORIZA_IND, AUTORIZA_USUARIO ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "AUTORIZA_USUARIO" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, DETALLE, AUTORIZA_IND, AUTORIZA_USUARIO ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "DETALLE" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "Los campos son CEDULA, NOMBRE, SALIDA, DETALLE, AUTORIZA_IND, AUTORIZA_USUARIO ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If
'FIN: Validación del Archivo


ProgressBarX.Visible = True
ProgressBarX.Caption = "Subiendo Archivo, Espere!"
DoEvents

Me.MousePointer = vbHourglass


'Sube, Revisa y Carga
With lswCarga.ListItems
    
    pLinea = 0
    strSQL = ""
    
    Do While Not rs.EOF
      If Trim(rs!Cedula) <> "" Then
        pCedula = rs!Cedula
        pNombre = rs!Nombre & ""
        pSalida = rs!Salida & ""
        pDetalle = rs!Detalle & ""
        pAutorizaInd = IIf(IsNull(rs!Autoriza_Ind), 0, rs!Autoriza_Ind)
        pAutorizaUsuario = rs!Autoriza_Usuario & ""
                
        pLinea = pLinea + 1
        
        If pLinea = 1 Then
            strSQL = strSQL & Space(10) & "exec spEXC_Mass_CS_Sube " & pPeriodoId & ", 'CS', '" & pCedula & "', '" & pNombre _
                   & "', '" & pSalida & "', '" & pDetalle & "', " & pAutorizaInd & ", '" & pAutorizaUsuario & "', '" & glogon.Usuario & "', 1"
            Call ConectionExecute(strSQL)
            strSQL = ""
        Else
            strSQL = strSQL & Space(10) & "exec spEXC_Mass_CS_Sube " & pPeriodoId & ", 'CS', '" & pCedula & "', '" & pNombre _
                   & "', '" & pSalida & "', '" & pDetalle & "', " & pAutorizaInd & ", '" & pAutorizaUsuario & "', '" & glogon.Usuario & "', 0"
        End If
        
        If Len(strSQL) > 40000 Then
           Call ConectionExecute(strSQL)
           If glogon.error Then
              Exit Sub
           End If
           
           ProgressBarX.Caption = "Subiendo Archivo, Registros Procesados:  " & pLinea & ", Espere!"
           DoEvents
           strSQL = ""
        End If
        
      End If
      rs.MoveNext
    Loop
    rs.Close

'Procesa Ultimo Bloque

If Len(strSQL) > 0 Then
   ProgressBarX.Caption = "Subiendo Archivo, Registros Procesados:  " & pLinea & ", Espere!"
   DoEvents
   Call ConectionExecute(strSQL)
   If glogon.error Then
      Exit Sub
   End If
   strSQL = ""
End If



ProgressBarX.Caption = "Revisando Registros e Inconsistencias"
DoEvents

Me.MousePointer = vbHourglass

'Revisa Lote y lo Carga
strSQL = "exec spEXC_Mass_CS_Valida " & pPeriodoId & ", 'CS'"
Call OpenRecordSet(rs, strSQL)
    txtCasos.Text = Format(rs!Total, "###,##0")
    txtCasosApl.Text = Format(rs!Aplica, "###,##0")
    txtCasosInco.Text = Format(rs!Inco, "###,##0")
rs.Close
                   
strSQL = "exec spEXC_Mass_CS_Consulta " & pPeriodoId & ", 'CS'"
Call OpenRecordSet(rs, strSQL)
If glogon.error Then
   Call sbMass_Limpia
   Exit Sub
End If
    
    .Clear
    Do While Not rs.EOF
       Set itmX = .Add(, , rs!Cedula)
           itmX.SubItems(1) = rs!Nombre
           itmX.SubItems(2) = rs!Salida
           itmX.SubItems(3) = rs!Detalle
           itmX.SubItems(4) = rs!Autoriza_Ind
           itmX.SubItems(5) = rs!Autoriza_Usuario
           
           itmX.SubItems(6) = rs!Inconsistencia
           If rs!Aplica = 0 Then
                itmX.TextBackColor = RGB(250, 219, 216)
           End If
      rs.MoveNext
    Loop
    rs.Close


End With 'LswCarga


Me.MousePointer = vbDefault
ProgressBarX.Visible = False

MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    ProgressBarX.Visible = False
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbMass_Limpia
End Sub


Private Sub btnCargar_Click()
Select Case True
 Case rbCarga(0).Value 'Casos Especiales
     Call sbCE_Mass_Carga
 
 Case rbCarga(1).Value 'Cambios de Salidas
     Call sbCS_Mass_Carga
End Select
End Sub

Private Sub btnCE_Click(Index As Integer)

Select Case Index
    Case 0
        Call sbCE_Lista
    Case 1
End Select

End Sub

Private Sub btnCS_Click(Index As Integer)

If Index <> 0 Then
 If Not fxCS_Valida Then
    Exit Sub
 End If
End If

Select Case Index
    Case 0 'Nuevo
        Call sbCS_Nuevo
    Case 1 'Eliminar
        Call sbCS_Elimina
    Case 2 'Guardar
        Call sbCS_Guarda
    Case 3 'Autorizar
        Call sbCS_Autoriza
End Select

End Sub

Private Sub btnExport_Click(Index As Integer)
 Select Case Index
    Case 0 'CE Solicitudos
        Call Excel_Exportar_Lsw(lsw, ProgressBarX)
    
    Case 1 'Carga
        Call Excel_Exportar_Lsw(lswCarga, ProgressBarX)
    
    Case 2 'CS
        Call Excel_Exportar_Lsw(lswCS, ProgressBarX)

    Case 3 'CE
        Call Excel_Exportar_Lsw(lswCE, ProgressBarX)

 End Select
End Sub

Private Sub sbPermite_Cambios()

On Error GoTo vError

strSQL = "select dbo.fxExc_Periodo_Aplicaciones_Valida(" & cboPeriodo.ItemData(cboPeriodo.ListIndex) & ") as 'Resultado'"
Call OpenRecordSet(rs, strSQL)
If rs!Resultado = 1 Then
 lblPermiteCambios.Caption = "Permite Cambios"
 lblPermiteCambios.Tag = 1
 lblPermiteCambios.ForeColor = vbBlack
Else
 lblPermiteCambios.Caption = "No se Permiten Cambios!"
 lblPermiteCambios.Tag = 0
 lblPermiteCambios.ForeColor = vbRed
End If
rs.Close

Exit Sub

vError:


End Sub

Private Sub btnInfo_Click()
Dim vMensaje As String

Select Case True
 Case rbCarga(0).Value 'Casos Especiales
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: IMPORT" & vbCrLf _
              & " 3. Columnas.: CEDULA, NOMBRE, SALIDA, PORCENTAJE, DETALLE"
 
 Case rbCarga(1).Value 'Cambios de Salidas
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: IMPORT" & vbCrLf _
              & " 3. Columnas.: CEDULA, NOMBRE, SALIDA, DETALLE, AUTORIZA_IND, AUTORIZA_USUARIO"
End Select


MsgBox vMensaje, vbInformation

End Sub

Private Sub cboPeriodo_Click()
If vPaso Then Exit Sub
If cboPeriodo.ListCount = 0 Then Exit Sub

Call sbPermite_Cambios

Select Case tcMain.SelectedItem
    Case 0
        Call sbListado_Load
    Case 1
        Call sbMass_Limpia
    Case 2
        Call sbCS_Lista
End Select

End Sub

Private Sub Form_Load()

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Cédula", 1500, vbCenter
    .Add , , "Nombre", 3500
    .Add , , "Porcentaje", 1500, vbRightJustify
    .Add , , "Salida", 1500, vbCenter
    .Add , , "Detalle", 2500
End With



vFecha = fxFechaServidor

 tcMain.Item(0).Selected = True

 Call sbBarra_Accion("nuevo")
 Call sbLimpia

 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub
If lsw.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError
   
txtId.Text = Item.Text

txtCedula.Text = Item.SubItems(1)
txtNombre.Text = Item.SubItems(2)
  
txtPorcentaje.Text = Item.SubItems(3)

txtSalida.Text = Item.SubItems(4)

txtDetalle.Text = Item.SubItems(5)


Call sbBarra_Accion("ACTIVO")

txtPorcentaje.SetFocus

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub sbCS_Limpia()

txtCS_ID.Text = ""
txtCS_Identificacion.Text = ""
txtCS_Nombre.Text = ""
txtCS_Detalle.Text = ""
txtCS_Salida.Text = ""
txtCS_Autoriza.Text = ""
End Sub

Private Sub sbCS_Nuevo()

Call sbCS_Limpia

gBusquedas.Convertir = "N"
gBusquedas.Columna = "CEDULA"
gBusquedas.Orden = "CEDULA"
gBusquedas.Consulta = "select CEDULA, NOMBRE, SALIDA_CODIGO From vExc_Cambio_Salida_Nuevo"
gBusquedas.Filtro = " and ID_Periodo = " & cboPeriodo.ItemData(cboPeriodo.ListIndex)

frmBusquedas.Show vbModal
If gBusquedas.Resultado <> "" Then
    txtCS_Identificacion.Text = gBusquedas.Resultado
    txtCS_Nombre.Text = gBusquedas.Resultado2
End If
    

End Sub

Private Sub sbCS_Guarda()
Dim pId As Long

On Error GoTo vError

Me.MousePointer = vbHourglass


If txtCS_ID.Text = "" Then
   pId = 0
Else
   pId = txtCS_ID.Text
End If

txtCS_Identificacion.Text = fxSysCleanTxtInject(txtCS_Identificacion.Text)
txtCS_Detalle.Text = fxSysCleanTxtInject(txtCS_Detalle.Text)

strSQL = "exec spExc_Cambio_Salida_Add " & pId & ", " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & ", '" & txtCS_Identificacion.Text & "', '" & txtCS_Detalle.Text _
       & "', '" & txtCS_Salida.Text & "', '" & glogon.Usuario & "'"

 Call OpenRecordSet(rs, strSQL)
 If rs!Pass = 1 Then
    Call Bitacora("Registra", "Excedente Cambio de Salida Id:  " & txtCS_ID.Text & ", Cedula: " & txtCS_Identificacion.Text & ", Salida: " & txtCS_Salida.Text & ", Periodo: " & cboPeriodo.Text)
    Me.MousePointer = vbDefault
    
    MsgBox "Registro realizado Satisfactoriamente!", vbInformation
    Call sbCS_Limpia
    Call sbCS_Lista
 Else
    Me.MousePointer = vbDefault
    MsgBox "No se puede registrar el Caso por que ya fué procesado!", vbExclamation
 End If


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCS_Elimina()
Dim i As Integer, pId As Long

On Error GoTo vError
     
i = MsgBox("Está Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
    Me.MousePointer = vbHourglass
       
    pId = txtCS_ID.Text
    txtCS_Identificacion.Text = fxSysCleanTxtInject(txtCS_Identificacion.Text)
    
    strSQL = "exec spExc_Cambio_Salida_Delete " & pId & ", " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & ", '" & txtCS_Identificacion.Text _
           & "', '" & txtCS_Salida.Text & "', '" & glogon.Usuario & "'"
    
     Call OpenRecordSet(rs, strSQL)
     If rs!Pass = 1 Then
         Call Bitacora("Borra", "Excedente Cambio de Salida Id:  " & txtCS_ID.Text & ", Cedula: " & txtCS_Identificacion.Text & ", Salida: " & txtCS_Salida.Text & ", Periodo: " & cboPeriodo.Text)
         
        MsgBox "Eliminación Realizada Satisfactoriamente!", vbInformation
        Call sbCS_Limpia
        Call sbCS_Lista
     
'        Call sbBarra_Accion("NUEVO")
'        Call sbListado_Load
'        Call RefrescaTags(Me)
     
     Else
         Me.MousePointer = vbDefault
         MsgBox "No se puede Eliminar el Caso por que ya fué procesado o no Existe!", vbExclamation
     End If
     
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCS_Autoriza()
Dim i As Integer

On Error GoTo vError

If txtCS_ID.Text = "" Then Exit Sub
If txtCS_Autoriza.Text = "" Or Mid(txtCS_Autoriza.Text, 1, 1) = "S" Then Exit Sub


i = MsgBox("Está Seguro que desea Autoriza esta caso?", vbYesNo)
If i = vbYes Then
    
    Me.MousePointer = vbHourglass
    'spExc_Cambio_Salida_Autoriza (@idCS int, @PeriodoId int, @Cedula varchar(20), @Usuario varchar(30), @Autoriza smallint = 1 )
    strSQL = "exec spExc_Cambio_Salida_Autoriza " & txtCS_ID.Text & ", " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
           & ", '" & txtCS_Identificacion.Text & "', '" & glogon.Usuario & "'"
    Call OpenRecordSet(rs, strSQL)
    
    Me.MousePointer = vbDefault
    
    If rs!Pass = 1 Then
        Call Bitacora("Borra", "Excedente Cambio de Salida Id:  " & txtCS_ID.Text & ", Cedula: " & txtCS_Identificacion.Text & ", Salida: " & txtCS_Salida.Text & ", Periodo: " & cboPeriodo.Text)
        MsgBox "Autorización Realizada Satisfactoriamente!", vbInformation
    Else
        MsgBox "Autorización No se Pudo realizar, verifique!", vbInformation
    End If
    
    Call sbCS_Limpia
    Call sbCS_Lista
    
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub

Private Sub sbCS_Lista()

On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pAutorizado As Integer

Select Case True
    Case rbCS_Autorizado.Item(0).Value
        pAutorizado = 2
    Case rbCS_Autorizado.Item(1).Value
        pAutorizado = 0
    Case rbCS_Autorizado.Item(2).Value
        pAutorizado = 1
End Select

lswCS.ListItems.Clear
With lswCS.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Cedula", 1400, vbCenter
    .Add , , "Nombre", 3500
    .Add , , "Salida", 1100, vbCenter
    .Add , , "Detalle", 3000
    .Add , , "Autoriza ?", 1400, vbCenter
    .Add , , "Aut.Fecha", 1800
    .Add , , "Aut.Usuario", 1500, vbCenter
    .Add , , "Reg.Fecha", 1800
    .Add , , "Reg.Usuario", 1500, vbCenter
    .Add , , "Mod.Fecha", 1800
    .Add , , "Mod.Usuario", 1500, vbCenter
    .Add , , "Aut.Id", 800, vbCenter
    .Add , , "Salida", 3100
End With

strSQL = "exec spExc_CambioSalida_Lista " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & ", '" & txtCS_Filtro.Text _
       & "', " & pAutorizado & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    Set itmX = lswCS.ListItems.Add(, , rs!consec)
        itmX.SubItems(1) = rs!Cedula
        itmX.SubItems(2) = rs!Nombre_Desc
        itmX.SubItems(3) = rs!Nueva_Salida
        itmX.SubItems(4) = rs!Detalle
        itmX.SubItems(5) = rs!Autorizado_desc
        itmX.SubItems(6) = rs!Autoriza_Fecha & ""
        itmX.SubItems(7) = rs!Autoriza_Usuario & ""
        itmX.SubItems(8) = rs!Registro_Fecha & ""
        itmX.SubItems(9) = rs!Registro_Usuario & ""
        itmX.SubItems(10) = rs!Modifica_Fecha & ""
        itmX.SubItems(11) = rs!Modifica_Usuario & ""
        itmX.SubItems(12) = rs!Ind_Autorizado
        itmX.SubItems(13) = rs!Salida_Desc
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswCarga_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswCarga.SortKey = ColumnHeader.Index - 1
  If lswCarga.SortOrder = 0 Then lswCarga.SortOrder = 1 Else lswCarga.SortOrder = 0
  lswCarga.Sorted = True
End Sub


Private Sub lswCS_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswCS.SortKey = ColumnHeader.Index - 1
  If lswCS.SortOrder = 0 Then lswCS.SortOrder = 1 Else lswCS.SortOrder = 0
  lswCS.Sorted = True
End Sub

Private Sub lswCS_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)


txtCS_ID.Text = Item.Text
txtCS_Identificacion.Text = Item.SubItems(1)
txtCS_Nombre.Text = Item.SubItems(2)
txtCS_Salida.Text = Item.SubItems(3)
txtCS_Detalle.Text = Item.SubItems(4)
txtCS_Autoriza.Text = Item.SubItems(5)

End Sub


Private Sub rbCarga_Click(Index As Integer)

lswCarga.ListItems.Clear

With lswCarga.ColumnHeaders
    .Clear
    .Add , , "Cedula", 1400, vbCenter
    .Add , , "Nombre", 3500

    If Index = 0 Then
        .Add , , "Porcentaje", 1500, vbRightJustify
        .Add , , "Salida", 1500, vbCenter
        .Add , , "Detalle", 2500
        .Add , , "Inconsistencia?", 3500
    Else
        .Add , , "Salida", 1100, vbCenter
        .Add , , "Detalle", 3000
        .Add , , "Autoriza ?", 1400, vbCenter
        .Add , , "Aut.Fecha", 1800
        .Add , , "Aut.Usuario", 1500, vbCenter
        .Add , , "Inconsistencia?", 3500
    End If

End With

Call sbMass_Limpia

End Sub

Private Sub rbCS_Autorizado_Click(Index As Integer)
Call sbCS_Lista
End Sub


Private Sub sbMass_Limpia()
    lswCarga.ListItems.Clear
    txtArchivo.Text = ""
    txtCasos.Text = "0"
    txtCasosApl.Text = "0"
    txtCasosInco.Text = "0"
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0
    Case 1
    Case 2
        Call sbCS_Lista

End Select

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Me.MousePointer = vbHourglass

vPaso = True


strSQL = "select IdX, ItmX from vExc_Periodos order by Idx desc"
Call sbCbo_Llena_New(cboPeriodo, strSQL, False, True)

vPaso = False


Call sbListado_Load

Me.MousePointer = vbDefault


End Sub


Private Sub txtCS_Filtro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbCS_Lista
End If
End Sub


Private Sub txtCS_Identificacion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Col1Name = "Cédula Colilla"
    gBusquedas.Col2Name = "Cédula Real"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from vExc_Cambio_Salida_Nuevo"
    gBusquedas.Filtro = " AND ID_PERIODO = " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    
    txtCS_Identificacion.Text = Trim(gBusquedas.Resultado)
    txtCS_Nombre.Text = Trim(gBusquedas.Resultado3)
End If

End Sub


Private Sub txtCS_Salida_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "COD_SALIDA"
    gBusquedas.Orden = "COD_SALIDA"
    gBusquedas.Consulta = "select COD_SALIDA, DESCRIPCION From vExc_Casos_Especial_Salidas_Cambio"
    gBusquedas.Filtro = " and REQUIERE_PORCENTAJE = 0"

    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
        txtCS_Salida.Text = Trim(gBusquedas.Resultado)
    End If
End If


End Sub

Private Sub txtPorcentaje_GotFocus()
On Error GoTo vError
  txtPorcentaje.Text = CCur(txtPorcentaje.Text)
vError:
End Sub

Private Sub txtPorcentaje_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
    txtSalida.SetFocus
End If
End Sub

Private Sub txtPorcentaje_LostFocus()
On Error GoTo vError
  
  txtPorcentaje.Text = Format(CCur(txtPorcentaje.Text), "Standard")

vError:

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
    
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Col1Name = "Cédula Colilla"
    gBusquedas.Col2Name = "Cédula Real"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from vExc_Casos_Especial_Nuevo"
    gBusquedas.Filtro = " AND ID_PERIODO = " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    
    txtCedula.Text = Trim(gBusquedas.Resultado)

    If Trim(txtCedula.Text) <> "" Then
        txtNombre.SetFocus
    End If
End If

End Sub

Private Sub txtCedula_LostFocus()

If Trim(txtCedula) <> "" Then
   txtNombre.Text = fxNombre(Trim(txtCedula))
   If Trim(txtNombre.Text) = "" Then
      txtId.Text = ""
      MsgBox "Cedula Incorrecta", vbInformation
      txtCedula.SetFocus
   Else
      Call sbCedula_Load
   End If
Else
   txtNombre.Text = ""
   txtId.Text = ""
End If

End Sub


Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbListado_Load
End If
End Sub


Private Sub txtLineas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbListado_Load
End If
End Sub


Private Sub txtSalida_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "COD_SALIDA"
    gBusquedas.Orden = "COD_SALIDA"
    
    gBusquedas.Consulta = "select COD_SALIDA, DESCRIPCION From vExc_Casos_Especial_Salidas_Cambio"
    gBusquedas.Filtro = " and REQUIERE_PORCENTAJE = 1"
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
       txtSalida.Text = gBusquedas.Resultado
    End If
End If

End Sub
