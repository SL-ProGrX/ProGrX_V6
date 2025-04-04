VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_Beneficios_Integral 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Beneficios y Ayudas Sociales: Registro Integral"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   13320
      Top             =   480
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   13815
      _Version        =   1441793
      _ExtentX        =   24368
      _ExtentY        =   11456
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
      ItemCount       =   11
      SelectedItem    =   2
      Item(0).Caption =   "Datos de la Persona"
      Item(0).ControlCount=   40
      Item(0).Control(0)=   "cboProvincia"
      Item(0).Control(1)=   "cboCanton"
      Item(0).Control(2)=   "cboDistrito"
      Item(0).Control(3)=   "cboNacionalidad"
      Item(0).Control(4)=   "txtEmail"
      Item(0).Control(5)=   "txtEmail_02"
      Item(0).Control(6)=   "txtApartado"
      Item(0).Control(7)=   "txtDireccion"
      Item(0).Control(8)=   "cboPaisNac"
      Item(0).Control(9)=   "Label10(0)"
      Item(0).Control(10)=   "Label11"
      Item(0).Control(11)=   "Label10(9)"
      Item(0).Control(12)=   "Label7"
      Item(0).Control(13)=   "Label18(2)"
      Item(0).Control(14)=   "Label18(10)"
      Item(0).Control(15)=   "cboSexo"
      Item(0).Control(16)=   "dtpNacimiento"
      Item(0).Control(17)=   "txtApellido1"
      Item(0).Control(18)=   "txtApellido2"
      Item(0).Control(19)=   "txtNombre(1)"
      Item(0).Control(20)=   "Label10(1)"
      Item(0).Control(21)=   "Label10(2)"
      Item(0).Control(22)=   "Label10(3)"
      Item(0).Control(23)=   "Label18(0)"
      Item(0).Control(24)=   "Label18(1)"
      Item(0).Control(25)=   "Label18(3)"
      Item(0).Control(26)=   "Label18(4)"
      Item(0).Control(27)=   "dtpFechaIngreso"
      Item(0).Control(28)=   "ShortcutCaption2"
      Item(0).Control(29)=   "lswTel"
      Item(0).Control(30)=   "Label18(5)"
      Item(0).Control(31)=   "txtNombre(2)"
      Item(0).Control(32)=   "Label18(6)"
      Item(0).Control(33)=   "Label18(7)"
      Item(0).Control(34)=   "cboNivelAcademico"
      Item(0).Control(35)=   "txtPuestoDesc"
      Item(0).Control(36)=   "PushButton1"
      Item(0).Control(37)=   "ShortcutCaption3(0)"
      Item(0).Control(38)=   "FlatEdit4"
      Item(0).Control(39)=   "cboEstadoCivil"
      Item(1).Caption =   "Consulta"
      Item(1).ControlCount=   16
      Item(1).Control(0)=   "Label1(0)"
      Item(1).Control(1)=   "Label1(1)"
      Item(1).Control(2)=   "Label1(2)"
      Item(1).Control(3)=   "Label1(3)"
      Item(1).Control(4)=   "Label1(4)"
      Item(1).Control(5)=   "Label1(5)"
      Item(1).Control(6)=   "lswConsulta"
      Item(1).Control(7)=   "Label1(6)"
      Item(1).Control(8)=   "txtC_Expediente"
      Item(1).Control(9)=   "dtpC_Inicio"
      Item(1).Control(10)=   "dtpC_Corte"
      Item(1).Control(11)=   "txtC_Identificacion"
      Item(1).Control(12)=   "txtC_Nombre"
      Item(1).Control(13)=   "cboC_Estado"
      Item(1).Control(14)=   "txtC_Usuario"
      Item(1).Control(15)=   "cboC_Tipo"
      Item(2).Caption =   "Orden de Pago"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "frmMonetario"
      Item(2).Control(1)=   "GroupBox1"
      Item(3).Caption =   "Generales"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "gbMain"
      Item(3).Control(1)=   "tcBene_Aux"
      Item(4).Caption =   "Apremiantes"
      Item(4).ControlCount=   12
      Item(4).Control(0)=   "Label3(15)"
      Item(4).Control(1)=   "cboA_Categoria"
      Item(4).Control(2)=   "cboA_Profesional"
      Item(4).Control(3)=   "Label3(16)"
      Item(4).Control(4)=   "Label3(17)"
      Item(4).Control(5)=   "cboA_Motivo"
      Item(4).Control(6)=   "lswA_Motivos"
      Item(4).Control(7)=   "ShortcutCaption3(1)"
      Item(4).Control(8)=   "Label3(18)"
      Item(4).Control(9)=   "lblA_NotasQty"
      Item(4).Control(10)=   "btnA_Motivo_Guarda"
      Item(4).Control(11)=   "txtA_Motivo"
      Item(5).Caption =   "Reconocimientos"
      Item(5).ControlCount=   55
      Item(5).Control(0)=   "ShortcutCaption3(2)"
      Item(5).Control(1)=   "Label1(14)"
      Item(5).Control(2)=   "Label1(15)"
      Item(5).Control(3)=   "Label1(16)"
      Item(5).Control(4)=   "Label1(17)"
      Item(5).Control(5)=   "Label1(18)"
      Item(5).Control(6)=   "Label1(19)"
      Item(5).Control(7)=   "Label1(20)"
      Item(5).Control(8)=   "Label1(21)"
      Item(5).Control(9)=   "Label1(22)"
      Item(5).Control(10)=   "Label1(23)"
      Item(5).Control(11)=   "Label1(24)"
      Item(5).Control(12)=   "Label1(25)"
      Item(5).Control(13)=   "txtR_Cedula"
      Item(5).Control(14)=   "txtR_Nombre"
      Item(5).Control(15)=   "dtpR_FechaNac"
      Item(5).Control(16)=   "txtR_Apellido1"
      Item(5).Control(17)=   "txtR_Apellido2"
      Item(5).Control(18)=   "txtR_Edad"
      Item(5).Control(19)=   "cboR_Centro"
      Item(5).Control(20)=   "cboR_GradoAnterior"
      Item(5).Control(21)=   "txtR_CentroEducativo"
      Item(5).Control(22)=   "cboR_Nivel"
      Item(5).Control(23)=   "cboR_Tipo"
      Item(5).Control(24)=   "btnR_Guardar"
      Item(5).Control(25)=   "txtR_Observaciones"
      Item(5).Control(26)=   "Label3(26)"
      Item(5).Control(27)=   "lblR_Observaciones"
      Item(5).Control(28)=   "cboR_Genero"
      Item(5).Control(29)=   "ShortcutCaption3(3)"
      Item(5).Control(30)=   "ShortcutCaption3(4)"
      Item(5).Control(31)=   "txtR_Matematicas"
      Item(5).Control(32)=   "cboR_MateAplica"
      Item(5).Control(33)=   "Label1(26)"
      Item(5).Control(34)=   "Label1(27)"
      Item(5).Control(35)=   "Label1(28)"
      Item(5).Control(36)=   "Label1(29)"
      Item(5).Control(37)=   "Label1(30)"
      Item(5).Control(38)=   "Label1(31)"
      Item(5).Control(39)=   "Label1(32)"
      Item(5).Control(40)=   "Label1(33)"
      Item(5).Control(41)=   "Label1(34)"
      Item(5).Control(42)=   "Label1(35)"
      Item(5).Control(43)=   "Label1(36)"
      Item(5).Control(44)=   "txtR_Espanol"
      Item(5).Control(45)=   "cboR_EspanolAplica"
      Item(5).Control(46)=   "txtR_Ciencias"
      Item(5).Control(47)=   "cboR_CienciasAplica"
      Item(5).Control(48)=   "txtR_EstudiosSociales"
      Item(5).Control(49)=   "cboR_EstudiosSocialesAplica"
      Item(5).Control(50)=   "txtR_Idioma"
      Item(5).Control(51)=   "txtR_Promedio"
      Item(5).Control(52)=   "cboR_Merito_Deportivo"
      Item(5).Control(53)=   "cboR_Merito_Cientifico"
      Item(5).Control(54)=   "cboR_Merito_Artistico"
      Item(6).Caption =   "Observaciones"
      Item(6).ControlCount=   5
      Item(6).Control(0)=   "ShortcutCaption4(0)"
      Item(6).Control(1)=   "Label3(9)"
      Item(6).Control(2)=   "lswObservaciones"
      Item(6).Control(3)=   "txtObservacionAdd"
      Item(6).Control(4)=   "btnObservacion"
      Item(7).Caption =   "Bit�cora"
      Item(7).ControlCount=   1
      Item(7).Control(0)=   "lswBitacora"
      Item(8).Caption =   "Requisitos"
      Item(8).ControlCount=   1
      Item(8).Control(0)=   "lswRequisitos"
      Item(9).Caption =   "CRECE"
      Item(9).ControlCount=   1
      Item(9).Control(0)=   "tcCRECE"
      Item(10).Caption=   "Sanciones"
      Item(10).ControlCount=   14
      Item(10).Control(0)=   "cboSancionMotivo"
      Item(10).Control(1)=   "Label3(10)"
      Item(10).Control(2)=   "Label3(11)"
      Item(10).Control(3)=   "Label3(12)"
      Item(10).Control(4)=   "txtSancionNotas"
      Item(10).Control(5)=   "Label3(13)"
      Item(10).Control(6)=   "txtSancionId"
      Item(10).Control(7)=   "dtpSancionInicio"
      Item(10).Control(8)=   "dtpSancionCorte"
      Item(10).Control(9)=   "lswSanciones"
      Item(10).Control(10)=   "ShortcutCaption4(1)"
      Item(10).Control(11)=   "chkSancionActica"
      Item(10).Control(12)=   "btnSancion(0)"
      Item(10).Control(13)=   "btnSancion(1)"
      Begin XtremeSuiteControls.ListView lswA_Motivos 
         Height          =   3135
         Left            =   -67840
         TabIndex        =   122
         Top             =   1920
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswConsulta 
         Height          =   5295
         Left            =   -69880
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   13575
         _Version        =   1441793
         _ExtentX        =   23945
         _ExtentY        =   9340
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswObservaciones 
         Height          =   4215
         Left            =   -69880
         TabIndex        =   92
         Top             =   2160
         Visible         =   0   'False
         Width           =   13575
         _Version        =   1441793
         _ExtentX        =   23945
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswBitacora 
         Height          =   6135
         Left            =   -70000
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   13815
         _Version        =   1441793
         _ExtentX        =   24368
         _ExtentY        =   10821
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswRequisitos 
         Height          =   6135
         Left            =   -70000
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   13815
         _Version        =   1441793
         _ExtentX        =   24368
         _ExtentY        =   10821
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswSanciones 
         Height          =   3135
         Left            =   -69880
         TabIndex        =   104
         Top             =   3240
         Visible         =   0   'False
         Width           =   13575
         _Version        =   1441793
         _ExtentX        =   23945
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtR_Apellido2 
         Height          =   315
         Left            =   -63280
         TabIndex        =   205
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.TabControl tcCRECE 
         Height          =   6015
         Left            =   -70000
         TabIndex        =   170
         Top             =   480
         Visible         =   0   'False
         Width           =   13815
         _Version        =   1441793
         _ExtentX        =   24368
         _ExtentY        =   10610
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
         Item(0).Caption =   "Educaci�n Financiera"
         Item(0).ControlCount=   3
         Item(0).Control(0)=   "cboC_Capacitacion"
         Item(0).Control(1)=   "Label1(7)"
         Item(0).Control(2)=   "gbSesiones"
         Item(1).Caption =   "Producto Financiero"
         Item(1).ControlCount=   9
         Item(1).Control(0)=   "cboC_ProdFinanciero"
         Item(1).Control(1)=   "Label1(11)"
         Item(1).Control(2)=   "vgCrece_PF"
         Item(1).Control(3)=   "Label1(12)"
         Item(1).Control(4)=   "txtC_PF_Obsevacion"
         Item(1).Control(5)=   "btnC_ProductoFinanciero"
         Item(1).Control(6)=   "lblC_PF_Obsevacion"
         Item(1).Control(7)=   "Label3(24)"
         Item(1).Control(8)=   "txtC_Liquidez"
         Item(2).Caption =   "Beneficio Solidario"
         Item(2).ControlCount=   7
         Item(2).Control(0)=   "Label1(13)"
         Item(2).Control(1)=   "fpSpread1"
         Item(2).Control(2)=   "PushButton2"
         Item(2).Control(3)=   "Label3(25)"
         Item(2).Control(4)=   "cboC_BeneficioSolidario"
         Item(2).Control(5)=   "lblC_BS_Notas"
         Item(2).Control(6)=   "txtC_BS_Notas"
         Begin XtremeSuiteControls.GroupBox gbSesiones 
            Height          =   4215
            Left            =   480
            TabIndex        =   173
            Top             =   1560
            Width           =   12255
            _Version        =   1441793
            _ExtentX        =   21616
            _ExtentY        =   7435
            _StockProps     =   79
            Caption         =   "Sesiones"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.ListView lswC_Sesiones 
               Height          =   3615
               Left            =   0
               TabIndex        =   174
               Top             =   480
               Width           =   6735
               _Version        =   1441793
               _ExtentX        =   11880
               _ExtentY        =   6376
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
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboC_Sesion 
               Height          =   330
               Left            =   7920
               TabIndex        =   175
               Top             =   720
               Width           =   3015
               _Version        =   1441793
               _ExtentX        =   5318
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
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.PushButton btnC_Sesion 
               Height          =   450
               Index           =   0
               Left            =   8160
               TabIndex        =   177
               Top             =   3480
               Width           =   1335
               _Version        =   1441793
               _ExtentX        =   2355
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "Agregar"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Picture         =   "frmAF_Beneficios_Integral.frx":0000
            End
            Begin XtremeSuiteControls.PushButton btnC_Sesion 
               Height          =   450
               Index           =   1
               Left            =   9480
               TabIndex        =   178
               Top             =   3480
               Width           =   1335
               _Version        =   1441793
               _ExtentX        =   2355
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "Eliminar"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Picture         =   "frmAF_Beneficios_Integral.frx":0731
            End
            Begin XtremeSuiteControls.ComboBox cboC_Asistencia 
               Height          =   330
               Left            =   7920
               TabIndex        =   179
               Top             =   1440
               Width           =   3015
               _Version        =   1441793
               _ExtentX        =   5318
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
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox cboC_Tareas 
               Height          =   330
               Left            =   7920
               TabIndex        =   181
               Top             =   2160
               Width           =   3015
               _Version        =   1441793
               _ExtentX        =   5318
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
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Index           =   10
               Left            =   7920
               TabIndex        =   182
               Top             =   1920
               Width           =   2175
               _Version        =   1441793
               _ExtentX        =   3836
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Tareas"
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
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Index           =   9
               Left            =   7920
               TabIndex        =   180
               Top             =   1200
               Width           =   2175
               _Version        =   1441793
               _ExtentX        =   3836
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Asistencia"
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
            Begin XtremeSuiteControls.Label Label1 
               Height          =   255
               Index           =   8
               Left            =   7920
               TabIndex        =   176
               Top             =   480
               Width           =   2175
               _Version        =   1441793
               _ExtentX        =   3836
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Sesi�n"
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
         End
         Begin XtremeSuiteControls.ComboBox cboC_Capacitacion 
            Height          =   330
            Left            =   2880
            TabIndex        =   171
            Top             =   960
            Width           =   4335
            _Version        =   1441793
            _ExtentX        =   7646
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboC_ProdFinanciero 
            Height          =   330
            Left            =   -64360
            TabIndex        =   183
            Top             =   600
            Visible         =   0   'False
            Width           =   2775
            _Version        =   1441793
            _ExtentX        =   4895
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin FPSpreadADO.fpSpread vgCrece_PF 
            Height          =   1815
            Left            =   -68800
            TabIndex        =   185
            Top             =   1080
            Visible         =   0   'False
            Width           =   10935
            _Version        =   524288
            _ExtentX        =   19288
            _ExtentY        =   3201
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
            MaxCols         =   3
            MaxRows         =   3
            ScrollBars      =   2
            SpreadDesigner  =   "frmAF_Beneficios_Integral.frx":0CD5
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtC_PF_Obsevacion 
            Height          =   795
            Left            =   -67360
            TabIndex        =   187
            Top             =   3960
            Visible         =   0   'False
            Width           =   7335
            _Version        =   1441793
            _ExtentX        =   12938
            _ExtentY        =   1402
            _StockProps     =   77
            ForeColor       =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnC_ProductoFinanciero 
            Height          =   330
            Left            =   -59680
            TabIndex        =   188
            Top             =   4320
            Visible         =   0   'False
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1080
            _ExtentY        =   573
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_Beneficios_Integral.frx":128D
         End
         Begin XtremeSuiteControls.FlatEdit txtC_Liquidez 
            Height          =   345
            Left            =   -67360
            TabIndex        =   191
            Top             =   3240
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   609
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
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
         Begin XtremeSuiteControls.ComboBox cboC_BeneficioSolidario 
            Height          =   330
            Left            =   -64720
            TabIndex        =   192
            Top             =   600
            Visible         =   0   'False
            Width           =   2775
            _Version        =   1441793
            _ExtentX        =   4895
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin FPSpreadADO.fpSpread fpSpread1 
            Height          =   2295
            Left            =   -69160
            TabIndex        =   194
            Top             =   1080
            Visible         =   0   'False
            Width           =   10935
            _Version        =   524288
            _ExtentX        =   19288
            _ExtentY        =   4048
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
            MaxCols         =   3
            MaxRows         =   3
            ScrollBars      =   2
            SpreadDesigner  =   "frmAF_Beneficios_Integral.frx":19BE
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   330
            Left            =   -60040
            TabIndex        =   195
            Top             =   4320
            Visible         =   0   'False
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1080
            _ExtentY        =   573
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_Beneficios_Integral.frx":1F67
         End
         Begin XtremeSuiteControls.FlatEdit txtC_BS_Notas 
            Height          =   795
            Left            =   -67720
            TabIndex        =   198
            Top             =   3960
            Visible         =   0   'False
            Width           =   7335
            _Version        =   1441793
            _ExtentX        =   12938
            _ExtentY        =   1402
            _StockProps     =   77
            ForeColor       =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   25
            Left            =   -69160
            TabIndex        =   197
            Top             =   3960
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observaci�n"
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
         Begin XtremeSuiteControls.Label lblC_BS_Notas 
            Height          =   255
            Left            =   -67720
            TabIndex        =   196
            Top             =   4800
            Visible         =   0   'False
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8281
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "0 caracteres de 300 permitidos"
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   13
            Left            =   -69040
            TabIndex        =   193
            Top             =   600
            Visible         =   0   'False
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Aplica Beneficio Solidario"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   24
            Left            =   -68800
            TabIndex        =   190
            Top             =   3960
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observaci�n"
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
         Begin XtremeSuiteControls.Label lblC_PF_Obsevacion 
            Height          =   255
            Left            =   -67360
            TabIndex        =   189
            Top             =   4800
            Visible         =   0   'False
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8281
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "0 caracteres de 300 permitidos"
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   12
            Left            =   -68800
            TabIndex        =   186
            Top             =   3240
            Visible         =   0   'False
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Liquidez"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   11
            Left            =   -68680
            TabIndex        =   184
            Top             =   600
            Visible         =   0   'False
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Aplicaci�n del Producto Financiero"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   172
            Top             =   960
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Capacitaci�n Completa"
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
      End
      Begin XtremeSuiteControls.CheckBox chkSancionActica 
         Height          =   375
         Left            =   -64840
         TabIndex        =   106
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Activa ?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpSancionCorte 
         Height          =   330
         Left            =   -65440
         TabIndex        =   103
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Expediente 
         Height          =   330
         Left            =   -61600
         TabIndex        =   88
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.ListView lswTel 
         Height          =   2295
         Left            =   -59320
         TabIndex        =   77
         Top             =   3480
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   4048
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Index           =   1
         Left            =   -62560
         TabIndex        =   66
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
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
      Begin XtremeSuiteControls.TabControl tcBene_Aux 
         Height          =   2655
         Left            =   -69400
         TabIndex        =   42
         Top             =   3600
         Visible         =   0   'False
         Width           =   12375
         _Version        =   1441793
         _ExtentX        =   21828
         _ExtentY        =   4683
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
         Item(0).Caption =   "Sepelio"
         Item(0).ControlCount=   3
         Item(0).Control(0)=   "txtNombreFallecido"
         Item(0).Control(1)=   "txtCedulaFallecido"
         Item(0).Control(2)=   "Label3(6)"
         Item(1).Caption =   "Desastres Natural"
         Item(1).ControlCount=   0
         Item(2).Caption =   "Fondo Emergencias Nacionales"
         Item(2).ControlCount=   4
         Item(2).Control(0)=   "Label3(7)"
         Item(2).Control(1)=   "Label3(8)"
         Item(2).Control(2)=   "txtFena_Descripcion"
         Item(2).Control(3)=   "txtFena_Emergencia"
         Item(3).Caption =   "Monto del Beneficio"
         Item(3).ControlCount=   8
         Item(3).Control(0)=   "Label3(19)"
         Item(3).Control(1)=   "Label3(20)"
         Item(3).Control(2)=   "Label3(21)"
         Item(3).Control(3)=   "txtMnt_Aplicado"
         Item(3).Control(4)=   "txtMnt_Aprobado"
         Item(3).Control(5)=   "txtMnt_Notas"
         Item(3).Control(6)=   "lblMnt_Notas"
         Item(3).Control(7)=   "btnMonto_Guarda"
         Item(4).Caption =   "Estado"
         Item(4).ControlCount=   6
         Item(4).Control(0)=   "cboBene_Estado"
         Item(4).Control(1)=   "Label3(22)"
         Item(4).Control(2)=   "Label3(23)"
         Item(4).Control(3)=   "txtEstado_Notas"
         Item(4).Control(4)=   "lblEstado_Notas"
         Item(4).Control(5)=   "btnEstado"
         Begin XtremeSuiteControls.FlatEdit txtEstado_Notas 
            Height          =   795
            Left            =   -68320
            TabIndex        =   169
            Top             =   1080
            Visible         =   0   'False
            Width           =   7335
            _Version        =   1441793
            _ExtentX        =   12938
            _ExtentY        =   1402
            _StockProps     =   77
            ForeColor       =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNombreFallecido 
            Height          =   315
            Left            =   3360
            TabIndex        =   43
            Top             =   720
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11245
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCedulaFallecido 
            Height          =   315
            Left            =   1680
            TabIndex        =   44
            Top             =   720
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2984
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFena_Descripcion 
            Height          =   1275
            Left            =   -67240
            TabIndex        =   89
            Top             =   1080
            Visible         =   0   'False
            Width           =   7575
            _Version        =   1441793
            _ExtentX        =   13361
            _ExtentY        =   2249
            _StockProps     =   77
            ForeColor       =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFena_Emergencia 
            Height          =   330
            Left            =   -67240
            TabIndex        =   121
            Top             =   480
            Visible         =   0   'False
            Width           =   7575
            _Version        =   1441793
            _ExtentX        =   13361
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMnt_Aplicado 
            Height          =   345
            Left            =   -63280
            TabIndex        =   131
            Top             =   600
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   609
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMnt_Aprobado 
            Height          =   345
            Left            =   -68320
            TabIndex        =   132
            Top             =   600
            Visible         =   0   'False
            Width           =   2535
            _Version        =   1441793
            _ExtentX        =   4471
            _ExtentY        =   609
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777215
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMnt_Notas 
            Height          =   795
            Left            =   -68320
            TabIndex        =   133
            Top             =   1080
            Visible         =   0   'False
            Width           =   7335
            _Version        =   1441793
            _ExtentX        =   12938
            _ExtentY        =   1402
            _StockProps     =   77
            ForeColor       =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnMonto_Guarda 
            Height          =   330
            Left            =   -60640
            TabIndex        =   135
            Top             =   1440
            Visible         =   0   'False
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1080
            _ExtentY        =   573
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_Beneficios_Integral.frx":2698
         End
         Begin XtremeSuiteControls.ComboBox cboBene_Estado 
            Height          =   315
            Left            =   -68320
            TabIndex        =   165
            Top             =   600
            Visible         =   0   'False
            Width           =   2535
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
         Begin XtremeSuiteControls.PushButton btnEstado 
            Height          =   330
            Left            =   -60640
            TabIndex        =   143
            Top             =   1440
            Visible         =   0   'False
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1080
            _ExtentY        =   573
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_Beneficios_Integral.frx":2DC9
         End
         Begin XtremeSuiteControls.Label lblEstado_Notas 
            Height          =   255
            Left            =   -68320
            TabIndex        =   153
            Top             =   1920
            Visible         =   0   'False
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8281
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "0 caracteres de 300 permitidos"
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   23
            Left            =   -70000
            TabIndex        =   168
            Top             =   1080
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observaci�n"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   22
            Left            =   -69040
            TabIndex        =   167
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Estado"
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
         Begin XtremeSuiteControls.Label lblMnt_Notas 
            Height          =   255
            Left            =   -68320
            TabIndex        =   134
            Top             =   1920
            Visible         =   0   'False
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8281
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "0 caracteres de 300 permitidos"
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   21
            Left            =   -69760
            TabIndex        =   130
            Top             =   1080
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observaci�n"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   20
            Left            =   -64840
            TabIndex        =   129
            Top             =   600
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Monto Aplicado"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   19
            Left            =   -69760
            TabIndex        =   128
            Top             =   600
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Monto Aprobado"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   615
            Index           =   8
            Left            =   -68800
            TabIndex        =   91
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Nombre de la Emergencia"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   615
            Index           =   7
            Left            =   -68800
            TabIndex        =   90
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Descripci�n de la Emergencia"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   45
            Top             =   720
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fallecido"
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
      End
      Begin XtremeSuiteControls.DateTimePicker dtpC_Inicio 
         Height          =   330
         Left            =   -68320
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpC_Corte 
         Height          =   330
         Left            =   -66880
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Identificacion 
         Height          =   330
         Left            =   -65440
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtC_Nombre 
         Height          =   330
         Left            =   -63640
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboC_Estado 
         Height          =   330
         Left            =   -59800
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtC_Usuario 
         Height          =   330
         Left            =   -58120
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.ComboBox cboC_Tipo 
         Height          =   330
         Left            =   -69880
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.GroupBox frmMonetario 
         Height          =   3255
         Left            =   1440
         TabIndex        =   26
         Top             =   480
         Width           =   10455
         _Version        =   1441793
         _ExtentX        =   18441
         _ExtentY        =   5741
         _StockProps     =   79
         Caption         =   "Registro para el desembolso del Beneficio o Ayuda:"
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.ListView lswPago 
            Height          =   1695
            Left            =   0
            TabIndex        =   27
            Top             =   960
            Width           =   10215
            _Version        =   1441793
            _ExtentX        =   18018
            _ExtentY        =   2990
            _StockProps     =   77
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   315
            Left            =   2760
            TabIndex        =   28
            Top             =   480
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   550
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDisponible 
            Height          =   315
            Left            =   8160
            TabIndex        =   29
            Top             =   480
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3619
            _ExtentY        =   550
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnPago 
            Height          =   330
            Index           =   0
            Left            =   8400
            TabIndex        =   162
            Top             =   2780
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1080
            _ExtentY        =   573
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frmAF_Beneficios_Integral.frx":34FA
         End
         Begin XtremeSuiteControls.PushButton btnPago 
            Height          =   330
            Index           =   1
            Left            =   9000
            TabIndex        =   163
            Top             =   2780
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1080
            _ExtentY        =   573
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frmAF_Beneficios_Integral.frx":3B2C
         End
         Begin XtremeSuiteControls.PushButton btnPago 
            Height          =   330
            Index           =   2
            Left            =   9600
            TabIndex        =   164
            Top             =   2780
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1080
            _ExtentY        =   573
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "frmAF_Beneficios_Integral.frx":425D
         End
         Begin XtremeShortcutBar.ShortcutCaption lblPagoCaso 
            Height          =   375
            Left            =   0
            TabIndex        =   146
            Top             =   2760
            Width           =   10215
            _Version        =   1441793
            _ExtentX        =   18018
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Seleccione un Caso de Pago"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   3
            Left            =   6000
            TabIndex        =   32
            Top             =   480
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Disponible"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   31
            Top             =   480
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Monto"
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
      Begin XtremeSuiteControls.GroupBox gbMain 
         Height          =   3975
         Left            =   -69280
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   12495
         _Version        =   1441793
         _ExtentX        =   22040
         _ExtentY        =   7011
         _StockProps     =   79
         Caption         =   "Datos del Beneficio o Ayuda: "
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.ComboBox cboBeneTipo 
            Height          =   330
            Left            =   6240
            TabIndex        =   34
            Top             =   720
            Width           =   2655
            _Version        =   1441793
            _ExtentX        =   4683
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
         Begin XtremeSuiteControls.FlatEdit txtBene_Notas 
            Height          =   795
            Left            =   1560
            TabIndex        =   35
            Top             =   1440
            Width           =   7335
            _Version        =   1441793
            _ExtentX        =   12938
            _ExtentY        =   1402
            _StockProps     =   77
            ForeColor       =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAutorizaUsuario 
            Height          =   345
            Left            =   4200
            TabIndex        =   110
            Top             =   2640
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   609
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
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
         Begin XtremeSuiteControls.FlatEdit txtAutorizaFecha 
            Height          =   345
            Left            =   6600
            TabIndex        =   111
            Top             =   2640
            Width           =   2295
            _Version        =   1441793
            _ExtentX        =   4048
            _ExtentY        =   609
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
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
         Begin XtremeSuiteControls.ComboBox cboBeneficio 
            Height          =   330
            Left            =   1560
            TabIndex        =   113
            Top             =   720
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8281
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
         Begin XtremeSuiteControls.FlatEdit txtEstado 
            Height          =   345
            Left            =   1560
            TabIndex        =   166
            Top             =   2640
            Width           =   2535
            _Version        =   1441793
            _ExtentX        =   4471
            _ExtentY        =   609
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   114
            Top             =   480
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Beneficio"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   5
            Left            =   6600
            TabIndex        =   41
            Top             =   2400
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Autorizado Fecha"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   4
            Left            =   4200
            TabIndex        =   40
            Top             =   2400
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Autorizado por"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   39
            Top             =   2400
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Estado"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   38
            Top             =   1200
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Notas"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   1
            Left            =   6240
            TabIndex        =   37
            Top             =   480
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo"
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
      End
      Begin XtremeSuiteControls.ComboBox cboProvincia 
         Height          =   330
         Left            =   -66040
         TabIndex        =   46
         Top             =   4680
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.ComboBox cboCanton 
         Height          =   330
         Left            =   -64120
         TabIndex        =   47
         Top             =   4680
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.ComboBox cboDistrito 
         Height          =   330
         Left            =   -61840
         TabIndex        =   48
         Top             =   4680
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.ComboBox cboNacionalidad 
         Height          =   330
         Left            =   -61960
         TabIndex        =   49
         Top             =   3120
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   315
         Left            =   -66040
         TabIndex        =   50
         Top             =   3600
         Visible         =   0   'False
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11456
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
      Begin XtremeSuiteControls.FlatEdit txtEmail_02 
         Height          =   315
         Left            =   -66040
         TabIndex        =   51
         Top             =   3960
         Visible         =   0   'False
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11456
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
      Begin XtremeSuiteControls.FlatEdit txtApartado 
         Height          =   315
         Left            =   -66040
         TabIndex        =   52
         Top             =   4320
         Visible         =   0   'False
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11456
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
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   675
         Left            =   -66040
         TabIndex        =   53
         Top             =   5040
         Visible         =   0   'False
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11456
         _ExtentY        =   1191
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
      Begin XtremeSuiteControls.ComboBox cboPaisNac 
         Height          =   330
         Left            =   -66040
         TabIndex        =   54
         Top             =   3120
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.ComboBox cboSexo 
         Height          =   330
         Left            =   -66040
         TabIndex        =   61
         Top             =   1560
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboEstadoCivil 
         Height          =   330
         Left            =   -66040
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.DateTimePicker dtpNacimiento 
         Height          =   315
         Left            =   -61000
         TabIndex        =   63
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
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
      Begin XtremeSuiteControls.FlatEdit txtApellido1 
         Height          =   315
         Left            =   -67600
         TabIndex        =   64
         Top             =   840
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
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
      Begin XtremeSuiteControls.FlatEdit txtApellido2 
         Height          =   315
         Left            =   -65080
         TabIndex        =   65
         Top             =   840
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
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
      Begin XtremeSuiteControls.DateTimePicker dtpFechaIngreso 
         Height          =   315
         Left            =   -61000
         TabIndex        =   75
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
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
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Index           =   2
         Left            =   -66040
         TabIndex        =   79
         Top             =   2160
         Visible         =   0   'False
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11456
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
      Begin XtremeSuiteControls.ComboBox cboNivelAcademico 
         Height          =   330
         Left            =   -66040
         TabIndex        =   82
         Top             =   2640
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.FlatEdit txtPuestoDesc 
         Height          =   330
         Left            =   -61960
         TabIndex        =   83
         Top             =   2640
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   450
         Left            =   -66040
         TabIndex        =   84
         Top             =   5880
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Actualizar Datos de Contacto"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Beneficios_Integral.frx":4801
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit4 
         Height          =   1515
         Left            =   -59320
         TabIndex        =   86
         Top             =   1200
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   2672
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtObservacionAdd 
         Height          =   1155
         Left            =   -67000
         TabIndex        =   94
         Top             =   480
         Visible         =   0   'False
         Width           =   7575
         _Version        =   1441793
         _ExtentX        =   13361
         _ExtentY        =   2037
         _StockProps     =   77
         ForeColor       =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnObservacion 
         Height          =   330
         Left            =   -59320
         TabIndex        =   36
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
         _Version        =   1441793
         _ExtentX        =   1080
         _ExtentY        =   573
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Beneficios_Integral.frx":4F32
      End
      Begin XtremeSuiteControls.ComboBox cboSancionMotivo 
         Height          =   345
         Left            =   -66880
         TabIndex        =   95
         Top             =   960
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin XtremeSuiteControls.FlatEdit txtSancionNotas 
         Height          =   915
         Left            =   -66880
         TabIndex        =   96
         Top             =   1800
         Visible         =   0   'False
         Width           =   7575
         _Version        =   1441793
         _ExtentX        =   13361
         _ExtentY        =   1614
         _StockProps     =   77
         ForeColor       =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSancionId 
         Height          =   450
         Left            =   -66880
         TabIndex        =   101
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   794
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpSancionInicio 
         Height          =   330
         Left            =   -66880
         TabIndex        =   102
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnSancion 
         Height          =   450
         Index           =   0
         Left            =   -63160
         TabIndex        =   107
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Nuevo"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmAF_Beneficios_Integral.frx":5663
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnSancion 
         Height          =   450
         Index           =   1
         Left            =   -61840
         TabIndex        =   108
         Top             =   480
         Visible         =   0   'False
         Width           =   615
         _Version        =   1441793
         _ExtentX        =   1085
         _ExtentY        =   794
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Beneficios_Integral.frx":5D83
      End
      Begin XtremeSuiteControls.ComboBox cboA_Categoria 
         Height          =   330
         Left            =   -67840
         TabIndex        =   112
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.ComboBox cboA_Profesional 
         Height          =   330
         Left            =   -63040
         TabIndex        =   117
         Top             =   600
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
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
      Begin XtremeSuiteControls.ComboBox cboA_Motivo 
         Height          =   330
         Left            =   -67840
         TabIndex        =   120
         Top             =   1080
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
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
      Begin XtremeSuiteControls.FlatEdit txtA_Motivo 
         Height          =   795
         Left            =   -67840
         TabIndex        =   125
         Top             =   5160
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   1402
         _StockProps     =   77
         ForeColor       =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnA_Motivo_Guarda 
         Height          =   330
         Left            =   -58000
         TabIndex        =   127
         Top             =   5640
         Visible         =   0   'False
         Width           =   615
         _Version        =   1441793
         _ExtentX        =   1080
         _ExtentY        =   573
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Beneficios_Integral.frx":64B4
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2415
         Left            =   1560
         TabIndex        =   136
         Top             =   3840
         Width           =   10095
         _Version        =   1441793
         _ExtentX        =   17806
         _ExtentY        =   4260
         _StockProps     =   79
         Caption         =   "Datos de la Cuenta Destino"
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
         Begin XtremeSuiteControls.FlatEdit txtP_Identificacion 
            Height          =   330
            Left            =   1680
            TabIndex        =   137
            Top             =   2040
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3831
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
         Begin XtremeSuiteControls.ComboBox cboCuenta 
            Height          =   315
            Left            =   5400
            TabIndex        =   138
            Top             =   2040
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
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
         Begin XtremeSuiteControls.PushButton btnCuenta 
            Height          =   315
            Left            =   9720
            TabIndex        =   139
            Top             =   2040
            Width           =   375
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
         Begin XtremeSuiteControls.ComboBox cboTipoId 
            Height          =   330
            Left            =   1680
            TabIndex        =   140
            Top             =   1680
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.ComboBox cboDivisa 
            Height          =   330
            Left            =   1680
            TabIndex        =   141
            Top             =   720
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtP_Correo 
            Height          =   315
            Left            =   5400
            TabIndex        =   142
            Top             =   720
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
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
         Begin XtremeSuiteControls.FlatEdit txtMtn_Girar 
            Height          =   330
            Left            =   1680
            TabIndex        =   144
            Top             =   360
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboBanco 
            Height          =   315
            Left            =   5400
            TabIndex        =   152
            Top             =   1680
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
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
         Begin XtremeSuiteControls.ComboBox cboEmite 
            Height          =   330
            Left            =   1680
            TabIndex        =   157
            Top             =   1080
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtP_Beneficiario 
            Height          =   330
            Left            =   5400
            TabIndex        =   155
            Top             =   360
            Width           =   4215
            _Version        =   1441793
            _ExtentX        =   7435
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtProveedorNombre 
            Height          =   315
            Left            =   5880
            TabIndex        =   159
            Top             =   1320
            Width           =   3735
            _Version        =   1441793
            _ExtentX        =   6583
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtProveedorId 
            Height          =   315
            Left            =   5400
            TabIndex        =   160
            Top             =   1320
            Width           =   510
            _Version        =   1441793
            _ExtentX        =   889
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
         Begin VB.Label lblProveedor 
            Alignment       =   1  'Right Justify
            Caption         =   "Proveedor"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   161
            Top             =   1320
            Width           =   975
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   12
            Left            =   -120
            TabIndex        =   158
            Top             =   1080
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Emite"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   5
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   11
            Left            =   3600
            TabIndex        =   156
            Top             =   360
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Beneficiario"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   7
            Left            =   3600
            TabIndex        =   154
            Top             =   1680
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Banco"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   10
            Left            =   3600
            TabIndex        =   151
            Top             =   720
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Correo"
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
         Begin XtremeSuiteControls.Label lblCuentaTitulo 
            Height          =   255
            Left            =   3600
            TabIndex        =   150
            Top             =   2040
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
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
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   8
            Left            =   -120
            TabIndex        =   149
            Top             =   2040
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Identificaci�n"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   6
            Left            =   -120
            TabIndex        =   148
            Top             =   720
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Divisa"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   5
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   4
            Left            =   -120
            TabIndex        =   147
            Top             =   1680
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo Id"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   5
            Left            =   -120
            TabIndex        =   145
            Top             =   360
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Giro"
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
      Begin XtremeSuiteControls.FlatEdit txtR_Cedula 
         Height          =   330
         Left            =   -67720
         TabIndex        =   200
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtR_Nombre 
         Height          =   315
         Left            =   -58840
         TabIndex        =   202
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.DateTimePicker dtpR_FechaNac 
         Height          =   315
         Left            =   -63280
         TabIndex        =   203
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
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
      Begin XtremeSuiteControls.FlatEdit txtR_Apellido1 
         Height          =   315
         Left            =   -67720
         TabIndex        =   204
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtR_Edad 
         Height          =   315
         Left            =   -61840
         TabIndex        =   210
         Top             =   960
         Visible         =   0   'False
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboR_Centro 
         Height          =   330
         Left            =   -67720
         TabIndex        =   212
         Top             =   1920
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.ComboBox cboR_GradoAnterior 
         Height          =   330
         Left            =   -63280
         TabIndex        =   214
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtR_CentroEducativo 
         Height          =   315
         Left            =   -63280
         TabIndex        =   216
         Top             =   1920
         Visible         =   0   'False
         Width           =   6735
         _Version        =   1441793
         _ExtentX        =   11880
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
      Begin XtremeSuiteControls.ComboBox cboR_Nivel 
         Height          =   330
         Left            =   -67720
         TabIndex        =   218
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.ComboBox cboR_Tipo 
         Height          =   330
         Left            =   -58720
         TabIndex        =   220
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.PushButton btnR_Guardar 
         Height          =   330
         Left            =   -58000
         TabIndex        =   222
         Top             =   3600
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Beneficios_Integral.frx":6BE5
      End
      Begin XtremeSuiteControls.FlatEdit txtR_Observaciones 
         Height          =   675
         Left            =   -67720
         TabIndex        =   223
         Top             =   2880
         Visible         =   0   'False
         Width           =   11175
         _Version        =   1441793
         _ExtentX        =   19711
         _ExtentY        =   1191
         _StockProps     =   77
         ForeColor       =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboR_Genero 
         Height          =   330
         Left            =   -58840
         TabIndex        =   226
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtR_Matematicas 
         Height          =   330
         Left            =   -67840
         TabIndex        =   229
         Top             =   4560
         Visible         =   0   'False
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboR_MateAplica 
         Height          =   330
         Left            =   -66880
         TabIndex        =   231
         Top             =   4560
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtR_Espanol 
         Height          =   330
         Left            =   -67840
         TabIndex        =   232
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboR_EspanolAplica 
         Height          =   330
         Left            =   -66880
         TabIndex        =   234
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtR_Ciencias 
         Height          =   330
         Left            =   -67840
         TabIndex        =   235
         Top             =   5520
         Visible         =   0   'False
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboR_CienciasAplica 
         Height          =   330
         Left            =   -66880
         TabIndex        =   237
         Top             =   5520
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtR_EstudiosSociales 
         Height          =   330
         Left            =   -67840
         TabIndex        =   238
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboR_EstudiosSocialesAplica 
         Height          =   330
         Left            =   -66880
         TabIndex        =   240
         Top             =   6000
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtR_Idioma 
         Height          =   330
         Left            =   -64120
         TabIndex        =   241
         Top             =   4560
         Visible         =   0   'False
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtR_Promedio 
         Height          =   330
         Left            =   -64120
         TabIndex        =   243
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboR_Merito_Deportivo 
         Height          =   330
         Left            =   -59200
         TabIndex        =   246
         Top             =   4680
         Visible         =   0   'False
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.ComboBox cboR_Merito_Cientifico 
         Height          =   330
         Left            =   -59200
         TabIndex        =   248
         Top             =   5280
         Visible         =   0   'False
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.ComboBox cboR_Merito_Artistico 
         Height          =   330
         Left            =   -59200
         TabIndex        =   250
         Top             =   5880
         Visible         =   0   'False
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   36
         Left            =   -59200
         TabIndex        =   253
         Top             =   5640
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "M�rito Art�stico"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   35
         Left            =   -59200
         TabIndex        =   252
         Top             =   5040
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "M�rito Cient�fico"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   34
         Left            =   -59200
         TabIndex        =   251
         Top             =   4440
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "M�rito Deportiivo"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   33
         Left            =   -61240
         TabIndex        =   249
         Top             =   5880
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fase Alcanzada"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   32
         Left            =   -61240
         TabIndex        =   247
         Top             =   5280
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nivel Cient�fico"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   31
         Left            =   -61240
         TabIndex        =   245
         Top             =   4680
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nivel Deportivo"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   30
         Left            =   -65320
         TabIndex        =   244
         Top             =   5040
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Promedio"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   29
         Left            =   -65320
         TabIndex        =   242
         Top             =   4560
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Idioma"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   28
         Left            =   -69760
         TabIndex        =   239
         Top             =   6000
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estudios Sociales"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   27
         Left            =   -69760
         TabIndex        =   236
         Top             =   5520
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ciencias"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   26
         Left            =   -69760
         TabIndex        =   233
         Top             =   5040
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Espa�ol"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   19
         Left            =   -69760
         TabIndex        =   230
         Top             =   4560
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Matem�ticas"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   375
         Index           =   4
         Left            =   -62200
         TabIndex        =   228
         Top             =   3960
         Visible         =   0   'False
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10610
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "M�ritos Especiales"
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
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   375
         Index           =   3
         Left            =   -70000
         TabIndex        =   227
         Top             =   3960
         Visible         =   0   'False
         Width           =   7815
         _Version        =   1441793
         _ExtentX        =   13785
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "M�rito Acad�mico"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblR_Observaciones 
         Height          =   255
         Left            =   -67720
         TabIndex        =   225
         Top             =   3600
         Visible         =   0   'False
         Width           =   4695
         _Version        =   1441793
         _ExtentX        =   8281
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "0 caracteres de 300 permitidos"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   26
         Left            =   -69760
         TabIndex        =   224
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Observaciones"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   25
         Left            =   -60760
         TabIndex        =   221
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Reconocimiento"
         BackColor       =   8421504
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   24
         Left            =   -69760
         TabIndex        =   219
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nivel"
         BackColor       =   8421504
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   23
         Left            =   -65320
         TabIndex        =   217
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Centro Educativo"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   22
         Left            =   -65320
         TabIndex        =   215
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Grado A�o anterior"
         BackColor       =   8421504
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   21
         Left            =   -69760
         TabIndex        =   213
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Centro"
         BackColor       =   8421504
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   20
         Left            =   -60760
         TabIndex        =   211
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Genero"
         BackColor       =   8421504
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   18
         Left            =   -60760
         TabIndex        =   209
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nombre"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   17
         Left            =   -65320
         TabIndex        =   208
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Segundo Apellido"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   16
         Left            =   -65320
         TabIndex        =   207
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha de Nacimiento"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   15
         Left            =   -69760
         TabIndex        =   206
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Primer Apellido"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   14
         Left            =   -69760
         TabIndex        =   201
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "C�dula del Estudiante"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   375
         Index           =   2
         Left            =   -70000
         TabIndex        =   199
         Top             =   360
         Visible         =   0   'False
         Width           =   13815
         _Version        =   1441793
         _ExtentX        =   24368
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Reconocimientos Mayra Soto: Datos del Estudiante"
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
      Begin XtremeSuiteControls.Label lblA_NotasQty 
         Height          =   255
         Left            =   -67840
         TabIndex        =   126
         Top             =   6000
         Visible         =   0   'False
         Width           =   4695
         _Version        =   1441793
         _ExtentX        =   8281
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "0 caracteres de 1200 permitidos"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   18
         Left            =   -68920
         TabIndex        =   124
         Top             =   5160
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Motivo"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   375
         Index           =   1
         Left            =   -67840
         TabIndex        =   123
         Top             =   1560
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Indique uno o m�s motivos de la solicitud"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   17
         Left            =   -69160
         TabIndex        =   119
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Motivo actual"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   16
         Left            =   -65320
         TabIndex        =   118
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Profesional Encargado"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   15
         Left            =   -69160
         TabIndex        =   116
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Categor�a"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   375
         Index           =   1
         Left            =   -69880
         TabIndex        =   105
         Top             =   2880
         Visible         =   0   'False
         Width           =   13575
         _Version        =   1441793
         _ExtentX        =   23945
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Lista de Sanciones Registradas"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   13
         Left            =   -68680
         TabIndex        =   100
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Id Sanci�n"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   12
         Left            =   -68680
         TabIndex        =   99
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha de Sanci�n"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   11
         Left            =   -68680
         TabIndex        =   98
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Motivo"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   10
         Left            =   -68680
         TabIndex        =   97
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Observaci�n"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   9
         Left            =   -68320
         TabIndex        =   67
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Observaci�n"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   375
         Index           =   0
         Left            =   -69880
         TabIndex        =   93
         Top             =   1800
         Visible         =   0   'False
         Width           =   13575
         _Version        =   1441793
         _ExtentX        =   23945
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Lista de Observaciones Registradas"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   6
         Left            =   -61600
         TabIndex        =   87
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Expediente"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   375
         Index           =   0
         Left            =   -59320
         TabIndex        =   85
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Estado Membres�a � Morosidad"
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
      Begin VB.Label Label18 
         Caption         =   "Nivel Acad�mico"
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
         Index           =   7
         Left            =   -67600
         TabIndex        =   81
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "Ocupaci�n"
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
         Index           =   6
         Left            =   -63280
         TabIndex        =   80
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "Lugar de Trabajo"
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
         Left            =   -67600
         TabIndex        =   78
         Top             =   2160
         Visible         =   0   'False
         Width           =   1695
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   -59320
         TabIndex        =   76
         Top             =   3120
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Lista de Tel�fonos"
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
      Begin VB.Label Label18 
         Caption         =   "Fecha Ingreso"
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
         Left            =   -62560
         TabIndex        =   74
         Top             =   1560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha Nacimiento"
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
         Left            =   -62560
         TabIndex        =   73
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "Genero:"
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
         Left            =   -67600
         TabIndex        =   72
         Top             =   1560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "Estado Civil"
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
         Index           =   0
         Left            =   -67600
         TabIndex        =   71
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   -62560
         TabIndex        =   70
         Top             =   600
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Apellido 2"
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
         Left            =   -65080
         TabIndex        =   69
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Apellido 1"
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
         Left            =   -67600
         TabIndex        =   68
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label18 
         Caption         =   "Pa�s Nacimiento:"
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
         Index           =   10
         Left            =   -67600
         TabIndex        =   60
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Nacionalidad:"
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
         Left            =   -64000
         TabIndex        =   59
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Direcci�n"
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
         Left            =   -67600
         TabIndex        =   58
         Top             =   4680
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Email No.2"
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
         Index           =   9
         Left            =   -67600
         TabIndex        =   57
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         Left            =   -67600
         TabIndex        =   56
         Top             =   4320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Email No.1"
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
         Left            =   -67600
         TabIndex        =   55
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   -69880
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Fecha"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   -58120
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   -59800
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   2
         Left            =   -63640
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nombre"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   -65440
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificaci�n"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   0
         Left            =   -68320
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Rango de Fechas"
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
   End
   Begin XtremeSuiteControls.PushButton btnNuevo 
      Height          =   330
      Left            =   10560
      TabIndex        =   1
      Top             =   1335
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   573
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmAF_Beneficios_Integral.frx":7316
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   330
      Left            =   11880
      TabIndex        =   2
      Top             =   1335
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   573
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_Beneficios_Integral.frx":7A36
   End
   Begin XtremeSuiteControls.PushButton btnBoleta 
      Height          =   330
      Left            =   12480
      TabIndex        =   3
      Top             =   1335
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   573
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_Beneficios_Integral.frx":8167
   End
   Begin XtremeSuiteControls.ComboBox cboTipoBeneficio 
      Height          =   345
      Left            =   5160
      TabIndex        =   5
      Top             =   1335
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   13080
      TabIndex        =   6
      Top             =   1335
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   573
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_Beneficios_Integral.frx":886E
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre_Completo 
      Height          =   330
      Left            =   5160
      TabIndex        =   21
      Top             =   480
      Width           =   5415
      _Version        =   1441793
      _ExtentX        =   9551
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   3480
      TabIndex        =   22
      Top             =   480
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
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
   Begin XtremeSuiteControls.FlatEdit txtBeneficioId 
      Height          =   345
      Left            =   3480
      TabIndex        =   109
      Top             =   1335
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   14
      Left            =   0
      TabIndex        =   115
      Top             =   0
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
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
      Height          =   135
      Index           =   1
      Left            =   5160
      TabIndex        =   25
      Top             =   240
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   238
      _StockProps     =   79
      Caption         =   "Nombre"
      ForeColor       =   16777215
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   135
      Index           =   0
      Left            =   3480
      TabIndex        =   24
      Top             =   240
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   238
      _StockProps     =   79
      Caption         =   "Identificaci�n"
      ForeColor       =   16777215
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
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   420
      Left            =   -120
      TabIndex        =   4
      Top             =   1305
      Width           =   13935
      _Version        =   1441793
      _ExtentX        =   24580
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "Beneficio activo: "
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
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   13935
   End
End
Attribute VB_Name = "frmAF_Beneficios_Integral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub btnAdjuntos_Click()

If txtCedula.Text <> "" Then
 gGA.Modulo = "CL_01"
 gGA.Llave_01 = txtCedula.Text
 gGA.Llave_02 = txtBeneficioId.Text
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End If

End Sub

Private Sub btnCuenta_Click()
Dim strSQL As String

On Error GoTo vError


GLOBALES.gTag = Trim(txtP_Identificacion.Text)
GLOBALES.gTag2 = "BENE"

frmCC_Cuentas_Bancarias.Show vbModal



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboBanco_Click()
If vPaso Or cboBanco.ListCount = 0 Or cboBanco.Text = "" Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & txtCedula.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub

Private Sub cboEmite_Click()

If vPaso Then Exit Sub
If cboEmite.ListCount = 0 Then Exit Sub


Dim pTipo As String


pTipo = fxTipoDocumento(cboEmite.Text)

lblProveedor.Top = lblCuentaTitulo.Top

If pTipo = "CP" Then
    lblCuentaTitulo.Visible = False
    lblProveedor.Visible = True
Else
    lblCuentaTitulo.Visible = True
    lblProveedor.Visible = False
End If

txtProveedorId.Top = lblProveedor.Top
txtProveedorNombre.Top = lblProveedor.Top

txtProveedorId.Visible = lblProveedor.Visible
txtProveedorNombre.Visible = lblProveedor.Visible


cboCuenta.Visible = lblCuentaTitulo.Visible

End Sub

Private Sub cboTipoBeneficio_Click()
If vPaso Then Exit Sub

tcMain(3).Visible = True
tcMain(4).Visible = False
tcMain(5).Visible = False
tcMain(9).Visible = False

tcBene_Aux(0).Visible = False
tcBene_Aux(1).Visible = False
tcBene_Aux(2).Visible = False



strSQL = "select * from afi_bene_categorias where cod_Categoria = '" _
       & cboTipoBeneficio.ItemData(cboTipoBeneficio.ListIndex) & "'"

Call OpenRecordSet(rs, strSQL)



Select Case True
    Case rs!I_FENA = 1  '"B_FENA" 'Fondo Nacional de Emergencias
        tcMain(3).Visible = True
        tcBene_Aux(2).Visible = True
        
        
    Case rs!I_SEPELIO = 1 '"B_SEPE" 'Sepelio
        tcMain(3).Visible = True
        tcBene_Aux(0).Visible = True
    
    Case rs!I_DESASTRES = 1 '"B_DESA" 'Desastre Natural o No Natural
        tcMain(3).Visible = True
        tcBene_Aux(1).Visible = True

    Case rs!I_APREMIANTE = 1 '"B_APRE" 'Apremiante
        tcMain(4).Visible = True
    
    Case rs!I_RECONOCIMIENTOS = 1 '"B_RECO" 'Reconocimientos Mayra Soto Hern�ndez
        tcMain(5).Visible = True
    
    Case rs!I_CRECE = 1 '"B_CRECE" 'Programa CRECE
        tcMain(9).Visible = True

End Select
rs.Close



tcMain(3).Selected = True

'Filtar los Tipos de Beneficios
vPaso = True
    strSQL = "select COD_BENEFICIO as 'IdX', DESCRIPCION as 'ItmX' " _
           & "  From afi_beneficios  " _
           & " Where COD_CATEGORIA = '" & cboTipoBeneficio.ItemData(cboTipoBeneficio.ListIndex) _
           & "'  and ESTADO = 'A' order by Descripcion"
    Call sbCbo_Llena_New(cboBeneficio, strSQL, False, True)
vPaso = False

'Llamar a datos del Beneficios
'Call cboBeneficio_Click

End Sub


Private Sub sbCargaCombos()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select cod_Pais as 'IdX', Descripcion as 'ItmX' from Paises" _
       & " where Activo = 1" _
       & " order by Omision desc, Descripcion asc"
Call sbCbo_Llena_New(cboPaisNac, strSQL, False, True)

strSQL = " select Catalogo_Id as 'IdX', Descripcion as 'ItmX' " _
       & " from AFI_CATALOGOS Where Tipo_Id = 3 order by Descripcion"
Call sbCbo_Llena_New(cboNivelAcademico, strSQL, False, True)


strSQL = "select cod_nacionalidad as 'IdX', Descripcion as 'ItmX' from Sys_nacionalidades" _
       & " where Activo = 1" _
       & " order by Omision desc, Descripcion asc"
Call sbCbo_Llena_New(cboNacionalidad, strSQL, False, True)

strSQL = "select Estado_Civil as 'IdX', Descripcion as 'ItmX' from SYS_ESTADO_CIVIL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoCivil, strSQL, False, True)

strSQL = "select ID_APT_CATEGORIA as 'IdX', DESCRIPCION as 'ItmX'" _
       & "  From AFI_BENE_APT_CATEGORIAS" _
       & "  Where ACTIVO = 1" _
       & " Order by ID_APT_CATEGORIA"
Call sbCbo_Llena_New(cboA_Categoria, strSQL, False, True)

strSQL = "select ID_PROFESIONAL as 'IdX', NOMBRE as 'ItmX'" _
       & "  From AFI_BENE_APT_PROFESIONALES" _
       & "  Where ACTIVO = 1" _
       & " Order by Nombre"
Call sbCbo_Llena_New(cboA_Profesional, strSQL, False, True)

'----Motivos
strSQL = "select COD_MOTIVO as 'IdX', DESCRIPCION as 'ItmX'" _
       & "  From AFI_BENE_MOTIVOS" _
       & "  Where ACTIVO = 1 order by DESCRIPCION"
Call sbCbo_Llena_New(cboA_Motivo, strSQL, False, True)


'TEMPORAL
strSQL = "select COD_MOTIVO , DESCRIPCION" _
       & "  From AFI_BENE_MOTIVOS" _
       & "  Where ACTIVO = 1 order by DESCRIPCION"
Call OpenRecordSet(rs, strSQL)
lswA_Motivos.ListItems.Clear
Do While Not rs.EOF
    Set itmX = lswA_Motivos.ListItems.Add(, , rs!Descripcion)
        itmX.Tag = rs!Cod_Motivo
rs.MoveNext
Loop
rs.Close

strSQL = "select COD_REQUISITO , DESCRIPCION" _
       & "  From AFI_BENE_REQUISITOS" _
       & "  Where ACTIVO = 1 order by DESCRIPCION"
Call OpenRecordSet(rs, strSQL)
lswRequisitos.ListItems.Clear
Do While Not rs.EOF
    Set itmX = lswRequisitos.ListItems.Add(, , rs!Descripcion)
        itmX.Tag = rs!Cod_Requisito
rs.MoveNext
Loop
rs.Close






'Opciones Limpias
cboNivelAcademico.AddItem ""

'Carga Tipos de Identificacion
vPaso = True
strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
vPaso = False


vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)

vPaso = False


'Call cboTipoId_Click


'----------------------------------------------------

vPaso = True

'Consulta todas las cuentas Bancarias
strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)


strSQL = "select COD_DIVISA AS 'IdX', DESCRIPCION as 'itmX' From vSys_Divisas" _
       & " Where DIVISA_LOCAL = 1"

Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)

vPaso = False


cboEmite.Clear
cboEmite.AddItem fxTipoDocumento("CK")
cboEmite.AddItem fxTipoDocumento("TE")
cboEmite.AddItem fxTipoDocumento("TS")
cboEmite.AddItem fxTipoDocumento("CP")
cboEmite.AddItem fxTipoDocumento("RC")
cboEmite.Text = fxTipoDocumento("TE")

Call cboEmite_Click

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbCargaCombos

End Sub

Private Sub txtA_Motivo_Change()
lblA_NotasQty.Caption = Len(txtA_Motivo.Text) & " caracteres de 1200 permitidos"
End Sub

Private Sub Form_Load()
 
 vModulo = 7
 
 imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture
 
 
cboSexo.AddItem "Masculino"
cboSexo.AddItem "Femenino"
cboSexo.AddItem "Otro"
cboSexo.Text = "Masculino"
 
 
 With lswConsulta.ColumnHeaders
    .Clear
    .Add , , "Id Expediente", 1200
    .Add , , "Id Beneficio", 1200
    .Add , , "Beneficio", 3000
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Estado", 1300, vbCenter
 End With
 
 
 With lswC_Sesiones.ColumnHeaders
    .Clear
    .Add , , "Sesi�n", 1200
    .Add , , "Asistencia", 1200, vbCenter
    .Add , , "Tarea", lswC_Sesiones.Width - 2600
 End With
 
 
 With lswA_Motivos.ColumnHeaders
    .Clear
    .Add , , "Motivo", lswA_Motivos.Width - 250
 End With
 lswA_Motivos.Checkboxes = True
 lswA_Motivos.HideColumnHeaders = True
 
 
 With lswPago.ColumnHeaders
    .Clear
    .Add , , "Identificaci�n", 1800
    .Add , , "Tipo", 700, vbCenter
    .Add , , "Nombre", 3000
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Emitir", 700, vbCenter
    .Add , , "Banco", 2500
    .Add , , "Cuenta", 1800
 End With
 
 With lswSanciones.ColumnHeaders
    .Clear
    .Add , , "Id", 800
    .Add , , "Motivo", 3500
    .Add , , "Activa?", 1100, vbCenter
    .Add , , "Fecha", 2500, vbCenter
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Notas", 1800
 End With

 
 With lswObservaciones.ColumnHeaders
    .Clear
    .Add , , "Id", 800
    .Add , , "Fecha", 2500, vbCenter
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Observaci�n", lswObservaciones.Width - 6100
 End With
 
 With lswRequisitos.ColumnHeaders
    .Clear
    .Add , , "Requisito", 4500
    .Add , , "Fecha", 2500, vbCenter
    .Add , , "Usuario", 2500, vbCenter
 End With
lswRequisitos.Checkboxes = True


With lswBitacora.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Fecha", 2500, vbCenter
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Detalle", lswBitacora.Width - 6100
 End With

vPaso = True
    strSQL = "select cod_categoria as 'IdX', descripcion as 'ItmX'" _
           & " From afi_bene_categorias  where Activo = 1 order by descripcion"
    Call sbCbo_Llena_New(cboTipoBeneficio, strSQL, False, True)
vPaso = False

 Call Formularios(Me)
 Call RefrescaTags(Me)


End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    
    gBusquedas.Col1Name = "C�dula Colilla"
    gBusquedas.Col2Name = "C�dula Real"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Filtro = " and EstadoActual = 'S'"
    gBusquedas.Columna = "Cedula"
    gBusquedas.Orden = "Cedula"
    
    
    frmBusquedas.Show vbModal
    
    txtCedula.Text = Trim(gBusquedas.Resultado)
    txtNombre_Completo.Text = gBusquedas.Resultado3
    
'    vCedula = txtCedula.Text
'    If Trim(txtCedula.Text) <> "" Then Call sbConsultaX(txtCedula)
    
End If
End Sub

Private Sub txtEstado_Notas_Change()
lblEstado_Notas.Caption = Len(txtEstado_Notas.Text) & " caracteres de 300 permitidos"
End Sub

Private Sub txtMnt_Notas_Change()
lblMnt_Notas.Caption = Len(txtMnt_Notas.Text) & " caracteres de 300 permitidos"
End Sub

Private Sub txtMtn_Girar_GotFocus()
On Error GoTo vError
    txtMtn_Girar.Text = CCur(txtMtn_Girar.Text)
vError:
End Sub


Private Sub txtMtn_Girar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboDivisa.SetFocus
End Sub

Private Sub txtMtn_Girar_LostFocus()
On Error GoTo vError
    txtMtn_Girar.Text = Format(CCur(txtMtn_Girar.Text), "Standard")
vError:
End Sub

Private Sub txtP_Identificacion_LostFocus()

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & txtP_Identificacion.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call OpenRecordSet(rs, strSQL)

cboCuenta.Clear
Do While Not rs.EOF
  cboCuenta.AddItem rs!IdX
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub
vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtProveedorId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedorNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal

  txtProveedorId.Text = gBusquedas.Resultado
  txtProveedorNombre.Text = gBusquedas.Resultado3
End If

End Sub



Private Sub txtProveedorNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMtn_Girar.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal

  txtProveedorId.Text = gBusquedas.Resultado
  txtProveedorNombre.Text = gBusquedas.Resultado3
End If

End Sub

