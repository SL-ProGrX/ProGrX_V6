VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_Reportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   8790
   ClientLeft      =   90
   ClientTop       =   480
   ClientWidth     =   9735
   Icon            =   "frmCR_Reportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   9735
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl tcPrincipal 
      Height          =   7575
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   9735
      _Version        =   1441793
      _ExtentX        =   17171
      _ExtentY        =   13361
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      PaintManager.Position=   2
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "Informes"
      Item(0).ControlCount=   23
      Item(0).Control(0)=   "ArbolExp"
      Item(0).Control(1)=   "tcFiltros"
      Item(0).Control(2)=   "lblSeguridad"
      Item(0).Control(3)=   "btnReporte"
      Item(0).Control(4)=   "chkFechas"
      Item(0).Control(5)=   "dtpInicio"
      Item(0).Control(6)=   "dtpCorte"
      Item(0).Control(7)=   "cboTipo"
      Item(0).Control(8)=   "cboFBase"
      Item(0).Control(9)=   "cboEOperacion"
      Item(0).Control(10)=   "cboEPersona"
      Item(0).Control(11)=   "cboOficina"
      Item(0).Control(12)=   "Label1(3)"
      Item(0).Control(13)=   "Label1(4)"
      Item(0).Control(14)=   "Label1(5)"
      Item(0).Control(15)=   "Label1(6)"
      Item(0).Control(16)=   "Label1(7)"
      Item(0).Control(17)=   "Label1(8)"
      Item(0).Control(18)=   "lblEstado"
      Item(0).Control(19)=   "Label1(11)"
      Item(0).Control(20)=   "Label1(26)"
      Item(0).Control(21)=   "cboESolicitud"
      Item(0).Control(22)=   "imgSeguridad"
      Item(1).Caption =   "Configuración"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "tcAux"
      Item(1).Control(1)=   "scTitulosTabs(0)"
      Item(2).Caption =   "Seguridad"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "tcAuxGrpAccs"
      Item(2).Control(1)=   "scTitulosTabs(1)"
      Begin MSComctlLib.TreeView ArbolExp 
         Height          =   7200
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   12700
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "imgArbol"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.TabControl tcFiltros 
         Height          =   4095
         Left            =   4200
         TabIndex        =   4
         Top             =   2520
         Width           =   5535
         _Version        =   1441793
         _ExtentX        =   9763
         _ExtentY        =   7223
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
         Item(0).ControlCount=   20
         Item(0).Control(0)=   "Label1(37)"
         Item(0).Control(1)=   "Label1(13)"
         Item(0).Control(2)=   "Label1(15)"
         Item(0).Control(3)=   "Label1(16)"
         Item(0).Control(4)=   "Label1(17)"
         Item(0).Control(5)=   "Label1(18)"
         Item(0).Control(6)=   "Label1(12)"
         Item(0).Control(7)=   "cboComite"
         Item(0).Control(8)=   "cboDestino"
         Item(0).Control(9)=   "cboRecurso"
         Item(0).Control(10)=   "cboInstitucion"
         Item(0).Control(11)=   "cboDeductora"
         Item(0).Control(12)=   "cboGarantia"
         Item(0).Control(13)=   "chkLineas"
         Item(0).Control(14)=   "txtCodigo"
         Item(0).Control(15)=   "txtDescripcion"
         Item(0).Control(16)=   "cboDivisa"
         Item(0).Control(17)=   "Label1(38)"
         Item(0).Control(18)=   "cboEspecial"
         Item(0).Control(19)=   "Label1(39)"
         Item(1).Caption =   "Adicionales"
         Item(1).ControlCount=   3
         Item(1).Control(0)=   "fraAdicional(2)"
         Item(1).Control(1)=   "fraAdicional(1)"
         Item(1).Control(2)=   "fraAdicional(0)"
         Item(2).Caption =   "F[1]"
         Item(2).ControlCount=   28
         Item(2).Control(0)=   "Label1(36)"
         Item(2).Control(1)=   "Label1(35)"
         Item(2).Control(2)=   "Label1(33)"
         Item(2).Control(3)=   "Label1(32)"
         Item(2).Control(4)=   "Label1(31)"
         Item(2).Control(5)=   "Label1(30)"
         Item(2).Control(6)=   "Label1(29)"
         Item(2).Control(7)=   "Label1(28)"
         Item(2).Control(8)=   "Label1(1)"
         Item(2).Control(9)=   "Label1(0)"
         Item(2).Control(10)=   "cboCobro"
         Item(2).Control(11)=   "cboProceso"
         Item(2).Control(12)=   "cboTiposTasas"
         Item(2).Control(13)=   "cboTipoOperacion"
         Item(2).Control(14)=   "cboAutorizaciones"
         Item(2).Control(15)=   "cboSigno(0)"
         Item(2).Control(16)=   "cboSigno(1)"
         Item(2).Control(17)=   "txtPlazoDesde"
         Item(2).Control(18)=   "txtPlazoHasta"
         Item(2).Control(19)=   "txtTasaDesde"
         Item(2).Control(20)=   "txtTasaHasta"
         Item(2).Control(21)=   "txtUltMov"
         Item(2).Control(22)=   "txtPrideduc"
         Item(2).Control(23)=   "chkPlazos"
         Item(2).Control(24)=   "chkTasas"
         Item(2).Control(25)=   "chkPriDeduc"
         Item(2).Control(26)=   "chkUltMov"
         Item(2).Control(27)=   "Label1(34)"
         Item(3).Caption =   "F[2]"
         Item(3).ControlCount=   21
         Item(3).Control(0)=   "Label18"
         Item(3).Control(1)=   "Label10"
         Item(3).Control(2)=   "Label9"
         Item(3).Control(3)=   "lblDepartamento"
         Item(3).Control(4)=   "lblSeccion"
         Item(3).Control(5)=   "Label1(25)"
         Item(3).Control(6)=   "cboZonas"
         Item(3).Control(7)=   "cboProvincia"
         Item(3).Control(8)=   "cboCanton"
         Item(3).Control(9)=   "cboDistrito"
         Item(3).Control(10)=   "txtDeptCodigo"
         Item(3).Control(11)=   "txtDeptDesc"
         Item(3).Control(12)=   "txtSecCodigo"
         Item(3).Control(13)=   "txtSecDesc"
         Item(3).Control(14)=   "chkProvincias"
         Item(3).Control(15)=   "chkCantones"
         Item(3).Control(16)=   "chkDistritos"
         Item(3).Control(17)=   "chkDepartamento"
         Item(3).Control(18)=   "chkSeccion"
         Item(3).Control(19)=   "cboUsuarios"
         Item(3).Control(20)=   "Label1(14)"
         Item(4).Caption =   "F[3]"
         Item(4).ControlCount=   13
         Item(4).Control(0)=   "Label1(20)"
         Item(4).Control(1)=   "Label1(21)"
         Item(4).Control(2)=   "Label1(22)"
         Item(4).Control(3)=   "Label1(23)"
         Item(4).Control(4)=   "Label1(24)"
         Item(4).Control(5)=   "cboProfesion"
         Item(4).Control(6)=   "cboSector"
         Item(4).Control(7)=   "cboSexo"
         Item(4).Control(8)=   "cboEstadoCivil"
         Item(4).Control(9)=   "cboCondicion"
         Item(4).Control(10)=   "txtEjecutivoId"
         Item(4).Control(11)=   "txtEjecutivoName"
         Item(4).Control(12)=   "Label1(40)"
         Begin XtremeSuiteControls.GroupBox fraAdicional 
            Height          =   1215
            Index           =   0
            Left            =   -69880
            TabIndex        =   5
            Top             =   360
            Visible         =   0   'False
            Width           =   5295
            _Version        =   1441793
            _ExtentX        =   9340
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Requisitos"
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
            Begin XtremeSuiteControls.ComboBox cboRequisitos 
               Height          =   330
               Left            =   1200
               TabIndex        =   6
               Top             =   360
               Width           =   4095
               _Version        =   1441793
               _ExtentX        =   7223
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
            Begin XtremeSuiteControls.ComboBox cboRequistoMarca 
               Height          =   330
               Left            =   1200
               TabIndex        =   7
               Top             =   720
               Width           =   4095
               _Version        =   1441793
               _ExtentX        =   7223
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
            Begin VB.Label Label1 
               Caption         =   "Requisitos"
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
               Left            =   120
               TabIndex        =   9
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Marca"
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
               Index           =   9
               Left            =   120
               TabIndex        =   8
               Top             =   720
               Width           =   732
            End
         End
         Begin XtremeSuiteControls.ComboBox cboComite 
            Height          =   330
            Left            =   1200
            TabIndex        =   10
            Top             =   1080
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
         Begin XtremeSuiteControls.ComboBox cboDestino 
            Height          =   330
            Left            =   1200
            TabIndex        =   11
            Top             =   2160
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
         Begin XtremeSuiteControls.ComboBox cboRecurso 
            Height          =   330
            Left            =   1200
            TabIndex        =   12
            Top             =   2520
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
         Begin XtremeSuiteControls.ComboBox cboInstitucion 
            Height          =   330
            Left            =   1200
            TabIndex        =   13
            Top             =   2880
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
         Begin XtremeSuiteControls.ComboBox cboDeductora 
            Height          =   330
            Left            =   1200
            TabIndex        =   14
            Top             =   3240
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
         Begin XtremeSuiteControls.ComboBox cboGarantia 
            Height          =   330
            Left            =   1200
            TabIndex        =   15
            Top             =   720
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
         Begin XtremeSuiteControls.CheckBox chkLineas 
            Height          =   255
            Left            =   4440
            TabIndex        =   16
            Top             =   1440
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
            _ExtentY        =   450
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtCodigo 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   1800
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
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
         Begin XtremeSuiteControls.FlatEdit txtDescripcion 
            Height          =   315
            Left            =   1200
            TabIndex        =   18
            Top             =   1800
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
         Begin XtremeSuiteControls.ComboBox cboCobro 
            Height          =   315
            Left            =   -69640
            TabIndex        =   19
            Top             =   720
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.ComboBox cboProceso 
            Height          =   315
            Left            =   -67960
            TabIndex        =   20
            Top             =   720
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
         Begin XtremeSuiteControls.ComboBox cboTiposTasas 
            Height          =   315
            Left            =   -69640
            TabIndex        =   21
            Top             =   1560
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1441793
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.ComboBox cboTipoOperacion 
            Height          =   315
            Left            =   -67960
            TabIndex        =   22
            Top             =   1560
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
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
         Begin XtremeSuiteControls.ComboBox cboAutorizaciones 
            Height          =   315
            Left            =   -66400
            TabIndex        =   23
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
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
         Begin XtremeSuiteControls.ComboBox cboSigno 
            Height          =   312
            Index           =   0
            Left            =   -68320
            TabIndex        =   24
            Top             =   3120
            Visible         =   0   'False
            Width           =   852
            _Version        =   1441793
            _ExtentX        =   1508
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
         Begin XtremeSuiteControls.ComboBox cboSigno 
            Height          =   312
            Index           =   1
            Left            =   -68320
            TabIndex        =   25
            Top             =   3480
            Visible         =   0   'False
            Width           =   852
            _Version        =   1441793
            _ExtentX        =   1508
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
         Begin XtremeSuiteControls.FlatEdit txtPlazoDesde 
            Height          =   312
            Left            =   -68320
            TabIndex        =   26
            Top             =   2400
            Visible         =   0   'False
            Width           =   852
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
         Begin XtremeSuiteControls.FlatEdit txtPlazoHasta 
            Height          =   312
            Left            =   -67480
            TabIndex        =   27
            Top             =   2400
            Visible         =   0   'False
            Width           =   852
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
         Begin XtremeSuiteControls.FlatEdit txtTasaDesde 
            Height          =   312
            Left            =   -68320
            TabIndex        =   28
            Top             =   2760
            Visible         =   0   'False
            Width           =   852
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
         Begin XtremeSuiteControls.FlatEdit txtTasaHasta 
            Height          =   312
            Left            =   -67480
            TabIndex        =   29
            Top             =   2760
            Visible         =   0   'False
            Width           =   852
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
         Begin XtremeSuiteControls.FlatEdit txtUltMov 
            Height          =   312
            Left            =   -67480
            TabIndex        =   30
            Top             =   3480
            Visible         =   0   'False
            Width           =   852
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
         Begin XtremeSuiteControls.FlatEdit txtPrideduc 
            Height          =   312
            Left            =   -67480
            TabIndex        =   31
            Top             =   3120
            Visible         =   0   'False
            Width           =   852
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
         Begin XtremeSuiteControls.CheckBox chkPlazos 
            Height          =   252
            Left            =   -66400
            TabIndex        =   32
            Top             =   2400
            Visible         =   0   'False
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todos"
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
         Begin XtremeSuiteControls.CheckBox chkTasas 
            Height          =   252
            Left            =   -66400
            TabIndex        =   33
            Top             =   2760
            Visible         =   0   'False
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
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
         Begin XtremeSuiteControls.CheckBox chkPriDeduc 
            Height          =   252
            Left            =   -66400
            TabIndex        =   34
            Top             =   3120
            Visible         =   0   'False
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
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
         Begin XtremeSuiteControls.CheckBox chkUltMov 
            Height          =   252
            Left            =   -66400
            TabIndex        =   35
            Top             =   3480
            Visible         =   0   'False
            Width           =   972
            _Version        =   1441793
            _ExtentX        =   1714
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
         Begin XtremeSuiteControls.ComboBox cboZonas 
            Height          =   330
            Left            =   -68920
            TabIndex        =   36
            Top             =   840
            Visible         =   0   'False
            Width           =   3495
            _Version        =   1441793
            _ExtentX        =   6165
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
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   330
            Left            =   -68920
            TabIndex        =   37
            Top             =   1200
            Visible         =   0   'False
            Width           =   3495
            _Version        =   1441793
            _ExtentX        =   6165
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
         Begin XtremeSuiteControls.ComboBox cboCanton 
            Height          =   330
            Left            =   -68920
            TabIndex        =   38
            Top             =   1560
            Visible         =   0   'False
            Width           =   3495
            _Version        =   1441793
            _ExtentX        =   6165
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
         Begin XtremeSuiteControls.ComboBox cboDistrito 
            Height          =   330
            Left            =   -68920
            TabIndex        =   39
            Top             =   1920
            Visible         =   0   'False
            Width           =   3495
            _Version        =   1441793
            _ExtentX        =   6165
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
         Begin XtremeSuiteControls.FlatEdit txtDeptCodigo 
            Height          =   315
            Left            =   -69880
            TabIndex        =   40
            Top             =   2760
            Visible         =   0   'False
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
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
         Begin XtremeSuiteControls.FlatEdit txtDeptDesc 
            Height          =   315
            Left            =   -68920
            TabIndex        =   41
            Top             =   2760
            Visible         =   0   'False
            Width           =   4335
            _Version        =   1441793
            _ExtentX        =   7646
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
         Begin XtremeSuiteControls.FlatEdit txtSecCodigo 
            Height          =   315
            Left            =   -69880
            TabIndex        =   42
            Top             =   3480
            Visible         =   0   'False
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1720
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
         Begin XtremeSuiteControls.FlatEdit txtSecDesc 
            Height          =   315
            Left            =   -68920
            TabIndex        =   43
            Top             =   3480
            Visible         =   0   'False
            Width           =   4335
            _Version        =   1441793
            _ExtentX        =   7646
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
         Begin XtremeSuiteControls.CheckBox chkProvincias 
            Height          =   255
            Left            =   -65320
            TabIndex        =   44
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1503
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
         Begin XtremeSuiteControls.CheckBox chkCantones 
            Height          =   255
            Left            =   -65320
            TabIndex        =   45
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1503
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todos"
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
         Begin XtremeSuiteControls.CheckBox chkDistritos 
            Height          =   255
            Left            =   -65320
            TabIndex        =   46
            Top             =   1920
            Visible         =   0   'False
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1503
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todos"
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
         Begin XtremeSuiteControls.CheckBox chkDepartamento 
            Height          =   255
            Left            =   -65560
            TabIndex        =   47
            Top             =   2520
            Visible         =   0   'False
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todos"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkSeccion 
            Height          =   255
            Left            =   -65560
            TabIndex        =   48
            Top             =   3240
            Visible         =   0   'False
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1714
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.ComboBox cboProfesion 
            Height          =   330
            Left            =   -68680
            TabIndex        =   49
            Top             =   480
            Visible         =   0   'False
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboSector 
            Height          =   330
            Left            =   -68680
            TabIndex        =   50
            Top             =   840
            Visible         =   0   'False
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboSexo 
            Height          =   330
            Left            =   -68680
            TabIndex        =   51
            Top             =   1440
            Visible         =   0   'False
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboEstadoCivil 
            Height          =   330
            Left            =   -68680
            TabIndex        =   52
            Top             =   1800
            Visible         =   0   'False
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboCondicion 
            Height          =   330
            Left            =   -68680
            TabIndex        =   53
            Top             =   2400
            Visible         =   0   'False
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.GroupBox fraAdicional 
            Height          =   1215
            Index           =   1
            Left            =   -69880
            TabIndex        =   54
            Top             =   1680
            Visible         =   0   'False
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Causas"
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
            Begin XtremeSuiteControls.ComboBox cboCausasTipos 
               Height          =   330
               Left            =   1200
               TabIndex        =   55
               Top             =   360
               Width           =   4095
               _Version        =   1441793
               _ExtentX        =   7223
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
            Begin XtremeSuiteControls.ComboBox cboCausas 
               Height          =   330
               Left            =   1200
               TabIndex        =   56
               Top             =   720
               Width           =   4095
               _Version        =   1441793
               _ExtentX        =   7223
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
            Begin VB.Label Label1 
               Caption         =   "Tipos"
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
               Index           =   19
               Left            =   120
               TabIndex        =   58
               Top             =   360
               Width           =   732
            End
            Begin VB.Label Label1 
               Caption         =   "Causas"
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
               Index           =   10
               Left            =   120
               TabIndex        =   57
               Top             =   720
               Width           =   732
            End
         End
         Begin XtremeSuiteControls.GroupBox fraAdicional 
            Height          =   1215
            Index           =   2
            Left            =   -69880
            TabIndex        =   59
            Top             =   3000
            Visible         =   0   'False
            Width           =   5415
            _Version        =   1441793
            _ExtentX        =   9551
            _ExtentY        =   2143
            _StockProps     =   79
            Caption         =   "Cortes"
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
            Begin XtremeSuiteControls.ComboBox cboCorte 
               Height          =   330
               Left            =   1200
               TabIndex        =   60
               Top             =   360
               Width           =   4095
               _Version        =   1441793
               _ExtentX        =   7223
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
            Begin VB.Label Label1 
               Caption         =   "Corte"
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
               Index           =   27
               Left            =   120
               TabIndex        =   61
               Top             =   360
               Width           =   852
            End
         End
         Begin XtremeSuiteControls.ComboBox cboDivisa 
            Height          =   330
            Left            =   1200
            TabIndex        =   62
            Top             =   360
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
         Begin XtremeSuiteControls.ComboBox cboUsuarios 
            Height          =   330
            Left            =   -68920
            TabIndex        =   63
            Top             =   480
            Visible         =   0   'False
            Width           =   3495
            _Version        =   1441793
            _ExtentX        =   6165
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
         Begin XtremeSuiteControls.ComboBox cboEspecial 
            Height          =   330
            Left            =   1200
            TabIndex        =   64
            Top             =   3720
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
         Begin XtremeSuiteControls.FlatEdit txtEjecutivoId 
            Height          =   315
            Left            =   -69880
            TabIndex        =   65
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   3435
            Visible         =   0   'False
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1714
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEjecutivoName 
            Height          =   315
            Left            =   -68920
            TabIndex        =   66
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   3435
            Visible         =   0   'False
            Width           =   4335
            _Version        =   1441793
            _ExtentX        =   7646
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
            TabIndex        =   99
            Top             =   3240
            Width           =   1095
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
            TabIndex        =   98
            Top             =   2160
            Width           =   1095
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
            TabIndex        =   97
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Garantía"
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
            Index           =   16
            Left            =   120
            TabIndex        =   96
            Top             =   720
            Width           =   732
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
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
            Height          =   252
            Index           =   17
            Left            =   120
            TabIndex        =   95
            Top             =   1080
            Width           =   732
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
            TabIndex        =   94
            Top             =   2880
            Width           =   1095
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
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   12
            Left            =   120
            TabIndex        =   93
            Top             =   1485
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Ult.Mov."
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
            Left            =   -69640
            TabIndex        =   92
            Top             =   3480
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.Label Label1 
            Caption         =   "Autorizaciones"
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
            Index           =   35
            Left            =   -66400
            TabIndex        =   91
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   33
            Left            =   -67480
            TabIndex        =   90
            Top             =   2160
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   32
            Left            =   -68320
            TabIndex        =   89
            Top             =   2160
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.Label Label1 
            Caption         =   "Tipos de Tasas"
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
            Index           =   31
            Left            =   -69640
            TabIndex        =   88
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Tasas"
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
            Index           =   30
            Left            =   -69640
            TabIndex        =   87
            Top             =   2760
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.Label Label1 
            Caption         =   "Plazos"
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
            Index           =   29
            Left            =   -69640
            TabIndex        =   86
            Top             =   2400
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Operación"
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
            Index           =   28
            Left            =   -67960
            TabIndex        =   85
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Proceso"
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
            Left            =   -67960
            TabIndex        =   84
            Top             =   480
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Cobro en"
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
            Left            =   -69640
            TabIndex        =   83
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Primer Deduc."
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
            Index           =   34
            Left            =   -69640
            TabIndex        =   82
            Top             =   3120
            Visible         =   0   'False
            Width           =   1332
         End
         Begin VB.Label Label18 
            Caption         =   "Distrito"
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
            Left            =   -69880
            TabIndex        =   81
            Top             =   1920
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia"
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
            Left            =   -69880
            TabIndex        =   80
            Top             =   1200
            Visible         =   0   'False
            Width           =   852
         End
         Begin VB.Label Label9 
            Caption         =   "Cantón"
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
            Left            =   -69880
            TabIndex        =   79
            Top             =   1560
            Visible         =   0   'False
            Width           =   612
         End
         Begin VB.Label lblDepartamento 
            Caption         =   "Departamento"
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
            Left            =   -69880
            TabIndex        =   78
            Top             =   2520
            Visible         =   0   'False
            Width           =   2292
         End
         Begin VB.Label lblSeccion 
            Caption         =   "Sección"
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
            Left            =   -69880
            TabIndex        =   77
            Top             =   3240
            Visible         =   0   'False
            Width           =   2292
         End
         Begin VB.Label Label1 
            Caption         =   "Zonas"
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
            Index           =   25
            Left            =   -69880
            TabIndex        =   76
            Top             =   840
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   20
            Left            =   -69880
            TabIndex        =   75
            Top             =   1440
            Visible         =   0   'False
            Width           =   972
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado Civil"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   21
            Left            =   -69880
            TabIndex        =   74
            Top             =   1800
            Visible         =   0   'False
            Width           =   972
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Condición Laboral"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   492
            Index           =   22
            Left            =   -69880
            TabIndex        =   73
            Top             =   2400
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.Label Label1 
            Caption         =   "Profesión"
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
            Index           =   23
            Left            =   -69880
            TabIndex        =   72
            Top             =   480
            Visible         =   0   'False
            Width           =   972
         End
         Begin VB.Label Label1 
            Caption         =   "Sector"
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
            Index           =   24
            Left            =   -69880
            TabIndex        =   71
            Top             =   840
            Visible         =   0   'False
            Width           =   972
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
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
            Height          =   252
            Index           =   38
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   732
         End
         Begin VB.Label Label1 
            Caption         =   "Usuarios"
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
            Index           =   14
            Left            =   -69880
            TabIndex        =   69
            Top             =   480
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Especial"
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
            Index           =   39
            Left            =   120
            TabIndex        =   68
            Top             =   3720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ejecutivo o Colocador del Crédito"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   40
            Left            =   -69880
            TabIndex        =   67
            Top             =   3240
            Visible         =   0   'False
            Width           =   3135
         End
      End
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   495
         Left            =   8040
         TabIndex        =   101
         Top             =   6720
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Picture         =   "frmCR_Reportes.frx":08CA
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.CheckBox chkFechas 
         Height          =   255
         Left            =   8520
         TabIndex        =   102
         Top             =   840
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   6360
         TabIndex        =   103
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   315
         Left            =   8280
         TabIndex        =   104
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   5400
         TabIndex        =   105
         Top             =   120
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
      Begin XtremeSuiteControls.ComboBox cboFBase 
         Height          =   315
         Left            =   6360
         TabIndex        =   106
         Top             =   840
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.ComboBox cboEOperacion 
         Height          =   330
         Left            =   6360
         TabIndex        =   107
         Top             =   1200
         Width           =   3255
         _Version        =   1441793
         _ExtentX        =   5741
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
      Begin XtremeSuiteControls.ComboBox cboEPersona 
         Height          =   330
         Left            =   6360
         TabIndex        =   108
         Top             =   1560
         Width           =   3255
         _Version        =   1441793
         _ExtentX        =   5741
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
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   330
         Left            =   5400
         TabIndex        =   109
         Top             =   1920
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
      Begin XtremeSuiteControls.ComboBox cboESolicitud 
         Height          =   330
         Left            =   6360
         TabIndex        =   119
         Top             =   1200
         Width           =   3255
         _Version        =   1441793
         _ExtentX        =   5741
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
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   6735
         Left            =   -70000
         TabIndex        =   120
         Top             =   480
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   11880
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   32
         PaintManager.BoldSelected=   -1  'True
         ItemCount       =   3
         Item(0).Caption =   "Grupos"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "vGrid"
         Item(0).Control(1)=   "Label2(1)"
         Item(1).Caption =   "Miembros"
         Item(1).ControlCount=   4
         Item(1).Control(0)=   "cboMiembros"
         Item(1).Control(1)=   "lswMiembros"
         Item(1).Control(2)=   "Label2(2)"
         Item(1).Control(3)=   "Label2(3)"
         Item(2).Caption =   "Informes"
         Item(2).ControlCount=   4
         Item(2).Control(0)=   "txtReportes"
         Item(2).Control(1)=   "vGridRep"
         Item(2).Control(2)=   "imgAddRep"
         Item(2).Control(3)=   "Label4"
         Begin XtremeSuiteControls.ListView lswMiembros 
            Height          =   5775
            Left            =   -68080
            TabIndex        =   138
            Top             =   840
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11245
            _ExtentY        =   10186
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
         Begin FPSpreadADO.fpSpread vGrid 
            Height          =   6135
            Left            =   1680
            TabIndex        =   121
            Top             =   480
            Width           =   6615
            _Version        =   524288
            _ExtentX        =   11668
            _ExtentY        =   10821
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
            MaxCols         =   496
            ScrollBars      =   2
            SpreadDesigner  =   "frmCR_Reportes.frx":0FD1
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridRep 
            Height          =   5655
            Left            =   -69640
            TabIndex        =   134
            Top             =   960
            Visible         =   0   'False
            Width           =   8895
            _Version        =   524288
            _ExtentX        =   15690
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
            MaxCols         =   496
            ScrollBars      =   2
            SpreadDesigner  =   "frmCR_Reportes.frx":14D8
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtReportes 
            Height          =   375
            Left            =   -67240
            TabIndex        =   135
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   661
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
            PasswordChar    =   "*"
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboMiembros 
            Height          =   330
            Left            =   -68080
            TabIndex        =   137
            Top             =   480
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11245
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   375
            Left            =   -69280
            TabIndex        =   136
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Clave de Edición (Admin)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin VB.Image imgAddRep 
            Height          =   375
            Left            =   -61600
            Picture         =   "frmCR_Reportes.frx":1B37
            Stretch         =   -1  'True
            ToolTipText     =   "Agregar & Actualizar lista de Reportes"
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupo de Usuarios"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Index           =   1
            Left            =   600
            TabIndex        =   124
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   2
            Left            =   -69280
            TabIndex        =   123
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Miembros"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   3
            Left            =   -69280
            TabIndex        =   122
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin XtremeSuiteControls.TabControl tcAuxGrpAccs 
         Height          =   6735
         Left            =   -70000
         TabIndex        =   125
         Top             =   480
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   11880
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   32
         PaintManager.BoldSelected=   -1  'True
         ItemCount       =   3
         Item(0).Caption =   "Grupos"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "vGridGrpAccss"
         Item(0).Control(1)=   "Label2(6)"
         Item(1).Caption =   "Miembros"
         Item(1).ControlCount=   4
         Item(1).Control(0)=   "cboGrpAccssM"
         Item(1).Control(1)=   "lswGrpAccssM"
         Item(1).Control(2)=   "Label2(8)"
         Item(1).Control(3)=   "Label2(9)"
         Item(2).Caption =   "Informes Autorizados"
         Item(2).ControlCount=   4
         Item(2).Control(0)=   "cboGrpAccssR"
         Item(2).Control(1)=   "lswGrpAccssR"
         Item(2).Control(2)=   "Label2(10)"
         Item(2).Control(3)=   "Label2(11)"
         Begin XtremeSuiteControls.ListView lswGrpAccssR 
            Height          =   5775
            Left            =   -68320
            TabIndex        =   141
            Top             =   840
            Visible         =   0   'False
            Width           =   6615
            _Version        =   1441793
            _ExtentX        =   11668
            _ExtentY        =   10186
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
         Begin XtremeSuiteControls.ListView lswGrpAccssM 
            Height          =   5775
            Left            =   -68200
            TabIndex        =   139
            Top             =   840
            Visible         =   0   'False
            Width           =   6495
            _Version        =   1441793
            _ExtentX        =   11456
            _ExtentY        =   10186
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
         Begin FPSpreadADO.fpSpread vGridGrpAccss 
            Height          =   6135
            Left            =   1320
            TabIndex        =   126
            Top             =   480
            Width           =   7575
            _Version        =   524288
            _ExtentX        =   13361
            _ExtentY        =   10821
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
            MaxCols         =   496
            ScrollBars      =   2
            SpreadDesigner  =   "frmCR_Reportes.frx":2247
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.ComboBox cboGrpAccssM 
            Height          =   330
            Left            =   -68200
            TabIndex        =   140
            Top             =   480
            Visible         =   0   'False
            Width           =   6495
            _Version        =   1441793
            _ExtentX        =   11456
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
         Begin XtremeSuiteControls.ComboBox cboGrpAccssR 
            Height          =   330
            Left            =   -68320
            TabIndex        =   142
            Top             =   480
            Visible         =   0   'False
            Width           =   6615
            _Version        =   1441793
            _ExtentX        =   11668
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupos de Acceso"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Index           =   6
            Left            =   240
            TabIndex        =   131
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   8
            Left            =   -69400
            TabIndex        =   130
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Miembros"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   9
            Left            =   -69400
            TabIndex        =   129
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   10
            Left            =   -69520
            TabIndex        =   128
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Reportes"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   11
            Left            =   -69520
            TabIndex        =   127
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulosTabs 
         Height          =   375
         Index           =   1
         Left            =   -70000
         TabIndex        =   133
         Top             =   0
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Grupos de Acceso"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   4210752
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulosTabs 
         Height          =   375
         Index           =   0
         Left            =   -70000
         TabIndex        =   132
         Top             =   0
         Visible         =   0   'False
         Width           =   9735
         _Version        =   1441793
         _ExtentX        =   17171
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Grupos de Trabajo"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   4210752
      End
      Begin VB.Image imgSeguridad 
         Height          =   255
         Left            =   4320
         Picture         =   "frmCR_Reportes.frx":27BF
         Stretch         =   -1  'True
         ToolTipText     =   "Requiere Grupo de Acceso Autorizado"
         Top             =   6840
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Oficina"
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
         Index           =   26
         Left            =   4320
         TabIndex        =   118
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Persona"
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
         Height          =   315
         Index           =   11
         Left            =   5400
         TabIndex        =   117
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Solicitud"
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
         Height          =   315
         Left            =   5400
         TabIndex        =   116
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Estados"
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
         Index           =   8
         Left            =   4320
         TabIndex        =   115
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Base"
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
         Height          =   315
         Index           =   7
         Left            =   5400
         TabIndex        =   114
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Corte"
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
         Height          =   315
         Index           =   6
         Left            =   7560
         TabIndex        =   113
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Inicio"
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
         Height          =   315
         Index           =   5
         Left            =   5400
         TabIndex        =   112
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Fechas"
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
         Left            =   4320
         TabIndex        =   111
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Reporte"
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
         Left            =   4320
         TabIndex        =   110
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblSeguridad 
         Caption         =   "[ Requiere Grupo de Acceso Autorizado ]"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4680
         TabIndex        =   100
         Top             =   6870
         Width           =   3255
      End
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_Reportes.frx":2EBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_Reportes.frx":971D
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_Reportes.frx":983B
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_Reportes.frx":9965
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_Reportes.frx":9A8B
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_Reportes.frx":9B99
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_Reportes.frx":9CA6
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_Reportes.frx":9DBF
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_Reportes.frx":9EED
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_Reportes.frx":9FFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblReporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   2520
      TabIndex        =   1
      Top             =   540
      Width           =   6972
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de Crédito"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   300
      Width           =   4572
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   9852
   End
End
Attribute VB_Name = "frmCR_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mModoSif As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub sbInicializa()

Me.MousePointer = vbHourglass

tcPrincipal.Item(0).Selected = True

tcFiltros.Item(0).Selected = True
tcFiltros.Item(1).Enabled = False

cboEspecial.Clear
cboEspecial.AddItem "TODOS"
cboEspecial.AddItem "Cartera Interna"
cboEspecial.AddItem "Cartera Administrada"
'cboEspecial.AddItem "Recaudos & Retenciones"
cboEspecial.Text = "TODOS"


imgSeguridad.Visible = False
lblSeguridad.Visible = imgSeguridad.Visible

lblReporte.Tag = ""
lblReporte.Caption = ">>> Seleccione Un Reporte <<<"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
chkFechas.Value = vbUnchecked
chkLineas.Value = vbChecked


cboProceso.Clear
cboProceso.AddItem "Normal"
cboProceso.AddItem "Traspaso Deuda"
cboProceso.AddItem "Cobro Judicial"
cboProceso.AddItem "TODOS"
cboProceso.Text = "TODOS"

cboCobro.Clear
cboCobro.AddItem "Cajas"
cboCobro.AddItem "Planilla"
cboCobro.AddItem "TODOS"
cboCobro.Text = "TODOS"


cboFBase.Clear
cboFBase.AddItem "Solicitud"
cboFBase.AddItem "Resolución"
cboFBase.AddItem "Formalización"
cboFBase.AddItem "Desembolso"
cboFBase.AddItem "Ultimo Mov."
cboFBase.Text = "Solicitud"

cboTipo.Clear
cboTipo.AddItem "Detalle"
cboTipo.AddItem "Resumen"
cboTipo.Text = "Detalle"

vPaso = True
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)
vPaso = False

strSQL = "select rtrim(Garantia) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from crd_garantia_tipos order by descripcion"
Call sbCbo_Llena_New(cboGarantia, strSQL, True, False)

cboESolicitud.Clear
cboESolicitud.AddItem "Recibida"
cboESolicitud.AddItem "Pendiente"
cboESolicitud.AddItem "Formalizada"
cboESolicitud.AddItem "Nula"
cboESolicitud.AddItem "Aprobada"
cboESolicitud.AddItem "Denegada"
cboESolicitud.AddItem "Todas"
cboESolicitud.Text = "Recibida"

cboEOperacion.Clear
cboEOperacion.AddItem "Activa"
cboEOperacion.AddItem "Cancelada"
cboEOperacion.AddItem "Nulas"
cboEOperacion.AddItem "Todas (Activas/Canceladas)"
cboEOperacion.Text = "Activa"


strSQL = "select rtrim(cod_oficina) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from SIF_Oficinas order by descripcion"
Call sbCbo_Llena_New(cboOficina, strSQL, True, False)


strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from afi_estados_persona order by descripcion"
Call sbCbo_Llena_New(cboEPersona, strSQL, True, False)
'Item Adicional
    cboEPersona.AddItem "Ex.Socios"
    cboEPersona.ItemData(cboEPersona.ListCount - 1) = "X"

strSQL = "select rtrim(descripcion) as 'Itmx', id_comite as 'Idx'" _
       & " from comites order by descripcion"
Call sbCbo_Llena_New(cboComite, strSQL, True, True)

strSQL = "select rtrim(cod_grupo) as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  crd_grupos"
Call sbCbo_Llena_New(cboUsuarios, strSQL, True, True)


strSQL = "select COD_DIVISA AS 'IdX', DESCRIPCION as 'ItmX'" _
       & " From vSys_Divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)

'Instituciones
vPaso = True
    strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
           & " from instituciones order by descripcion"
    Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)
vPaso = False

'Adicionales
'Causas de Pendientes y Denegadas
vPaso = True
    cboCausasTipos.Clear
    cboCausasTipos.AddItem "Pendientes"
    cboCausasTipos.AddItem "Denegadas"
    cboCausasTipos.Text = "Pendientes"
vPaso = False

'Requisitos Tipo de Marca
cboRequistoMarca.Clear
cboRequistoMarca.AddItem "Cumple"
cboRequistoMarca.AddItem "No Cumple"
cboRequistoMarca.AddItem "En Blanco"
cboRequistoMarca.Text = "Cumple"

'Cortes (Frecuencia)
cboCorte.Clear
cboCorte.AddItem "Diario"
cboCorte.AddItem "Semanal"
cboCorte.AddItem "Mensual"
cboCorte.AddItem "Trimestral"
cboCorte.AddItem "Semestral"
cboCorte.AddItem "Anual"
cboCorte.Text = "Mensual"

'Adicionales y Distribuicion
strSQL = "select OBJECT_ID('UPROGRAMATICA') as Resultado"
Call OpenRecordSet(rs, strSQL)
If IsNull(rs!Resultado) Then
  mModoSif = True
  lblDepartamento.Caption = "Departamento"
  lblSeccion.Caption = "Sección"
Else
  mModoSif = False
  lblDepartamento.Caption = "Unidad Programatica"
  lblSeccion.Caption = "Unidad de Trabajo"
End If
rs.Close

strSQL = "select COD_PROFESION as 'Idx',descripcion as 'ItmX' from AFI_PROFESIONES order by descripcion"
Call sbCbo_Llena_New(cboProfesion, strSQL, True, True)

strSQL = "select COD_SECTOR as 'Idx',descripcion as 'ItmX' from AFI_SECTORES order by descripcion"
Call sbCbo_Llena_New(cboSector, strSQL, True, True)

strSQL = "select rtrim(COD_ZONA) as 'IdX', rtrim(descripcion) as 'ItmX' from AFI_ZONAS order by descripcion"
Call sbCbo_Llena_New(cboZonas, strSQL, True, True)

cboTipoOperacion.Clear
cboTipoOperacion.AddItem "TODAS"
cboTipoOperacion.AddItem "Originales"
cboTipoOperacion.AddItem "Derivadas"
cboTipoOperacion.Text = "TODAS"

cboTiposTasas.Clear
cboTiposTasas.AddItem "TODAS"
cboTiposTasas.AddItem "Revisables"
cboTiposTasas.AddItem "Indizadas"
cboTiposTasas.Text = "TODAS"

cboAutorizaciones.Clear
cboAutorizaciones.AddItem "Autorizadas"
cboAutorizaciones.AddItem "Normales"
cboAutorizaciones.AddItem "TODAS"
cboAutorizaciones.Text = "TODAS"

cboSigno(0).Clear
cboSigno(0).AddItem ">"
cboSigno(0).AddItem "<"
cboSigno(0).AddItem "="
cboSigno(0).Text = "="

cboSigno(1).Clear
cboSigno(1).AddItem ">"
cboSigno(1).AddItem "<"
cboSigno(1).AddItem "="
cboSigno(1).Text = "="

txtPrideduc.Text = GLOBALES.glngFechaCR
txtUltMov.Text = GLOBALES.glngFechaCR

txtPlazoDesde.Text = 1
txtPlazoHasta.Text = 999

txtTasaDesde.Text = 0
txtTasaHasta.Text = 100

chkTasas_Click
chkPlazos_Click
chkPriDeduc_Click
chkUltMov_Click


cboSexo.Clear
cboSexo.AddItem "TODOS"
cboSexo.AddItem "Femenino"
cboSexo.AddItem "Masculino"
cboSexo.Text = "TODOS"


strSQL = "select Estado_Laboral as 'IdX', Descripcion as 'ItmX'" _
       & " from AFI_ESTADO_LABORAL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboCondicion, strSQL, True, True)

strSQL = "select Estado_Civil as 'IdX', Descripcion as 'ItmX' from SYS_ESTADO_CIVIL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoCivil, strSQL, True, True)


Call chkFechas_Click
Call chkLineas_Click
Call chkProvincias_Click
Call cboInstitucion_Click

Call sbRefrescaArbol

Me.MousePointer = vbDefault

End Sub


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer, rsTmp As New ADODB.Recordset

On Error GoTo vError

If Right(Node.Key, 1) = "Z" Then
  lblReporte.Caption = Node.Text
  lblReporte.Tag = fxIndiceCodigo(Node.Key)
  
  strSQL = "select Tipo,isnull(Adicional,0) as Adicional,isnull(Seguridad,0) as Seguridad" _
         & " from crd_reportes where id = " & lblReporte.Tag
  Call OpenRecordSet(rs, strSQL)
  
       cboESolicitud.Enabled = True
       cboEOperacion.Enabled = True
       
  If rs!seguridad = 1 Then
     imgSeguridad.Visible = True
     
    'Verificar que la persona tenga acceso a este reporte
    strSQL = "select isnull(COUNT(*),0) as Existe" _
           & " From CRD_REPORTES_GRP_AUT where id = " & lblReporte.Tag _
           & " and cod_grupo in(select cod_grupo from crd_reportes_grp_usr where usuario = '" & glogon.Usuario & "')"
    Call OpenRecordSet(rsTmp, strSQL, 0)
    If rsTmp!Existe = 0 Then
       lblSeguridad.Caption = "[ Requiere Grupo de Acceso Autorizado ]"
       lblSeguridad.ForeColor = vbRed
    Else
       lblSeguridad.Caption = "[ Usuario Tiene Acceso Autorizado ]"
       lblSeguridad.ForeColor = vbBlue
    End If
    rsTmp.Close
     
  Else
     imgSeguridad.Visible = False
  End If
  lblSeguridad.Visible = imgSeguridad.Visible
       
       
    tcFiltros.Item(0).Selected = True
    tcFiltros.Item(1).Enabled = False
 
  For i = 0 To 2
    fraAdicional.Item(i).Visible = False
  Next i
  
  Select Case rs!adicional
    Case 1 'Requisitos
      tcFiltros.Item(1).Enabled = True
      fraAdicional.Item(0).Visible = True
         
         
    Case 2 'Causas
      tcFiltros.Item(1).Enabled = True
      fraAdicional.Item(1).Visible = True
      
    Case 3 'Cortes
      tcFiltros.Item(1).Enabled = True
      fraAdicional.Item(2).Visible = True
      
    Case Else
  
  End Select
  
   
  'Pone el frame visible en el top adecuado.
  For i = 0 To 2
    If fraAdicional.Item(i).Visible Then
       fraAdicional.Item(i).top = 360
    End If
  Next i
  
  
  Select Case UCase(Trim(rs!Tipo))
     Case "CRD", "CBR", "RET"
       lblEstado.Caption = "Operación"
       cboEOperacion.Visible = True
       cboESolicitud.Visible = False
     
     Case "SGT"
       lblEstado.Caption = "Solicitud"
       cboESolicitud.Visible = True
       cboEOperacion.Visible = False
     
     Case "ESP"
       lblEstado.Caption = "N/A"
       cboESolicitud.Enabled = False
       cboEOperacion.Enabled = False
     
  End Select
  rs.Close
End If

vError:

End Sub


Private Sub btnReporte_Click()
Dim vTipo As String

On Error GoTo vError

strSQL = "select Tipo from crd_reportes where id = " & lblReporte.Tag
Call OpenRecordSet(rs, strSQL)
vTipo = UCase(Trim(rs!Tipo))
rs.Close


Select Case vTipo
 Case "CRD", "SGT"
     Call sbReporteCRD
  Case "RET"
     Call sbReporteRET

End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboCausasTipos_Click()

If vPaso Then Exit Sub

strSQL = "select rtrim(cod_causas) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from  operacion_causas where tipo = '" & Mid(cboCausasTipos.Text, 1, 1) _
       & "' order by descripcion"
       
Call sbCbo_Llena_New(cboCausas, strSQL, True, True)

End Sub

Private Sub cboGrpAccssM_Click()

If vPaso Then Exit Sub

If cboGrpAccssM.ListCount <= 0 Then Exit Sub

vPaso = True

With lswGrpAccssM
 .ListItems.Clear
  
 strSQL = "select U.nombre,U.descripcion,A.usuario" _
        & " from Usuarios U left join CRD_REPORTES_GRP_USR A on U.nombre = A.usuario" _
        & " and U.estado = 'A'  and A.cod_grupo = " & cboGrpAccssM.ItemData(cboGrpAccssM.ListIndex) _
        & " order by A.usuario desc,U.nombre asc"
 Call OpenRecordSet(rs, strSQL, 0)
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!Nombre)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Usuario) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
End With

vPaso = False

End Sub


Private Sub cboGrpAccssR_Click()

If vPaso Then Exit Sub

If cboGrpAccssR.ListCount <= 0 Then Exit Sub

vPaso = True

With lswGrpAccssR
 .ListItems.Clear
  
 strSQL = "select R.tipo,R.id,R.reporte,A.cod_grupo" _
        & " from CRD_REPORTES R left join CRD_REPORTES_GRP_AUT A on R.id = A.id" _
        & " and A.cod_grupo = " & cboGrpAccssR.ItemData(cboGrpAccssR.ListIndex) _
        & " order by A.cod_grupo desc,R.tipo asc, R.id asc"
 Call OpenRecordSet(rs, strSQL, 0)
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!Id)
      itmX.SubItems(1) = rs!Tipo
      itmX.SubItems(2) = rs!Reporte
      If Not IsNull(rs!Cod_Grupo) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
End With

vPaso = False

End Sub

Private Sub cboInstitucion_Click()

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

Private Sub cboMiembros_Click()

If vPaso Then Exit Sub
If cboMiembros.ListCount <= 0 Then Exit Sub

vPaso = True

With lswMiembros
 .ListItems.Clear
  
 strSQL = "select U.nombre,U.descripcion,A.usuario" _
        & " from Usuarios U left join crd_grpusers A on U.nombre = A.usuario" _
        & " and U.estado = 'A'  and A.cod_grupo = '" & cboMiembros.ItemData(cboMiembros.ListIndex) & "'" _
        & " order by A.usuario desc,U.nombre asc"
 Call OpenRecordSet(rs, strSQL, 0)
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!Nombre)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Usuario) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
End With

vPaso = False

End Sub

Private Sub cboCanton_Click()

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

Private Sub cboProvincia_Click()

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click

End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub


Private Sub chkCantones_Click()
If chkCantones.Value = vbChecked Then
   cboCanton.Enabled = False
Else
   cboCanton.Enabled = True
End If

chkDistritos.Value = chkCantones.Value
chkDistritos_Click
End Sub

Private Sub chkDistritos_Click()

If chkDistritos.Value = vbChecked Then
   cboDistrito.Enabled = False
Else
   cboDistrito.Enabled = True
End If

End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled
cboFBase.Enabled = dtpInicio.Enabled

End Sub


Private Sub chkLineas_Click()

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

Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xkey As String = "N")
Dim nodX As Node, vKey As String

On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub


Private Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim vPadre As String


With ArbolExp
  .Nodes.Clear
  Set vNode = .Nodes.Add(, , "Reportes", "Reportes", "imgRoot")
  Call sbCreaNodos("Reportes", "Créditos", "imgCRD", False, "0x0CRD")
  Call sbCreaNodos("Reportes", "Trámites", "imgSGT", False, "0x0SGT")
  Call sbCreaNodos("Reportes", "Cobro", "imgCBR", False, "0x0CBR")
  Call sbCreaNodos("Reportes", "Retenciones", "imgRetenciones", False, "0x0RET")
  Call sbCreaNodos("Reportes", "Especiales", "imgEspecial", False, "0x0ESP")
  
  strSQL = "select Id,Reporte,Tipo,isnull(seguridad,0) as Seguridad from crd_reportes order by reporte"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
    vPadre = "0x0" & Trim(UCase(rs!Tipo))
    If rs!seguridad = 0 Then
        Call sbCreaNodos(vPadre, rs!Reporte, "imgDetalle", False, "0x0" & rs!Id & "Z")
    Else
        Call sbCreaNodos(vPadre, rs!Reporte, "imgSeguridad", False, "0x0" & rs!Id & "Z")
    End If
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With

End Sub


Private Sub chkPlazos_Click()
If chkPlazos.Value = vbChecked Then
 txtPlazoDesde.Enabled = False
Else
 txtPlazoDesde.Enabled = True
End If
txtPlazoHasta.Enabled = txtPlazoDesde.Enabled
End Sub

Private Sub chkPriDeduc_Click()
If chkPriDeduc.Value = vbChecked Then
   txtPrideduc.Enabled = False
Else
   txtPrideduc.Enabled = True
End If
End Sub

Private Sub chkProvincias_Click()
If chkProvincias.Value = vbChecked Then
   cboProvincia.Enabled = False
Else
   cboProvincia.Enabled = True
End If

chkCantones.Value = chkProvincias.Value
chkCantones_Click

End Sub

Private Sub chkTasas_Click()
If chkTasas.Value = vbChecked Then
 txtTasaDesde.Enabled = False
Else
 txtTasaDesde.Enabled = True
End If
txtTasaHasta.Enabled = txtTasaDesde.Enabled

End Sub

Private Sub chkUltMov_Click()
If chkUltMov.Value = vbChecked Then
   txtUltMov.Enabled = False
Else
   txtUltMov.Enabled = True
End If
End Sub


Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vGrid.AppearanceStyle = fxGridStyle
vGridGrpAccss.AppearanceStyle = vGrid.AppearanceStyle
vGridRep.AppearanceStyle = vGrid.AppearanceStyle


With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2800
    .Add , , "Nombre", lswMiembros.Width - 2900
End With


With lswGrpAccssM.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2800
    .Add , , "Descripción", lswGrpAccssM.Width - 2900
End With

With lswGrpAccssR.ColumnHeaders
    .Clear
    .Add , , "Id", 900
    .Add , , "Tipo", 1000, vbCenter
    .Add , , "Reporte", lswGrpAccssR.Width - 2100
End With



Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub imgAddRep_Click()
'Inicializa Base de Datos con Reportes
glogon.Conection.Execute "exec spCRDReportesGen"
MsgBox "Lista de Reportes Actualizada...", vbInformation

End Sub



Private Sub lswGrpAccssM_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
  strSQL = "insert CRD_REPORTES_GRP_USR(cod_grupo,usuario) values(" & cboGrpAccssM.ItemData(cboGrpAccssM.ListIndex) _
         & ",'" & Item.Text & "')"
Else
  strSQL = "delete CRD_REPORTES_GRP_USR where cod_grupo = " & cboGrpAccssM.ItemData(cboGrpAccssM.ListIndex) _
         & " and usuario = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswGrpAccssR_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
  strSQL = "insert CRD_REPORTES_GRP_AUT(cod_grupo,id) values(" & cboGrpAccssR.ItemData(cboGrpAccssR.ListIndex) _
         & "," & Item.Text & ")"
Else
  strSQL = "delete CRD_REPORTES_GRP_AUT where cod_grupo = " & cboGrpAccssR.ItemData(cboGrpAccssR.ListIndex) _
         & " and id = " & Item.Text
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError


If Item.Checked Then
  'Preguntar si ya Existe el Usuario en Otro Grupo. / de ser asi no continuar
  strSQL = "select isnull(count(*),0) as Existe from crd_grpUsers where cod_grupo <> '" _
         & cboMiembros.ItemData(cboMiembros.ListIndex) & "' and usuario = '" & Item.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe > 0 Then
     rs.Close
     Item.Checked = False
     MsgBox "El Usuario ya ha sido asignado a otro grupo, proceda a excluirlo primero del otro grupo antes de agregarlo a este", vbExclamation
     Exit Sub
  End If
  rs.Close
End If


If Item.Checked Then
  strSQL = "insert crd_grpusers(cod_grupo,usuario) values('" & cboMiembros.ItemData(cboMiembros.ListIndex) _
         & "','" & Item.Text & "')"
Else
  strSQL = "delete crd_grpusers where cod_grupo = '" & cboMiembros.ItemData(cboMiembros.ListIndex) _
         & "' and usuario = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Function fxReporteFile() As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select prefijo from crd_reportes where id = " & lblReporte.Tag
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  fxReporteFile = ""
Else
  fxReporteFile = Trim(rs!prefijo)
End If
rs.Close

End Function



Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Grupos
    strSQL = "select cod_grupo,descripcion from crd_grupos order by cod_grupo"
    Call sbCargaGrid(vGrid, 2, strSQL)
  
  Case 1 'Miembros
    vPaso = True
    strSQL = "select cod_grupo as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  crd_grupos"
    Call sbCbo_Llena_New(cboMiembros, strSQL, False, True)
    vPaso = False
    
    Call cboMiembros_Click
    
  Case 2 'Reportes
    strSQL = "select ID,Tipo,Reporte,Prefijo,Adicional,isnull(Seguridad,0) from crd_reportes order by tipo,reporte"
    Call sbCargaGrid(vGridRep, 6, strSQL)

End Select

End Sub

Private Sub tcAuxGrpAccs_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  
  Case 0 'Grupos
    strSQL = "select cod_grupo,descripcion,activo from crd_reportes_grp order by cod_grupo"
    Call sbCargaGrid(vGridGrpAccss, 3, strSQL)
  
  Case 1 'Miembros
    vPaso = True
    strSQL = "select cod_grupo as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  crd_reportes_grp where activo = 1"
    Call sbCbo_Llena_New(cboGrpAccssM, strSQL, False, True)
    vPaso = False
    
    Call cboGrpAccssM_Click
  
  Case 2 'Reportes
    vPaso = True
    strSQL = "select cod_grupo as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  crd_reportes_grp where activo = 1"
    Call sbCbo_Llena_New(cboGrpAccssR, strSQL, False, True)
    vPaso = False
    
    Call cboGrpAccssR_Click

End Select

End Sub

Private Sub tcFiltros_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Me.MousePointer = vbHourglass

If Item.Index = 1 Then
 
 'Actualiza requisitos
 If fraAdicional.Item(0).Visible Then
    If chkLineas.Value = vbChecked Then
       strSQL = "select rtrim(R.cod_requisito) as 'IdX', rtrim(R.descripcion) as 'Itmx'" _
               & " from requisitos_adicionales R order by descripcion"
    Else
        strSQL = "select rtrim(R.cod_requisito) as 'IdX', rtrim(R.descripcion) as 'Itmx'" _
               & " from requisitos_adicionales R inner join requisitos_asignacion A" _
               & " on R.cod_requisito = A.cod_requisito and A.Codigo = '" _
               & txtCodigo & "'"
    End If
   Call sbCbo_Llena_New(cboRequisitos, strSQL, True, True)
 
 Else
   Call cboCausasTipos_Click
 End If

End If

Me.MousePointer = vbDefault

End Sub

Private Sub tcPrincipal_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Inicial
      Call sbInicializa
    
    Case 1 'Configuracion
      tcAux.Item(0).Selected = True
      
      strSQL = "select cod_grupo,descripcion from crd_grupos order by cod_grupo"
      Call sbCargaGrid(vGrid, 2, strSQL)
    
      lblReporte.Caption = "Configuración de Reportes"
    
    Case 2 'Seguridad
      tcAuxGrpAccs.Item(0).Selected = True
      
      strSQL = "select cod_grupo,descripcion,activo from crd_reportes_grp order by cod_grupo"
      Call sbCargaGrid(vGridGrpAccss, 3, strSQL)

      lblReporte.Caption = "Seguridad de Reportes"
End Select

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbInicializa
End Sub


Private Sub sbReporteCRD()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String, vTemp As String

On Error GoTo vError

If lblReporte.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass


If imgSeguridad.Visible Then
  'Verificar que la persona tenga acceso a este reporte
  strSQL = "select isnull(COUNT(*),0) as Existe" _
         & " From CRD_REPORTES_GRP_AUT where id = " & lblReporte.Tag _
         & " and cod_grupo in(select cod_grupo from crd_reportes_grp_usr where usuario = '" & glogon.Usuario & "')"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then
     Me.MousePointer = vbDefault
     rs.Close
     MsgBox "El usuario actual no tiene acceso autorizado a este reporte, verifique...", vbExclamation
     Exit Sub
  End If
  rs.Close
End If


 If fraAdicional.Item(2).Visible Then
    vTitulo = UCase(lblReporte.Caption & " [" & cboCorte.Text & "] " & cboTipo.Text)
 Else
    vTitulo = UCase(lblReporte.Caption & " : " & cboTipo.Text)
 End If



vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Créditos"
 
 .Connect = glogon.ConectRPT
  
 If chkFechas.Value = vbUnchecked Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    Select Case Mid(cboFBase.Text, 1, 1)
      Case "S" 'Recepción
        strSQL = strSQL & "{vCRDCreditosReportes01.fechaSol}"
        vSubTitulo = "Solicitadas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case "R" 'Resolución
        strSQL = strSQL & "{vCRDCreditosReportes01.fechares}"
        vSubTitulo = "Resueltas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case "F" 'Formalizacion
        strSQL = strSQL & "{vCRDCreditosReportes01.fechaforp}"
        vSubTitulo = "Formalizadas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case "D" 'Desembolso
        strSQL = strSQL & "{vCRDCreditosReportes01.fecha_inicio_Calculo}"
        vSubTitulo = "Desembolsos entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case "U" 'Ultimo Movimiento
        strSQL = strSQL & "{vCRDCreditosReportes01.Ultimo_Movimiento}"
        vSubTitulo = "Ultimo Movimiento entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
    End Select
    strSQL = strSQL & " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" _
           & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
 Else
   vSubTitulo = "Historico"
 End If
 
 
 If cboDivisa.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.COD_DIVISA} = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
 End If
 vSubTitulo = vSubTitulo & " ¦ Divisa: " & cboDivisa.Text


 If cboEspecial.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   
   Select Case cboEspecial.Text
        Case "Cartera Interna"
            strSQL = strSQL & "{vCRDCreditosReportes01.Linea_Interna} = 1"
        Case "Cartera Administrada"
            strSQL = strSQL & "{vCRDCreditosReportes01.Linea_Interna} = 0"
   End Select
 End If
 vSubTitulo = vSubTitulo & " ¦ Listado: " & cboEspecial.Text
 
 If cboESolicitud.Visible Then
    If Mid(cboESolicitud.Text, 1, 1) = "T" Then
      vSubTitulo = vSubTitulo & ", Estado Solicitud: " & cboESolicitud.Text
    Else
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{vCRDCreditosReportes01.estadosol} = '" & Mid(cboESolicitud.Text, 1, 1) & "'"
      vSubTitulo = vSubTitulo & ", Estado Solicitud: " & cboESolicitud.Text
    End If
 End If
 
 If cboEOperacion.Visible Then
    If Mid(cboEOperacion.Text, 1, 1) = "T" Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "({vCRDCreditosReportes01.estado} = 'A' OR {vCRDCreditosReportes01.estado} = 'C')"
      vSubTitulo = vSubTitulo & ", Estado Operación : " & cboEOperacion.Text
    Else
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{vCRDCreditosReportes01.estado} = '" & Mid(cboEOperacion.Text, 1, 1) & "'"
      vSubTitulo = vSubTitulo & ", Estado Operación: " & cboEOperacion.Text
    End If
 End If
    
    
If cboEPersona.Text <> "TODOS" Then
     Select Case cboEPersona.ItemData(cboEPersona.ListIndex)
         Case "X" 'Opex
           If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
           strSQL = strSQL & "({vCRDCreditosReportes01.estadoactual} = 'A' OR {vCRDCreditosReportes01.estadoactual} = 'P')"
         Case Else
           If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
           strSQL = strSQL & "{vCRDCreditosReportes01.estadoactual} = '" & cboEPersona.ItemData(cboEPersona.ListIndex) & "'"
    End Select
End If
 vFiltro = vFiltro & ", Condición: " & cboEPersona.Text
    
 If cboOficina.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.cod_oficina_R} = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
 End If
 vSubTitulo = vSubTitulo & ", Oficina: " & cboOficina.Text
    
    
 If cboGarantia.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.garantia} = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
 End If
 vSubTitulo = vSubTitulo & ", Garantía: " & cboGarantia.Text
 
 If cboUsuarios.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   If cboESolicitud.Enabled Then
       strSQL = strSQL & "{vCRDCreditosReportes01.GRUPOREC} = '" & cboUsuarios.ItemData(cboUsuarios.ListIndex) & "'"
   Else
       strSQL = strSQL & "{vCRDCreditosReportes01.COD_GRUPO} = '" & cboUsuarios.ItemData(cboUsuarios.ListIndex) & "'"
   End If
 End If
 vFiltro = ", Grupo: " & cboUsuarios
 
 If chkLineas.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.Codigo} = '" & Trim(txtCodigo.Text) & "'"
   vFiltro = vFiltro & ", Línea: " & UCase(txtCodigo.Text)
 Else
   vFiltro = vFiltro & ", Todas las Líneas"
 End If
 
 If cboRecurso.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.recurso} = '" & Trim(cboRecurso.ItemData(cboRecurso.ListIndex)) & "'"
 End If
 vFiltro = vFiltro & ", Recurso: " & cboRecurso.Text
 
 If cboDestino.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.cod_destino} = '" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
 End If
 vFiltro = vFiltro & ", Destino: " & cboDestino.Text
 
 If cboComite.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.id_comite} = " & cboComite.ItemData(cboComite.ListIndex) & ""
   vFiltro = vFiltro & ", Comité: " & cboComite.Text
 End If
 
 If cboInstitucion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.cod_institucion} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ""
   vFiltro = vFiltro & ", Institución: " & cboInstitucion.Text
 End If
  
 'Deductora
 If cboDeductora.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.cod_deductora} = " & cboDeductora.ItemData(cboDeductora.ListIndex) & ""
   vFiltro = vFiltro & ", Deductora: " & cboDeductora.Text
 End If
 
 
 
'Otros Filtros
 If cboProceso.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
     Select Case Mid(cboProceso.Text, 1, 1)
       Case "N"
            strSQL = strSQL & "{vCRDCreditosReportes01.Proceso} = 'N'"
       Case "T"
            strSQL = strSQL & "{vCRDCreditosReportes01.Proceso} = 'T'"
       Case "C"
            strSQL = strSQL & "{vCRDCreditosReportes01.Proceso} = 'J'"
     End Select
     vFiltro = vFiltro & ", Proceso: " & cboProceso.Text
 End If


 If cboTipoOperacion.Text <> "TODAS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   If Mid(cboTipoOperacion.Text, 1, 1) = "O" Then
       strSQL = strSQL & "ISNULL({vCRDCreditosReportes01.REFERENCIA}) = TRUE"
   Else
       strSQL = strSQL & "ISNULL({vCRDCreditosReportes01.REFERENCIA}) = FALSE"
   End If
   
   vFiltro = vFiltro & ", Operaciones: " & cboTipoOperacion.Text
 End If


 If cboCobro.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   If cboCobro.Text = "Cajas" Then
       strSQL = strSQL & "{vCRDCreditosReportes01.Ind_deduce_Planilla} = 'N'"
   Else
       strSQL = strSQL & "{vCRDCreditosReportes01.Ind_deduce_Planilla} = 'S'"
   End If
   
   vFiltro = vFiltro & ", Cobro vía: " & cboCobro.Text
 End If

 If cboTiposTasas.Text <> "TODAS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   If Mid(cboTiposTasas.Text, 1, 1) = "I" Then 'Indizadas
       strSQL = strSQL & "ISNULL({vCRDCreditosReportes01.TBP_PuntosAdd}) = TRUE"
   Else 'Revisables
       strSQL = strSQL & "ISNULL({vCRDCreditosReportes01.TBP_PuntosAdd}) = FALSE"
   End If
   
   vFiltro = vFiltro & ", Tipo Tasa: " & cboTiposTasas.Text
 End If
 
 
 If cboAutorizaciones.Text <> "TODAS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   If Mid(cboAutorizaciones.Text, 1, 1) = "N" Then 'Normales
       strSQL = strSQL & "ISNULL({vCRDCreditosReportes01.Autoriza_Fecha}) = TRUE"
   Else 'Autorizadas
       strSQL = strSQL & "ISNULL({vCRDCreditosReportes01.Autoriza_Fecha}) = FALSE"
   End If
   
   vFiltro = vFiltro & ", Autoriza: " & cboAutorizaciones.Text
 End If
 
 
 If chkPlazos.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.Plazo} >= " & txtPlazoDesde.Text & " AND {vCRDCreditosReportes01.Plazo} <=" & txtPlazoHasta.Text
    
    vFiltro = vFiltro & ",Plazos: " & txtPlazoDesde.Text & " - " & txtPlazoHasta.Text
 End If
 
 If chkTasas.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.Interesv} >= " & txtTasaDesde.Text & " AND {vCRDCreditosReportes01.Interesv} <=" & txtTasaHasta.Text
    
    vFiltro = vFiltro & ", Tasas : " & txtTasaDesde.Text & " - " & txtTasaHasta.Text
 End If
 
 'Primer Deducción
 If chkPriDeduc.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.PriDeduc} " & cboSigno(0).Text & txtPrideduc.Text
    
   vFiltro = vFiltro & ", Pri.Deduc. " & cboSigno(0).Text & " " & txtPrideduc.Text
 End If
 
 'Ultimo Movimiento
 If chkUltMov.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.FecUlt} " & cboSigno(1).Text & txtUltMov.Text
    
   vFiltro = vFiltro & ", Ult.Mov. " & cboSigno(1).Text & " " & txtUltMov.Text
 End If
 
 
 
 
 'Ejecutivo Colocador
 If IsNumeric(txtEjecutivoId.Text) Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.Id_Promotor} = " & txtEjecutivoId.Text
   vSubTitulo = vSubTitulo & ", Ejecutivo: " & txtEjecutivoName.Text
 End If
 
 'Parametros Adicionales
 If tcFiltros.Item(1).Enabled Then
 
   Select Case True
     Case fraAdicional.Item(0).Visible 'Requisitos
     
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        Select Case cboRequistoMarca.Text
          Case "Cumple"
             strSQL = strSQL & "{OPERACION_REQUISITOS.ESTADO} = 1"
          Case "No Cumple"
             strSQL = strSQL & "{OPERACION_REQUISITOS.ESTADO} = 2"
          Case "En Blanco"
             strSQL = strSQL & "{OPERACION_REQUISITOS.ESTADO} = 0"
        End Select
        
        If cboRequisitos.Text <> "TODOS" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{OPERACION_REQUISITOS.COD_REQUISITO} = '" & cboRequisitos.ItemData(cboRequisitos.ListIndex) & "'"
        End If
        
        vFiltro = vFiltro & ", Requisitos: " & cboRequisitos.Text & " ESTADO : " & cboRequistoMarca.Text
        
        vSubTitulo = vSubTitulo & ", Requisito Estado: " & cboRequistoMarca.Text
     
     
     Case fraAdicional.Item(1).Visible 'Causas
        
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{OPERACION_GESTION.TIPO} = '" & Mid(cboCausasTipos.Text, 1, 1) & "'"
        
        If cboCausas.Text <> "TODOS" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{OPERACION_GESTION.COD_CAUSAS} = '" & cboCausas.ItemData(cboCausas.ListIndex) & "'"
        End If
        
        vFiltro = vFiltro & ", Causa de Rechazo: " & cboCausas.Text & " TIPO : " & cboCausasTipos.Text
        
        vSubTitulo = vSubTitulo & ", Tipo de Causa: " & cboCausasTipos.Text
     
   End Select
 End If
 
 'Nuevos Filtros
 
  If Mid(cboSexo.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.sexo} = '" & Mid(cboSexo.Text, 1, 1) & "'"
 End If
 vSubTitulo = vSubTitulo & ", Sexo: " & cboSexo.Text
  
  If cboEstadoCivil.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.EstadoCivil} = '" & cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex) & "'"
 End If
 vSubTitulo = vSubTitulo & ", Estado Civil: " & cboEstadoCivil.Text
 
 If cboCondicion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
     strSQL = strSQL & "{vCRDCreditosReportes01.EstadoLaboral} = '" & cboCondicion.ItemData(cboCondicion.ListIndex) & "'"
 End If
 vFiltro = vFiltro & ", Laboral: " & cboCondicion.Text

  If cboSector.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.cod_sector} = " & cboSector.ItemData(cboSector.ListIndex) & ""
 End If
 vFiltro = vFiltro & ", Sector: " & cboSector.Text
 
 
 If cboProfesion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.cod_profesion} = " & cboProfesion.ItemData(cboProfesion.ListIndex) & ""
 End If
 vFiltro = vFiltro & ", Profesión: " & cboProfesion.Text
 
 
 If cboZonas.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.cod_zona} = '" & cboZonas.ItemData(cboZonas.ListIndex) & "'"
 End If
 vFiltro = vFiltro & ", Zona: " & cboZonas.Text
 
 
 'Filtros Adicionales
If chkProvincias.Value = vbUnchecked And cboProvincia.ListCount > 0 Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.provincia} = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "'"
   vFiltro = vFiltro & ", Provincia: " & cboProvincia.Text
Else
   vFiltro = vFiltro & ", Provincia: Todas"
End If
 
If chkCantones.Value = vbUnchecked And cboCanton.ListCount > 0 Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.canton} = '" & cboCanton.ItemData(cboCanton.ListIndex) & "'"
   vFiltro = vFiltro & " [" & cboCanton.Text & "]"
End If
 
If chkDistritos.Value = vbUnchecked And cboDistrito.ListCount > 0 Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.distrito} = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "'"
   vFiltro = vFiltro & " [" & cboDistrito.Text & "]"
End If
 
If chkDepartamento.Value = vbUnchecked And txtDeptCodigo.Text <> "" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.DeptCod} = '" & txtDeptCodigo.Text & "'"
   vFiltro = vFiltro & ", Dept.: " & txtDeptCodigo.Text
    
   If chkSeccion.Value = vbUnchecked And txtSecCodigo.Text <> "" Then
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "{vCRDCreditosReportes01.SecCod} = '" & txtSecCodigo.Text & "'"
       vFiltro = vFiltro & " [" & txtSecCodigo.Text & "]"
   End If
End If
 
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='" & vTitulo & "'"
 .Formulas(4) = "fxSubTitulo='" & Mid(vSubTitulo, 1, 250) & "'"
 .Formulas(5) = "fxFiltro='" & Mid(vFiltro, 1, 250) & "'"
 
 If fraAdicional.Item(2).Visible Then
    .ReportFileName = SIFGlobal.fxPathReportes(Trim(fxReporteFile) & "_" & cboCorte.Text & "_" & Trim(cboTipo.Text) & ".rpt")
 Else
    .ReportFileName = SIFGlobal.fxPathReportes(Trim(fxReporteFile) & "_" & Trim(cboTipo.Text) & ".rpt")
 End If
 .SelectionFormula = strSQL

 .PrintReport
End With

Me.MousePointer = vbDefault

Call Bitacora("Imprime", lblReporte.Caption)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbReporteRET()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String, vTemp As String

On Error GoTo vError

If lblReporte.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass


If imgSeguridad.Visible Then
  'Verificar que la persona tenga acceso a este reporte
  strSQL = "select isnull(COUNT(*),0) as Existe" _
         & " From CRD_REPORTES_GRP_AUT where id = " & lblReporte.Tag _
         & " and cod_grupo in(select cod_grupo from crd_reportes_grp_usr where usuario = '" & glogon.Usuario & "')"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then
     Me.MousePointer = vbDefault
     rs.Close
     MsgBox "El usuario actual no tiene acceso autorizado a este reporte, verifique...", vbExclamation
     Exit Sub
  End If
  rs.Close
End If


 If fraAdicional.Item(2).Visible Then
    vTitulo = UCase(lblReporte.Caption & " [" & cboCorte.Text & "] " & cboTipo.Text)
 Else
    vTitulo = UCase(lblReporte.Caption & " : " & cboTipo.Text)
 End If



vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Créditos"
 
 .Connect = glogon.ConectRPT
  
 If chkFechas.Value = vbUnchecked Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    Select Case Mid(cboFBase.Text, 1, 1)
      Case "S"
        strSQL = strSQL & "{vCRDRetencionesReportes01.fechaSol}"
        vSubTitulo = "Solicitadas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case "R"
        strSQL = strSQL & "{vCRDRetencionesReportes01.fechares}"
        vSubTitulo = "Resueltas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case "F"
        strSQL = strSQL & "{vCRDRetencionesReportes01.fechaforp}"
        vSubTitulo = "Formalizadas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case "D"
        strSQL = strSQL & "{vCRDRetencionesReportes01.fecha_inicio_Calculo}"
        vSubTitulo = "Desembolsos entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case "U" 'Ultimo Movimiento
        strSQL = strSQL & "{vCRDRetencionesReportes01.Ultimo_Movimiento}"
        vSubTitulo = "Ultimo Movimiento entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
    End Select
    strSQL = strSQL & " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" _
           & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
 Else
   vSubTitulo = "Historico"
 End If
 
' If cboESolicitud.Visible Then
'    If Mid(cboESolicitud.Text, 1, 1) = "T" Then
'      vSubTitulo = vSubTitulo & " / Estado Solicitud : " & cboESolicitud.Text
'    Else
'      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'      strSQL = strSQL & "{vCRDRetencionesReportes01.estadosol} = '" & Mid(cboESolicitud.Text, 1, 1) & "'"
'      vSubTitulo = vSubTitulo & " / Estado Solicitud : " & cboESolicitud.Text
'    End If
' End If
 
 If cboEOperacion.Visible Then
    If Mid(cboEOperacion.Text, 1, 1) = "T" Then
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "({vCRDRetencionesReportes01.estado} = 'A' OR {vCRDRetencionesReportes01.estado} = 'C')"
      vSubTitulo = vSubTitulo & ", Estado Operación: " & cboEOperacion.Text
    Else
      If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
      strSQL = strSQL & "{vCRDRetencionesReportes01.estado} = '" & Mid(cboEOperacion.Text, 1, 1) & "'"
      vSubTitulo = vSubTitulo & ", Estado Operación: " & cboEOperacion.Text
    End If
 End If
    
    
If cboEPersona.Text <> "TODOS" Then
     Select Case Mid(cboEPersona.Text, 1, 2)
         Case "X" 'Opex
           If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
           strSQL = strSQL & "({vCRDRetencionesReportes01.estadoactual} = 'A' OR {vCRDRetencionesReportes01.estadoactual} = 'P')"
         Case Else
           If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
           strSQL = strSQL & "{vCRDRetencionesReportes01.estadoactual} = '" & cboEPersona.ItemData(cboEPersona.ListIndex) & "'"
    End Select
End If
 vFiltro = vFiltro & ", Condición: " & cboEPersona.Text
    
 If cboOficina.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.cod_oficina_R} = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
 End If
 vSubTitulo = vSubTitulo & ", Oficina: " & cboOficina.Text
    
    
 If cboGarantia.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.garantia} = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
 End If
 vSubTitulo = vSubTitulo & ", Garantía: " & cboGarantia.Text
 
 If cboUsuarios.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   If cboESolicitud.Enabled Then
       strSQL = strSQL & "{vCRDRetencionesReportes01.GRUPOREC} = '" & cboUsuarios.ItemData(cboUsuarios.ListIndex) & "'"
   Else
       strSQL = strSQL & "{vCRDRetencionesReportes01.COD_GRUPO} = '" & cboUsuarios.ItemData(cboUsuarios.ListIndex) & "'"
   End If
 End If
 vFiltro = ", Grupo: " & cboUsuarios
 
 If chkLineas.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.Codigo} = '" & Trim(txtCodigo) & "'"
   vFiltro = vFiltro & ", Línea: " & UCase(txtCodigo)
 Else
   vFiltro = vFiltro & ", Todas las Líneas"
 End If
 
 
 If cboDestino.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.cod_destino} = '" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
 End If
 vFiltro = vFiltro & ", Destino: " & cboDestino.Text
 
 If cboComite.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.id_comite} = " & cboComite.ItemData(cboComite.ListIndex) & ""
   vFiltro = vFiltro & ", Comité: " & cboComite.Text
 End If
 
 If cboInstitucion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.cod_institucion} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ""
   vFiltro = vFiltro & ", Institución: " & cboInstitucion.Text
 End If
 
 'Deductora
 If cboDeductora.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.cod_deductora} = " & cboDeductora.ItemData(cboDeductora.ListIndex) & ""
   vFiltro = vFiltro & ", Deductora: " & cboDeductora.Text
 End If
 

 If cboCobro.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   If cboCobro.Text = "Cajas" Then
       strSQL = strSQL & "{vCRDRetencionesReportes01.Ind_deduce_Planilla} = 'N'"
   Else
       strSQL = strSQL & "{vCRDRetencionesReportes01.Ind_deduce_Planilla} = 'S'"
   End If
   
   vFiltro = vFiltro & ", Cobro vía: " & cboCobro.Text
 End If



 If chkPlazos.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.Plazo} >= " & txtPlazoDesde.Text & " AND {vCRDRetencionesReportes01.Plazo} <=" & txtPlazoHasta.Text
    
    vFiltro = vFiltro & ", Plazos: " & txtPlazoDesde.Text & " - " & txtPlazoHasta.Text
 End If
 
 If chkTasas.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.Interesv} >= " & txtTasaDesde.Text & " AND {vCRDRetencionesReportes01.Interesv} <=" & txtTasaHasta.Text
    
    vFiltro = vFiltro & ", Tasas: " & txtTasaDesde.Text & " - " & txtTasaHasta.Text
 End If
 
 'Primer Deducción
 If chkPriDeduc.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.PriDeduc} " & cboSigno(0).Text & txtPrideduc.Text
    
   vFiltro = vFiltro & ", Pri.Deduc. " & cboSigno(0).Text & " " & txtPrideduc.Text
 End If
 
 'Ultimo Movimiento
 If chkUltMov.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.FecUlt} " & cboSigno(1).Text & txtUltMov.Text
    
   vFiltro = vFiltro & ", Ult.Mov. " & cboSigno(1).Text & " " & txtUltMov.Text
 End If
 
 
 'Nuevos Filtros
 
  If Mid(cboSexo.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.sexo} = '" & Mid(cboSexo.Text, 1, 1) & "'"
 End If
 vSubTitulo = vSubTitulo & ", Sexo: " & cboSexo.Text
  
  If cboEstadoCivil.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.EstadoCivil} = '" & cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex) & "'"
 End If
 vSubTitulo = vSubTitulo & ", Estado Civil: " & cboEstadoCivil.Text
 
 If cboCondicion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
     strSQL = strSQL & "{vCRDRetencionesReportes01.EstadoLaboral} = '" & cboCondicion.ItemData(cboCondicion.ListIndex) & "'"
 End If
 vFiltro = vFiltro & ", Laboral: " & cboCondicion.Text

  If cboSector.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.cod_sector} = " & cboSector.ItemData(cboSector.ListIndex) & ""
 End If
 vFiltro = vFiltro & ", Sector: " & cboSector.Text
 
 
 If cboProfesion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.cod_profesion} = " & cboProfesion.ItemData(cboProfesion.ListIndex) & ""
 End If
 vFiltro = vFiltro & ", Profesión: " & cboProfesion.Text
 
 
 If cboZonas.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.cod_zona} = '" & cboZonas.ItemData(cboZonas.ListIndex) & "'"
 End If
 vFiltro = vFiltro & ", Zona: " & cboZonas.Text
 
 
 'Filtros Adicionales
If chkProvincias.Value = vbUnchecked And cboProvincia.ListCount > 0 Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.provincia} = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "'"
   vFiltro = vFiltro & ", Provincia: " & cboProvincia.Text
Else
   vFiltro = vFiltro & ", Provincia: Todas"
End If
 
If chkCantones.Value = vbUnchecked And cboCanton.ListCount > 0 Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.canton} = '" & cboCanton.ItemData(cboCanton.ListIndex) & "'"
   vFiltro = vFiltro & " [" & cboCanton.Text & "]"
End If
 
If chkDistritos.Value = vbUnchecked And cboDistrito.ListCount > 0 Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.distrito} = '" & cboDistrito.ItemData(cboDistrito.ListIndex) & "'"
   vFiltro = vFiltro & " [" & cboDistrito.Text & "]"
End If
 
If chkDepartamento.Value = vbUnchecked And txtDeptCodigo.Text <> "" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDRetencionesReportes01.DeptCod} = '" & txtDeptCodigo.Text & "'"
   vFiltro = vFiltro & ", Dept.: " & txtDeptCodigo.Text
    
   If chkSeccion.Value = vbUnchecked And txtSecCodigo.Text <> "" Then
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "{vCRDRetencionesReportes01.SecCod} = '" & txtSecCodigo.Text & "'"
       vFiltro = vFiltro & " [" & txtSecCodigo.Text & "]"
   End If
End If
 
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='" & vTitulo & "'"
 .Formulas(4) = "fxSubTitulo='" & Mid(vSubTitulo, 1, 250) & "'"
 .Formulas(5) = "fxFiltro='" & Mid(vFiltro, 1, 250) & "'"
 
 If fraAdicional.Item(2).Visible Then
    .ReportFileName = SIFGlobal.fxPathReportes(Trim(fxReporteFile) & "_" & cboCorte.Text & "_" & Trim(cboTipo.Text) & ".rpt")
 Else
    .ReportFileName = SIFGlobal.fxPathReportes(Trim(fxReporteFile) & "_" & Trim(cboTipo.Text) & ".rpt")
 End If
 .SelectionFormula = strSQL

 .PrintReport

End With

Me.MousePointer = vbDefault

Call Bitacora("Imprime", lblReporte.Caption)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from crd_Grupos" _
       & " where cod_grupo = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into crd_Grupos(cod_grupo,descripcion) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "')"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Grupo de Usuarios: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update crd_Grupos set descripcion = '" & vGrid.Text & "'"
 strSQL = strSQL & " where cod_grupo = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Grupo de Usuarios : " & vGrid.Text)


End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub txtDeptCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptDesc.SetFocus
If KeyCode = vbKeyF4 Then
  
    If mModoSif Then
      gBusquedas.Columna = "cod_departamento"
      gBusquedas.Orden = "cod_departamento"
      gBusquedas.Consulta = "select cod_departamento as codigo,descripcion from afDepartamentos"
      gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    Else
      gBusquedas.Columna = "codigo"
      gBusquedas.Orden = "codigo"
      gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
      gBusquedas.Filtro = ""
    End If

  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDeptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecCodigo.SetFocus
If KeyCode = vbKeyF4 Then

    If mModoSif Then
      gBusquedas.Columna = "descripcion"
      gBusquedas.Orden = "descripcion"
      gBusquedas.Consulta = "select cod_departamento as codigo,descripcion from afDepartamentos"
      gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    Else
      gBusquedas.Columna = "codigo"
      gBusquedas.Orden = "codigo"
      gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
      gBusquedas.Filtro = ""
    End If
  
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If
End Sub




Private Sub txtEjecutivoId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Ejectuvo Id"
  gBusquedas.Col2Name = "Nombre"
  gBusquedas.Col3Name = "Identificación"
  gBusquedas.Columna = "id_promotor"
  gBusquedas.Orden = "id_promotor"
  gBusquedas.Consulta = "select id_promotor,nombre, cod_Comision from promotores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtEjecutivoId.Text = gBusquedas.Resultado
  txtEjecutivoName.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtEjecutivoName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Ejectuvo Id"
  gBusquedas.Col2Name = "Nombre"
  gBusquedas.Col3Name = "Identificación"
  gBusquedas.Columna = "id_promotor"
  gBusquedas.Orden = "id_promotor"
  gBusquedas.Consulta = "select id_promotor,nombre, cod_Comision from promotores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtEjecutivoId.Text = gBusquedas.Resultado
  txtEjecutivoName.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtSecCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecDesc.SetFocus
If KeyCode = vbKeyF4 Then
  
    If mModoSif Then
        gBusquedas.Columna = "cod_seccion"
        gBusquedas.Orden = "cod_seccion"
        gBusquedas.Consulta = "select cod_seccion as codigo,descripcion from afSecciones"
        gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
                    & " and cod_departamento = '" & txtDeptCodigo.Text & "'"
    Else
        gBusquedas.Columna = "ut_codigo"
        gBusquedas.Orden = "ut_codigo"
        gBusquedas.Consulta = "select ut_codigo,ut_descripcion from UTRABAJO"
        gBusquedas.Filtro = ""
    End If
  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtSecDesc_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then optPropiedad.Item(0).SetFocus
If KeyCode = vbKeyF4 Then
    If mModoSif Then
        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Consulta = "select cod_seccion as codigo,descripcion from afSecciones"
        gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
                    & " and cod_departamento = '" & txtDeptCodigo.Text & "'"
    Else
        gBusquedas.Columna = "ut_descripcion"
        gBusquedas.Orden = "ut_descripcion"
        gBusquedas.Consulta = "select ut_codigo,ut_descripcion from UTRABAJO"
        gBusquedas.Filtro = ""
    End If
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If

End Sub



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

End Sub

Private Function fxGuardarRep() As Long

On Error GoTo vError

fxGuardarRep = 0
vGridRep.Row = vGridRep.ActiveRow
vGridRep.col = 1

If vGridRep.Text = "" Then 'Insertar
  vGridRep.col = 2
  strSQL = "insert into crd_reportes(tipo,reporte,prefijo,adicional,seguridad) values('" _
         & UCase(vGridRep.Text) & "','"
  vGridRep.col = 3
  strSQL = strSQL & vGridRep.Text & "','"
  vGridRep.col = 4
  strSQL = strSQL & vGridRep.Text & "',"
  vGridRep.col = 5
  strSQL = strSQL & vGridRep.Text & ","
  vGridRep.col = 6
  strSQL = strSQL & vGridRep.Value & ")"
  
  Call ConectionExecute(strSQL)

  vGridRep.col = 1
  
  strSQL = "select isnull(max(id),0) as Ultimo from crd_reportes"
  Call OpenRecordSet(rs, strSQL)
   vGridRep.Text = CStr(rs!ultimo)
  rs.Close
  
  Call Bitacora("Registra", "Reportes de Credito: " & vGridRep.Text)

Else 'Actualizar

 vGridRep.col = 2
 strSQL = "update crd_reportes set tipo = '" & vGridRep.Text & "',reporte = '"
 vGridRep.col = 3
 strSQL = strSQL & vGridRep.Text & "',prefijo = '"
 vGridRep.col = 4
 strSQL = strSQL & vGridRep.Text & "',adicional = "
 vGridRep.col = 5
 strSQL = strSQL & vGridRep.Text & ",seguridad = "
 vGridRep.col = 6
 strSQL = strSQL & vGridRep.Value & " Where ID = "
 vGridRep.col = 1
 strSQL = strSQL & vGridRep.Text
 
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Reportes de Credito : " & vGridRep.Text)


End If

fxGuardarRep = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Function



Private Function fxGuardarGrpAccss() As Long

On Error GoTo vError

fxGuardarGrpAccss = 0
vGridGrpAccss.Row = vGridGrpAccss.ActiveRow
vGridGrpAccss.col = 1

If vGridGrpAccss.Text = "" Then 'Insertar
  vGridGrpAccss.col = 2
  strSQL = "insert into crd_reportes_grp(descripcion,activo) values('" _
         & Trim(vGridGrpAccss.Text) & "',"
  vGridGrpAccss.col = 3
  strSQL = strSQL & vGridGrpAccss.Value & ")"
  
  Call ConectionExecute(strSQL)

  vGridGrpAccss.col = 1
  
  strSQL = "select isnull(max(cod_grupo),0) as Ultimo from crd_reportes_grp"
  Call OpenRecordSet(rs, strSQL)
   vGridGrpAccss.Text = CStr(rs!ultimo)
  rs.Close
  
  Call Bitacora("Registra", "Reportes > Grupo de Acceso: " & vGridGrpAccss.Text)

Else 'Actualizar

 vGridGrpAccss.col = 2
 strSQL = "update crd_reportes_grp set descripcion = '" & Trim(vGridGrpAccss.Text) & "',activo = "
 vGridGrpAccss.col = 3
 strSQL = strSQL & vGridGrpAccss.Value & " where cod_grupo = "
 vGridGrpAccss.col = 1
 strSQL = strSQL & vGridGrpAccss.Text
 
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Reportes > Grupo de Acceso: " & vGridGrpAccss.Text)


End If

fxGuardarGrpAccss = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Function


Private Sub vGridGrpAccss_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGridGrpAccss.ActiveCol = vGridGrpAccss.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarGrpAccss
  If i = 0 Then Exit Sub
  vGridGrpAccss.Row = vGridGrpAccss.ActiveRow
  If vGridGrpAccss.MaxRows <= vGridGrpAccss.ActiveRow Then
    vGridGrpAccss.MaxRows = vGridGrpAccss.MaxRows + 1
    vGridGrpAccss.Row = vGridGrpAccss.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridGrpAccss.MaxRows = vGridGrpAccss.MaxRows + 1
    vGridGrpAccss.InsertRows vGridGrpAccss.ActiveRow, 1
    vGridGrpAccss.Row = vGridGrpAccss.ActiveRow
End If
End Sub



Private Sub vGridRep_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGridRep.ActiveCol = vGridRep.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  If txtReportes = "trick" Then
        i = fxGuardarRep
        If i = 0 Then Exit Sub
        vGridRep.Row = vGridRep.ActiveRow
        If vGridRep.MaxRows <= vGridRep.ActiveRow Then
          vGridRep.MaxRows = vGridRep.MaxRows + 1
          vGridRep.Row = vGridRep.MaxRows
        End If
  Else
    MsgBox "Proporcione la contraseña de Administrador", vbInformation
  End If

End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridRep.MaxRows = vGridRep.MaxRows + 1
    vGridRep.InsertRows vGridRep.ActiveRow, 1
    vGridRep.Row = vGridRep.ActiveRow
End If

End Sub
