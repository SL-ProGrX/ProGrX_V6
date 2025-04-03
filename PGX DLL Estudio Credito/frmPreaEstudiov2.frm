VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmPreaEstudiov2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Estudio de Credito (Preanálisis v2)"
   ClientHeight    =   11460
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18180
   LinkTopic       =   "Form1"
   ScaleHeight     =   897.504
   ScaleMode       =   0  'User
   ScaleWidth      =   1212
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.Resizer Resizer1 
      Height          =   11460
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18180
      _Version        =   1572864
      _ExtentX        =   32067
      _ExtentY        =   20214
      _StockProps     =   1
      VScrollLargeChange=   1000
      VScrollSmallChange=   200
      HScrollLargeChange=   1000
      HScrollSmallChange=   200
      BorderStyle     =   2
      AutoSize        =   -1  'True
      ShowSizeIcon    =   -1  'True
      ControlCount    =   8
      Control(0).Caption=   "tcMain"
      Control(0).Width=   100
      Control(0).Height=   50
      Control(1).Caption=   "stBar"
      Control(1).Y    =   100
      Control(1).Width=   100
      Control(2).Caption=   "gbComite"
      Control(3).Caption=   "gbCredito"
      Control(3).Width=   100
      Control(4).Caption=   "gbSalarios(0)"
      Control(4).Width=   33
      Control(5).Caption=   "gbSalarios(1)"
      Control(5).X    =   34
      Control(5).Width=   33
      Control(6).Caption=   "gbSalarios(2)"
      Control(6).X    =   64
      Control(6).Width=   33
      Control(6).Height=   80
      Control(7).Caption=   "scTitulo"
      Control(7).Width=   100
      Begin XtremeSuiteControls.UpDown UpDownExpediente 
         Height          =   405
         Left            =   3840
         TabIndex        =   302
         Top             =   600
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1085
         _ExtentY        =   714
         _StockProps     =   64
         Orientation     =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
         BuddyControl    =   ""
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.GroupBox gbTitulo 
         Height          =   735
         Left            =   10560
         TabIndex        =   73
         Top             =   480
         Width           =   7455
         _Version        =   1572864
         _ExtentX        =   13150
         _ExtentY        =   1296
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtSalarioMinimoInembargable 
            Height          =   315
            Left            =   1800
            TabIndex        =   76
            Top             =   120
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSalarioNormativa 
            Height          =   315
            Left            =   5520
            TabIndex        =   77
            Top             =   120
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   375
            Index           =   4
            Left            =   3960
            TabIndex        =   75
            Top             =   120
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Salario Normativa"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   74
            Top             =   120
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Salario Mínimo Inembargable"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox gbComite 
         Height          =   1455
         Left            =   10560
         TabIndex        =   30
         Top             =   1200
         Width           =   7455
         _Version        =   1572864
         _ExtentX        =   13150
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Asigna Comité Resolutivo"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboComite 
            Height          =   330
            Left            =   0
            TabIndex        =   31
            Top             =   360
            Width           =   7335
            _Version        =   1572864
            _ExtentX        =   12938
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
         Begin XtremeSuiteControls.PushButton btnComiteCambio 
            Height          =   330
            Left            =   6120
            TabIndex        =   59
            ToolTipText     =   "Cambio de Evaluador"
            Top             =   720
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Asigna"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Picture         =   "frmPreaEstudiov2.frx":0000
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.FlatEdit txtDiasIntereses 
            Height          =   315
            Left            =   1440
            TabIndex        =   270
            Top             =   720
            Visible         =   0   'False
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1714
            _ExtentY        =   550
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
            Text            =   "0"
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnNotificacion 
            Height          =   330
            Left            =   4800
            TabIndex        =   321
            ToolTipText     =   "Enviar Notificación de Resolución a la Persona"
            Top             =   720
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Notificar"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmPreaEstudiov2.frx":0727
         End
         Begin VB.Label lblDiasInt 
            BackStyle       =   0  'Transparent
            Caption         =   "Días de Intereses"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   271
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin XtremeSuiteControls.GroupBox gbDesecho 
         Height          =   1095
         Left            =   17760
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   6015
         _Version        =   1572864
         _ExtentX        =   10610
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   "Desecho"
         UseVisualStyle  =   -1  'True
         Begin VB.ComboBox cboSexo 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmPreaEstudiov2.frx":0892
            Left            =   1320
            List            =   "frmPreaEstudiov2.frx":089C
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFecNac 
            Height          =   315
            Left            =   4320
            TabIndex        =   25
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   550
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nacimiento"
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
            Index           =   38
            Left            =   3120
            TabIndex        =   27
            Top             =   390
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Genero"
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
            Index           =   37
            Left            =   360
            TabIndex        =   26
            Top             =   390
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin XtremeSuiteControls.TabControl tcMain 
         Height          =   6226
         Left            =   0
         TabIndex        =   1
         Top             =   5160
         Width           =   18135
         _Version        =   1572864
         _ExtentX        =   31988
         _ExtentY        =   10993
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
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   13
         SelectedItem    =   6
         Item(0).Caption =   "Salarios"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "gbSalarios(0)"
         Item(1).Caption =   "Deducciones"
         Item(1).ControlCount=   14
         Item(1).Control(0)=   "fraDCargas"
         Item(1).Control(1)=   "gDeducciones"
         Item(1).Control(2)=   "ShortcutCaption1(4)"
         Item(1).Control(3)=   "txtD_TotalMensual"
         Item(1).Control(4)=   "txtD_TotalColilla"
         Item(1).Control(5)=   "Label1(62)"
         Item(1).Control(6)=   "gbEstudioCIC"
         Item(1).Control(7)=   "cboDeduccion"
         Item(1).Control(8)=   "Label1(64)"
         Item(1).Control(9)=   "txtD_Descripcion"
         Item(1).Control(10)=   "Label1(65)"
         Item(1).Control(11)=   "btnDeduccion"
         Item(1).Control(12)=   "txtD_Monto"
         Item(1).Control(13)=   "Label1(66)"
         Item(2).Caption =   "Créditos"
         Item(2).ControlCount=   10
         Item(2).Control(0)=   "gCuotasCancela"
         Item(2).Control(1)=   "txtC_CuotaCancelaTotal"
         Item(2).Control(2)=   "ShortcutCaption1(6)"
         Item(2).Control(3)=   "Label1(67)"
         Item(2).Control(4)=   "gCuotasCobrar"
         Item(2).Control(5)=   "txtC_CuotaPorCobrarTotal"
         Item(2).Control(6)=   "ShortcutCaption1(7)"
         Item(2).Control(7)=   "Label1(68)"
         Item(2).Control(8)=   "btnCreditos(0)"
         Item(2).Control(9)=   "btnCreditos(1)"
         Item(3).Caption =   "Refundiciones"
         Item(3).ControlCount=   11
         Item(3).Control(0)=   "txtR_TotalCuotas"
         Item(3).Control(1)=   "txtR_TotalRefunde"
         Item(3).Control(2)=   "Label1(82)"
         Item(3).Control(3)=   "txtR_TotalMora"
         Item(3).Control(4)=   "Label1(83)"
         Item(3).Control(5)=   "Label1(84)"
         Item(3).Control(6)=   "btnRefundiciones_Actualiza"
         Item(3).Control(7)=   "ShortcutCaption1(11)"
         Item(3).Control(8)=   "ShortcutCaption1(12)"
         Item(3).Control(9)=   "dtpR_Formaliza"
         Item(3).Control(10)=   "gRefunde"
         Item(4).Caption =   "Desembolsos"
         Item(4).ControlCount=   8
         Item(4).Control(0)=   "ShortcutCaption1(8)"
         Item(4).Control(1)=   "Label1(69)"
         Item(4).Control(2)=   "gbDesembolsos"
         Item(4).Control(3)=   "txtDS_TotalMonto"
         Item(4).Control(4)=   "txtDS_TotalCuota"
         Item(4).Control(5)=   "btnDesembolso(2)"
         Item(4).Control(6)=   "btnDesembolso(0)"
         Item(4).Control(7)=   "gDesembolsos"
         Item(5).Caption =   "Fianzas"
         Item(5).ControlCount=   6
         Item(5).Control(0)=   "gFianzas"
         Item(5).Control(1)=   "btnFianzas_Actualiza"
         Item(5).Control(2)=   "txtF_TotalSaldos"
         Item(5).Control(3)=   "txtF_TotalCuotas"
         Item(5).Control(4)=   "Label1(81)"
         Item(5).Control(5)=   "ShortcutCaption1(10)"
         Item(6).Caption =   "Resumen"
         Item(6).ControlCount=   2
         Item(6).Control(0)=   "gbResumen(0)"
         Item(6).Control(1)=   "gbResumen(1)"
         Item(7).Caption =   "Historial"
         Item(7).ControlCount=   6
         Item(7).Control(0)=   "tcHistorial"
         Item(7).Control(1)=   "ShortcutCaption1(13)"
         Item(7).Control(2)=   "cboEtiquetas"
         Item(7).Control(3)=   "txtEtiqueta_Nota"
         Item(7).Control(4)=   "btnEtiqueta"
         Item(7).Control(5)=   "ShortcutCaption1(19)"
         Item(8).Caption =   "Adjuntos"
         Item(8).ControlCount=   8
         Item(8).Control(0)=   "txtArchivo"
         Item(8).Control(1)=   "btnArchivo"
         Item(8).Control(2)=   "Label1(96)"
         Item(8).Control(3)=   "lswArchivos"
         Item(8).Control(4)=   "btnAdjunto_Guardar"
         Item(8).Control(5)=   "ShortcutCaption1(16)"
         Item(8).Control(6)=   "lblLoading"
         Item(8).Control(7)=   "btnAdjunto_Elimina"
         Item(9).Caption =   "Hipotecario"
         Item(9).ControlCount=   12
         Item(9).Control(0)=   "Label1(86)"
         Item(9).Control(1)=   "btnHipoteca(0)"
         Item(9).Control(2)=   "Label1(87)"
         Item(9).Control(3)=   "btnHipoteca(1)"
         Item(9).Control(4)=   "Label1(88)"
         Item(9).Control(5)=   "btnHipoteca(2)"
         Item(9).Control(6)=   "Label1(89)"
         Item(9).Control(7)=   "btnHipoteca(3)"
         Item(9).Control(8)=   "Label1(90)"
         Item(9).Control(9)=   "btnHipoteca(4)"
         Item(9).Control(10)=   "Label1(91)"
         Item(9).Control(11)=   "txtCFIA_Avaluo"
         Item(10).Caption=   "Prendario"
         Item(10).ControlCount=   2
         Item(10).Control(0)=   "gbPrendario"
         Item(10).Control(1)=   "gbPrendaExamenes"
         Item(11).Caption=   "Resolución"
         Item(11).ControlCount=   2
         Item(11).Control(0)=   "GroupBox4"
         Item(11).Control(1)=   "GroupBox5"
         Item(12).Caption=   "Causas"
         Item(12).ControlCount=   7
         Item(12).Control(0)=   "lsw"
         Item(12).Control(1)=   "optCausas(0)"
         Item(12).Control(2)=   "optCausas(1)"
         Item(12).Control(3)=   "cmdGuardaObservaciones"
         Item(12).Control(4)=   "txtObservaciones"
         Item(12).Control(5)=   "GroupBox3"
         Item(12).Control(6)=   "ShortcutCaption1(17)"
         Begin XtremeSuiteControls.ListView lswArchivos 
            Height          =   3495
            Left            =   -68560
            TabIndex        =   307
            Top             =   1680
            Visible         =   0   'False
            Width           =   13815
            _Version        =   1572864
            _ExtentX        =   24368
            _ExtentY        =   6165
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Checkboxes      =   -1  'True
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   4455
            Left            =   -69880
            TabIndex        =   276
            Top             =   960
            Visible         =   0   'False
            Width           =   12015
            _Version        =   1572864
            _ExtentX        =   21193
            _ExtentY        =   7858
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.GroupBox GroupBox5 
            Height          =   5535
            Left            =   -59680
            TabIndex        =   331
            Top             =   480
            Visible         =   0   'False
            Width           =   7695
            _Version        =   1572864
            _ExtentX        =   13573
            _ExtentY        =   9763
            _StockProps     =   79
            Appearance      =   6
            Begin FPSpreadADO.fpSpread vGrid 
               Height          =   2775
               Left            =   240
               TabIndex        =   332
               Top             =   240
               Width           =   6615
               _Version        =   524288
               _ExtentX        =   11668
               _ExtentY        =   4895
               _StockProps     =   64
               BorderStyle     =   0
               DisplayRowHeaders=   0   'False
               EditEnterAction =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ScrollBars      =   0
               SpreadDesigner  =   "frmPreaEstudiov2.frx":08B5
               AppearanceStyle =   1
            End
            Begin XtremeSuiteControls.FlatEdit txtCumplimientoNotas 
               Height          =   2235
               Left            =   240
               TabIndex        =   333
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   3075
               Width           =   7095
               _Version        =   1572864
               _ExtentX        =   12515
               _ExtentY        =   3942
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
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2
               Appearance      =   6
               FlatStyle       =   -1  'True
               UseVisualStyle  =   0   'False
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox4 
            Height          =   5535
            Left            =   -69880
            TabIndex        =   324
            Top             =   480
            Visible         =   0   'False
            Width           =   10215
            _Version        =   1572864
            _ExtentX        =   18018
            _ExtentY        =   9763
            _StockProps     =   79
            Appearance      =   6
            Begin XtremeSuiteControls.ListView lswAutorizadores 
               Height          =   4215
               Left            =   120
               TabIndex        =   334
               Top             =   1200
               Width           =   9975
               _Version        =   1572864
               _ExtentX        =   17595
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.FlatEdit txtActa 
               Height          =   315
               Left            =   1320
               TabIndex        =   325
               Top             =   360
               Width           =   1935
               _Version        =   1572864
               _ExtentX        =   3413
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
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtActaFecha 
               Height          =   315
               Left            =   7920
               TabIndex        =   326
               Top             =   360
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   550
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
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtActaSesion 
               Height          =   315
               Left            =   4800
               TabIndex        =   327
               Top             =   360
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   550
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
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.RadioButton rbActas 
               Height          =   255
               Index           =   0
               Left            =   1920
               TabIndex        =   335
               Top             =   840
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Resoluciones"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   21
            End
            Begin XtremeSuiteControls.RadioButton rbActas 
               Height          =   255
               Index           =   1
               Left            =   4320
               TabIndex        =   336
               Top             =   840
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Autorizadores"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   21
            End
            Begin XtremeSuiteControls.RadioButton rbActas 
               Height          =   255
               Index           =   2
               Left            =   6480
               TabIndex        =   337
               Top             =   840
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Asistencia"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   21
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
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
               Index           =   35
               Left            =   6960
               TabIndex        =   330
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "No. Acta"
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
               Index           =   9
               Left            =   120
               TabIndex        =   329
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "No. Sesión"
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
               Index           =   44
               Left            =   3600
               TabIndex        =   328
               Top             =   360
               Width           =   1575
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   735
            Left            =   -57760
            TabIndex        =   314
            Top             =   960
            Visible         =   0   'False
            Width           =   5415
            _Version        =   1572864
            _ExtentX        =   9551
            _ExtentY        =   1296
            _StockProps     =   79
            Appearance      =   6
            Begin XtremeSuiteControls.RadioButton optObservacion 
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   315
               Top             =   240
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Analistas de Crédito"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   21
            End
            Begin XtremeSuiteControls.RadioButton optObservacion 
               Height          =   375
               Index           =   1
               Left            =   1920
               TabIndex        =   316
               Top             =   240
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Resolución del Comité"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   21
            End
            Begin XtremeSuiteControls.RadioButton optObservacion 
               Height          =   375
               Index           =   2
               Left            =   3960
               TabIndex        =   317
               Top             =   240
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Junta Directiva"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   21
            End
         End
         Begin XtremeSuiteControls.GroupBox gbPrendario 
            Height          =   3375
            Left            =   -68440
            TabIndex        =   291
            Top             =   960
            Visible         =   0   'False
            Width           =   7335
            _Version        =   1572864
            _ExtentX        =   12938
            _ExtentY        =   5953
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Begin XtremeSuiteControls.FlatEdit txtPrendaValor 
               Height          =   315
               Left            =   4800
               TabIndex        =   299
               Top             =   720
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.PushButton btnPrenda 
               Height          =   450
               Index           =   0
               Left            =   4800
               TabIndex        =   206
               ToolTipText     =   "Eliminar Todas!"
               Top             =   1320
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3625
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "Garantía Prendaria"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Picture         =   "frmPreaEstudiov2.frx":0EA5
            End
            Begin XtremeSuiteControls.PushButton btnPrenda 
               Height          =   450
               Index           =   1
               Left            =   4800
               TabIndex        =   218
               ToolTipText     =   "Eliminar Todas!"
               Top             =   2040
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3625
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "Gastos Honorarios"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Picture         =   "frmPreaEstudiov2.frx":15C5
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Agregar Honorarios de Traspaso y Constitución"
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
               Index           =   95
               Left            =   960
               TabIndex        =   252
               Top             =   2040
               Width           =   3735
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Agregar la Garantía Prendaria"
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
               Index           =   94
               Left            =   960
               TabIndex        =   301
               Top             =   1440
               Width           =   3735
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Agregar el Valor Preliminar de la Prenda"
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
               Index           =   93
               Left            =   960
               TabIndex        =   300
               Top             =   720
               Width           =   3855
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   14
               Left            =   0
               TabIndex        =   292
               Top             =   0
               Width           =   9615
               _Version        =   1572864
               _ExtentX        =   16960
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Garantía Prendaria:"
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
         Begin XtremeSuiteControls.PushButton btnHipoteca 
            Height          =   615
            Index           =   0
            Left            =   -61840
            TabIndex        =   280
            Top             =   960
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Montos Hipoteca"
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.TabControl tcHistorial 
            Height          =   5895
            Left            =   -70000
            TabIndex        =   272
            Top             =   960
            Visible         =   0   'False
            Width           =   12855
            _Version        =   1572864
            _ExtentX        =   22675
            _ExtentY        =   10398
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
            Item(0).Caption =   "Historial Ejecutivos"
            Item(0).ControlCount=   1
            Item(0).Control(0)=   "gH_Ejecutivos"
            Item(1).Caption =   "Historial General"
            Item(1).ControlCount=   1
            Item(1).Control(0)=   "gH_General"
            Begin FPSpreadADO.fpSpread gH_Ejecutivos 
               Height          =   4575
               Left            =   0
               TabIndex        =   275
               Top             =   480
               Width           =   12855
               _Version        =   524288
               _ExtentX        =   22675
               _ExtentY        =   8070
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
               MaxCols         =   6
               SpreadDesigner  =   "frmPreaEstudiov2.frx":1A2E
               VScrollSpecialType=   2
               AppearanceStyle =   1
            End
            Begin FPSpreadADO.fpSpread gH_General 
               Height          =   4575
               Left            =   -70000
               TabIndex        =   311
               Top             =   480
               Visible         =   0   'False
               Width           =   12855
               _Version        =   524288
               _ExtentX        =   22675
               _ExtentY        =   8070
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
               MaxCols         =   6
               SpreadDesigner  =   "frmPreaEstudiov2.frx":2176
               VScrollSpecialType=   2
               AppearanceStyle =   1
            End
         End
         Begin XtremeSuiteControls.GroupBox gbDesembolsos 
            Height          =   5535
            Left            =   -69880
            TabIndex        =   222
            Top             =   600
            Visible         =   0   'False
            Width           =   8655
            _Version        =   1572864
            _ExtentX        =   15266
            _ExtentY        =   9763
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            BorderStyle     =   2
            Begin XtremeSuiteControls.ListView lswD_Lista 
               Height          =   1335
               Left            =   120
               TabIndex        =   224
               Top             =   840
               Width           =   8415
               _Version        =   1572864
               _ExtentX        =   14843
               _ExtentY        =   2355
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
               Appearance      =   21
            End
            Begin XtremeSuiteControls.ComboBox cboD_Ordinario 
               Height          =   330
               Left            =   1200
               TabIndex        =   225
               Top             =   480
               Width           =   1215
               _Version        =   1572864
               _ExtentX        =   2143
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
            Begin XtremeSuiteControls.FlatEdit txtD_Filtro 
               Height          =   330
               Left            =   2400
               TabIndex        =   227
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   480
               Width           =   6135
               _Version        =   1572864
               _ExtentX        =   10821
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
            Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
               Height          =   330
               Left            =   6120
               TabIndex        =   228
               Top             =   3240
               Width           =   2415
               _Version        =   1572864
               _ExtentX        =   4260
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   1973790
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
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.GroupBox GroupBox1 
               Height          =   1935
               Left            =   0
               TabIndex        =   229
               Top             =   3720
               Width           =   9495
               _Version        =   1572864
               _ExtentX        =   16748
               _ExtentY        =   3413
               _StockProps     =   79
               Caption         =   "Datos de la Cuenta Destino"
               ForeColor       =   4210752
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               BorderStyle     =   2
               Begin XtremeSuiteControls.FlatEdit txtIdentificación 
                  Height          =   330
                  Left            =   4920
                  TabIndex        =   230
                  Top             =   0
                  Width           =   2655
                  _Version        =   1572864
                  _ExtentX        =   4683
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
                  Height          =   330
                  Left            =   4920
                  TabIndex        =   231
                  Top             =   360
                  Width           =   2655
                  _Version        =   1572864
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
               Begin XtremeSuiteControls.PushButton btnCuenta 
                  Height          =   315
                  Left            =   7800
                  TabIndex        =   232
                  Top             =   360
                  Width           =   375
                  _Version        =   1572864
                  _ExtentX        =   656
                  _ExtentY        =   550
                  _StockProps     =   79
                  Caption         =   "..."
                  BackColor       =   16777215
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
                  Left            =   1320
                  TabIndex        =   233
                  Top             =   0
                  Width           =   2175
                  _Version        =   1572864
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
                  Left            =   1320
                  TabIndex        =   234
                  Top             =   360
                  Width           =   2175
                  _Version        =   1572864
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
               Begin XtremeSuiteControls.FlatEdit txtEntidad 
                  Height          =   330
                  Left            =   1320
                  TabIndex        =   235
                  Top             =   840
                  Width           =   2175
                  _Version        =   1572864
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
               Begin XtremeSuiteControls.FlatEdit txtCorreo 
                  Height          =   315
                  Left            =   4920
                  TabIndex        =   236
                  Top             =   840
                  Width           =   3615
                  _Version        =   1572864
                  _ExtentX        =   6376
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
               Begin XtremeSuiteControls.FlatEdit txtDetalle 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   237
                  Top             =   1320
                  Width           =   7215
                  _Version        =   1572864
                  _ExtentX        =   12726
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
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cuenta"
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
                  Index           =   78
                  Left            =   3600
                  TabIndex        =   244
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label1 
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
                  Height          =   255
                  Index           =   77
                  Left            =   3600
                  TabIndex        =   243
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
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
                  Index           =   76
                  Left            =   120
                  TabIndex        =   242
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Moneda"
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
                  Index           =   75
                  Left            =   120
                  TabIndex        =   241
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Entidad"
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
                  Index           =   74
                  Left            =   120
                  TabIndex        =   240
                  Top             =   840
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Correo"
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
                  Index           =   73
                  Left            =   3600
                  TabIndex        =   239
                  Top             =   840
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Detalle"
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
                  Index           =   72
                  Left            =   120
                  TabIndex        =   238
                  Top             =   1320
                  Width           =   1215
               End
            End
            Begin XtremeSuiteControls.ComboBox cboBanco 
               Height          =   330
               Left            =   1320
               TabIndex        =   245
               Top             =   3240
               Width           =   4815
               _Version        =   1572864
               _ExtentX        =   8493
               _ExtentY        =   582
               _StockProps     =   77
               ForeColor       =   1973790
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
               Style           =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.FlatEdit txtDS_Descripcion 
               Height          =   315
               Left            =   1320
               TabIndex        =   246
               Top             =   2400
               Width           =   7215
               _Version        =   1572864
               _ExtentX        =   12726
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
            Begin XtremeSuiteControls.FlatEdit txtDS_Monto 
               Height          =   315
               Left            =   1320
               TabIndex        =   268
               Top             =   2760
               Width           =   2175
               _Version        =   1572864
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtDS_Cuota 
               Height          =   315
               Left            =   6360
               TabIndex        =   372
               Top             =   2760
               Width           =   2175
               _Version        =   1572864
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cuota Liberada"
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
               Index           =   41
               Left            =   4560
               TabIndex        =   373
               Top             =   2760
               Width           =   1695
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Monto"
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
               Index           =   71
               Left            =   120
               TabIndex        =   267
               Top             =   2760
               Width           =   1095
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Banco [ Tipo ]"
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
               Index           =   79
               Left            =   120
               TabIndex        =   266
               Top             =   3240
               Width           =   1095
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción"
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
               Index           =   80
               Left            =   120
               TabIndex        =   247
               Top             =   2400
               Width           =   1095
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Ordinario"
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
               Index           =   70
               Left            =   240
               TabIndex        =   226
               Top             =   480
               Width           =   735
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   9
               Left            =   0
               TabIndex        =   223
               Top             =   0
               Width           =   9615
               _Version        =   1572864
               _ExtentX        =   16960
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Registro de Desembolsos:"
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
         Begin XtremeSuiteControls.GroupBox gbEstudioCIC 
            Height          =   1815
            Left            =   -59200
            TabIndex        =   188
            Top             =   3960
            Visible         =   0   'False
            Width           =   4935
            _Version        =   1572864
            _ExtentX        =   8705
            _ExtentY        =   3201
            _StockProps     =   79
            Caption         =   "Estudio CIC"
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
            Appearance      =   21
            Begin XtremeSuiteControls.FlatEdit txtCIC_Puntaje 
               Height          =   330
               Left            =   3240
               TabIndex        =   191
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   720
               Width           =   1575
               _Version        =   1572864
               _ExtentX        =   2778
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
            Begin XtremeSuiteControls.FlatEdit txtCIC_NivelHistorico 
               Height          =   330
               Left            =   3240
               TabIndex        =   192
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   1080
               Width           =   1575
               _Version        =   1572864
               _ExtentX        =   2778
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
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   18
               Left            =   0
               TabIndex        =   338
               Top             =   0
               Width           =   10575
               _Version        =   1572864
               _ExtentX        =   18653
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Estudio CIC"
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
            Begin XtremeSuiteControls.Label Label4 
               Height          =   495
               Index           =   1
               Left            =   360
               TabIndex        =   190
               Top             =   960
               Width           =   2775
               _Version        =   1572864
               _ExtentX        =   4895
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Nivel de Comportamiento Histórico"
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
            Begin XtremeSuiteControls.Label Label4 
               Height          =   495
               Index           =   0
               Left            =   360
               TabIndex        =   189
               Top             =   600
               Width           =   1935
               _Version        =   1572864
               _ExtentX        =   3413
               _ExtentY        =   873
               _StockProps     =   79
               Caption         =   "Puntaje Final del Deudor"
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
         End
         Begin XtremeSuiteControls.GroupBox fraDCargas 
            Height          =   2775
            Left            =   -59200
            TabIndex        =   169
            Top             =   600
            Visible         =   0   'False
            Width           =   4935
            _Version        =   1572864
            _ExtentX        =   8705
            _ExtentY        =   4895
            _StockProps     =   79
            Caption         =   "Cargas"
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
            Appearance      =   21
            Begin XtremeSuiteControls.UpDown UpDownFrap 
               Height          =   315
               Left            =   2025
               TabIndex        =   193
               Top             =   840
               Width           =   270
               _Version        =   1572864
               _ExtentX        =   476
               _ExtentY        =   556
               _StockProps     =   64
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Max             =   10
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtFrapPorc"
               BuddyProperty   =   ""
            End
            Begin XtremeSuiteControls.CheckBox chkCargaAsociacion 
               Height          =   255
               Left            =   3720
               TabIndex        =   175
               Top             =   480
               Width           =   1095
               _Version        =   1572864
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Aplicar"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   21
            End
            Begin XtremeSuiteControls.CheckBox chkCargaFrap 
               Height          =   255
               Left            =   3720
               TabIndex        =   176
               Top             =   840
               Width           =   1095
               _Version        =   1572864
               _ExtentX        =   1931
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Aplicar"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   21
            End
            Begin XtremeSuiteControls.FlatEdit txtD_TotalCargas 
               Height          =   315
               Left            =   1560
               TabIndex        =   185
               Top             =   2160
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3625
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtFrapPorc 
               Height          =   315
               Left            =   1560
               TabIndex        =   187
               Top             =   840
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "0"
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   5
               Left            =   0
               TabIndex        =   194
               Top             =   0
               Width           =   10575
               _Version        =   1572864
               _ExtentX        =   18653
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Cargas Sociales e Impositivas"
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
            Begin VB.Image imgFraCerrar 
               Height          =   240
               Left            =   4680
               Picture         =   "frmPreaEstudiov2.frx":28BE
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Cargas"
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
               Index           =   61
               Left            =   360
               TabIndex        =   186
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "(-) Imp.Salario"
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
               Index           =   23
               Left            =   120
               TabIndex        =   184
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "(-) FAP/FRAP"
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
               Index           =   17
               Left            =   120
               TabIndex        =   183
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "(-) Asociación"
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
               Index           =   15
               Left            =   120
               TabIndex        =   182
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "(-) C.C.S.S."
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
               Index           =   14
               Left            =   120
               TabIndex        =   181
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lblCargaAsociacion 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   5130
                  SubFormatType   =   1
               EndProperty
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
               Height          =   315
               Left            =   1560
               TabIndex        =   180
               Top             =   480
               Width           =   2055
            End
            Begin VB.Label lblCargaFrap 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   5130
                  SubFormatType   =   1
               EndProperty
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
               Height          =   315
               Left            =   2280
               TabIndex        =   179
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label lblCargaCCSS 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   5130
                  SubFormatType   =   1
               EndProperty
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
               Height          =   315
               Left            =   1560
               TabIndex        =   178
               Top             =   1320
               Width           =   2055
            End
            Begin VB.Label lblCargaImpSalario 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   5130
                  SubFormatType   =   1
               EndProperty
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
               Height          =   315
               Left            =   1560
               TabIndex        =   177
               Top             =   1680
               Width           =   2055
            End
         End
         Begin XtremeSuiteControls.GroupBox gbSalarios 
            Height          =   5775
            Index           =   0
            Left            =   -70000
            TabIndex        =   78
            Top             =   360
            Visible         =   0   'False
            Width           =   18135
            _Version        =   1572864
            _ExtentX        =   31988
            _ExtentY        =   10186
            _StockProps     =   79
            BackColor       =   -2147483633
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            BorderStyle     =   2
            Begin XtremeSuiteControls.ListView lswIncapacidades 
               Height          =   1695
               Left            =   12360
               TabIndex        =   371
               Top             =   2280
               Width           =   5535
               _Version        =   1572864
               _ExtentX        =   9763
               _ExtentY        =   2990
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               View            =   3
               FullRowSelect   =   -1  'True
               Appearance      =   21
            End
            Begin XtremeSuiteControls.CheckBox chkS_Constancia 
               Height          =   255
               Left            =   120
               TabIndex        =   163
               Top             =   1680
               Width           =   1935
               _Version        =   1572864
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Salario Constancia"
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
               Appearance      =   21
               Alignment       =   1
            End
            Begin XtremeSuiteControls.DateTimePicker dtpCorte 
               Height          =   330
               Left            =   2280
               TabIndex        =   79
               Top             =   600
               Width           =   1935
               _Version        =   1572864
               _ExtentX        =   3413
               _ExtentY        =   582
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
            Begin XtremeSuiteControls.ComboBox cboSalario 
               Height          =   330
               Left            =   2280
               TabIndex        =   153
               Top             =   240
               Width           =   3615
               _Version        =   1572864
               _ExtentX        =   6376
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
            Begin XtremeSuiteControls.FlatEdit txtS_Devengado 
               Height          =   315
               Left            =   2280
               TabIndex        =   154
               Top             =   960
               Width           =   1935
               _Version        =   1572864
               _ExtentX        =   3413
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtS_Mensual 
               Height          =   315
               Left            =   2280
               TabIndex        =   156
               Top             =   1320
               Width           =   1935
               _Version        =   1572864
               _ExtentX        =   3413
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin FPSpreadADO.fpSpread gExtras 
               Height          =   1815
               Left            =   0
               TabIndex        =   158
               Top             =   3360
               Width           =   5895
               _Version        =   524288
               _ExtentX        =   10398
               _ExtentY        =   3201
               _StockProps     =   64
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
               ScrollBars      =   2
               SpreadDesigner  =   "frmPreaEstudiov2.frx":2FC4
               AppearanceStyle =   1
            End
            Begin XtremeSuiteControls.FlatEdit txtT_Extras 
               Height          =   315
               Left            =   4080
               TabIndex        =   159
               Top             =   5280
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "0.00"
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtS_Privado 
               Height          =   315
               Left            =   2280
               TabIndex        =   161
               Top             =   2520
               Width           =   1935
               _Version        =   1572864
               _ExtentX        =   3413
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.CheckBox chkS_OrdenPatronal 
               Height          =   255
               Left            =   120
               TabIndex        =   164
               Top             =   2040
               Width           =   1935
               _Version        =   1572864
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Sal.Orden Patronal"
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
               Appearance      =   21
               Alignment       =   1
            End
            Begin XtremeSuiteControls.FlatEdit txtS_Constancia 
               Height          =   315
               Left            =   2280
               TabIndex        =   165
               Top             =   1680
               Width           =   1935
               _Version        =   1572864
               _ExtentX        =   3413
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtS_OrdenPatronal 
               Height          =   315
               Left            =   2280
               TabIndex        =   166
               Top             =   2040
               Width           =   1935
               _Version        =   1572864
               _ExtentX        =   3413
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.UpDown UpDownSPrivado 
               Height          =   315
               Left            =   5625
               TabIndex        =   348
               Top             =   2520
               Width           =   270
               _Version        =   1572864
               _ExtentX        =   476
               _ExtentY        =   556
               _StockProps     =   64
               Appearance      =   6
               UseVisualStyle  =   0   'False
               Value           =   100
               AutoBuddy       =   -1  'True
               BuddyControl    =   "txtS_Privado_Porc"
               BuddyProperty   =   ""
            End
            Begin XtremeSuiteControls.FlatEdit txtS_Privado_Porc 
               Height          =   315
               Left            =   4680
               TabIndex        =   349
               Top             =   2520
               Width           =   975
               _Version        =   1572864
               _ExtentX        =   1720
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
               Text            =   "100"
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin FPSpreadADO.fpSpread gSalarios 
               Height          =   5055
               Left            =   6120
               TabIndex        =   350
               Top             =   600
               Width           =   6015
               _Version        =   524288
               _ExtentX        =   10610
               _ExtentY        =   8916
               _StockProps     =   64
               BorderStyle     =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   5
               SpreadDesigner  =   "frmPreaEstudiov2.frx":35BE
               AppearanceStyle =   1
            End
            Begin XtremeSuiteControls.PushButton btnS_Copy 
               Height          =   330
               Index           =   0
               Left            =   10920
               TabIndex        =   351
               ToolTipText     =   "Copia Mes desde el porta papeles"
               Top             =   255
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
               Picture         =   "frmPreaEstudiov2.frx":3C40
            End
            Begin XtremeSuiteControls.PushButton btnS_Copy 
               Height          =   330
               Index           =   1
               Left            =   11280
               TabIndex        =   352
               ToolTipText     =   "Copia Salario RH desde el porta papeles"
               Top             =   255
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
               Picture         =   "frmPreaEstudiov2.frx":4330
            End
            Begin XtremeSuiteControls.PushButton btnS_Copy 
               Height          =   330
               Index           =   2
               Left            =   11640
               TabIndex        =   353
               ToolTipText     =   "Copia CA desde el porta papeles"
               Top             =   255
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
               Picture         =   "frmPreaEstudiov2.frx":4A20
            End
            Begin XtremeSuiteControls.ComboBox cboS_ComponenteAdicional 
               Height          =   330
               Left            =   14760
               TabIndex        =   355
               Top             =   840
               Width           =   3135
               _Version        =   1572864
               _ExtentX        =   5530
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
            Begin XtremeSuiteControls.FlatEdit txtS_ComponenteAdicional 
               Height          =   315
               Left            =   14760
               TabIndex        =   356
               Top             =   1320
               Width           =   2535
               _Version        =   1572864
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.PushButton btnS_Copy 
               Height          =   330
               Index           =   3
               Left            =   17040
               TabIndex        =   357
               ToolTipText     =   "Copia Incapacidades desde el porta papeles"
               Top             =   1935
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
               Picture         =   "frmPreaEstudiov2.frx":5110
            End
            Begin XtremeSuiteControls.PushButton btnS_Copy 
               Height          =   330
               Index           =   4
               Left            =   17400
               TabIndex        =   358
               ToolTipText     =   "Elimina Lista de Incapacidades"
               Top             =   1935
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
               Picture         =   "frmPreaEstudiov2.frx":5800
            End
            Begin XtremeSuiteControls.FlatEdit txtOficina 
               Height          =   315
               Left            =   13200
               TabIndex        =   359
               ToolTipText     =   "Presione F4 para consultar"
               Top             =   4920
               Width           =   4695
               _Version        =   1572864
               _ExtentX        =   8281
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
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.PushButton btnOficinaCambia 
               Height          =   330
               Left            =   16680
               TabIndex        =   360
               ToolTipText     =   "Cambiar de Oficina/Agencia"
               Top             =   5280
               Width           =   1215
               _Version        =   1572864
               _ExtentX        =   2143
               _ExtentY        =   582
               _StockProps     =   79
               Caption         =   "Cambiar"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   21
               Picture         =   "frmPreaEstudiov2.frx":5DA4
               ImageAlignment  =   0
            End
            Begin XtremeSuiteControls.FlatEdit txtEjecutivo 
               Height          =   315
               Left            =   13200
               TabIndex        =   361
               ToolTipText     =   "Presione F4 para consultar"
               Top             =   4440
               Width           =   4695
               _Version        =   1572864
               _ExtentX        =   8281
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
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtS_ComponenteAdicionalPorc 
               Height          =   315
               Left            =   17280
               TabIndex        =   362
               Top             =   1320
               Width           =   615
               _Version        =   1572864
               _ExtentX        =   1085
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
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   2
               Left            =   12360
               TabIndex        =   370
               Top             =   240
               Width           =   5535
               _Version        =   1572864
               _ExtentX        =   9763
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Componentes Adicionales:"
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
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "(%) Componente adicional"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   57
               Left            =   12480
               TabIndex        =   369
               Top             =   840
               Width           =   2055
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "(+) Componentes Adicionales"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   58
               Left            =   12480
               TabIndex        =   368
               Top             =   1320
               Width           =   2295
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   3
               Left            =   12360
               TabIndex        =   367
               Top             =   1920
               Width           =   5535
               _Version        =   1572864
               _ExtentX        =   9763
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Lista de Incapacidades (Información)"
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
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Ejecutivo"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   59
               Left            =   12360
               TabIndex        =   366
               Top             =   4440
               Width           =   975
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Oficina"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   60
               Left            =   12360
               TabIndex        =   365
               Top             =   4920
               Width           =   975
            End
            Begin VB.Label lblRegistro 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Usuario"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   0
               Left            =   13200
               TabIndex        =   364
               Top             =   4080
               Width           =   1935
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblRegistro 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   1
               Left            =   15840
               TabIndex        =   363
               Top             =   4080
               Width           =   1935
               WordWrap        =   -1  'True
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   1
               Left            =   6120
               TabIndex        =   354
               Top             =   240
               Width           =   5895
               _Version        =   1572864
               _ExtentX        =   10398
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Tabla de Salarios:"
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
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   0
               Left            =   0
               TabIndex        =   168
               Top             =   3000
               Width           =   5895
               _Version        =   1572864
               _ExtentX        =   10398
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Extras:"
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
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Extras"
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
               Index           =   56
               Left            =   2760
               TabIndex        =   167
               Top             =   5280
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "%"
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
               Index           =   55
               Left            =   4200
               TabIndex        =   162
               Top             =   2520
               Width           =   375
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Ingreso por actividades privadas"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   54
               Left            =   120
               TabIndex        =   160
               Top             =   2400
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Salario Mensual"
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
               Index           =   53
               Left            =   120
               TabIndex        =   157
               Top             =   1320
               Width           =   1815
            End
            Begin VB.Label Label1 
               Caption         =   "Salario Devengado"
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
               Index           =   52
               Left            =   120
               TabIndex        =   155
               Top             =   960
               Width           =   1575
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Corte de Colilla"
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
               Index           =   18
               Left            =   120
               TabIndex        =   81
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Salario"
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
               Index           =   5
               Left            =   120
               TabIndex        =   80
               Top             =   240
               Width           =   1815
            End
         End
         Begin XtremeSuiteControls.GroupBox gbResumen 
            Height          =   5655
            Index           =   0
            Left            =   0
            TabIndex        =   82
            Top             =   360
            Width           =   9255
            _Version        =   1572864
            _ExtentX        =   16325
            _ExtentY        =   9975
            _StockProps     =   79
            Caption         =   "Cálculos del Estudio"
            ForeColor       =   16711680
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
            Appearance      =   21
            Begin XtremeSuiteControls.FlatEdit txtSalarioReal 
               Height          =   315
               Left            =   2280
               TabIndex        =   84
               Top             =   2040
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   550
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
            Begin XtremeSuiteControls.FlatEdit txtCompAdicionalBase 
               Height          =   315
               Left            =   2280
               TabIndex        =   85
               Top             =   2400
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   550
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtDevengadoMes 
               Height          =   315
               Left            =   2280
               TabIndex        =   86
               Top             =   3120
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   550
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
            Begin XtremeSuiteControls.FlatEdit txtTotal_Cargas_CCSS 
               Height          =   315
               Left            =   2280
               TabIndex        =   87
               Top             =   3480
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPorcSobreSalario 
               Height          =   315
               Left            =   2280
               TabIndex        =   88
               Top             =   3840
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtDeducciones 
               Height          =   315
               Left            =   2280
               TabIndex        =   89
               Top             =   4200
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCrdTransitoCancelados 
               Height          =   315
               Left            =   2280
               TabIndex        =   90
               Top             =   4560
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCrdTransitoXCobrar 
               Height          =   315
               Left            =   2280
               TabIndex        =   91
               Top             =   4920
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtRebajoExtras 
               Height          =   315
               Left            =   2280
               TabIndex        =   92
               Top             =   1680
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtSalarioDevengado 
               Height          =   315
               Left            =   2280
               TabIndex        =   93
               Top             =   1320
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   550
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
            Begin XtremeSuiteControls.FlatEdit txtTipoSalario 
               Height          =   315
               Left            =   2280
               TabIndex        =   106
               Top             =   480
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   550
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
            Begin XtremeSuiteControls.FlatEdit txtColillaCorte 
               Height          =   315
               Left            =   2280
               TabIndex        =   107
               Top             =   840
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   550
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
            Begin XtremeSuiteControls.FlatEdit txtFianzas 
               Height          =   315
               Left            =   6600
               TabIndex        =   108
               Top             =   2280
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtMontoGirar 
               Height          =   315
               Left            =   6600
               TabIndex        =   109
               Top             =   4320
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCuotaDiferencia 
               Height          =   315
               Left            =   6600
               TabIndex        =   110
               Top             =   4680
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtTotalLiquido 
               Height          =   315
               Left            =   6600
               TabIndex        =   111
               Top             =   1560
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtRefundiciones 
               Height          =   315
               Left            =   6600
               TabIndex        =   112
               Top             =   840
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtDesembolsos 
               Height          =   315
               Left            =   6600
               TabIndex        =   113
               Top             =   1200
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtSalarioLiquido 
               Height          =   315
               Left            =   6600
               TabIndex        =   114
               Top             =   480
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtTotalLiquidoGrupo 
               Height          =   315
               Left            =   6600
               TabIndex        =   115
               Top             =   1920
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtComisiones 
               Height          =   315
               Left            =   6600
               TabIndex        =   131
               Top             =   3120
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtIntereses 
               Height          =   315
               Left            =   6600
               TabIndex        =   133
               Top             =   3480
               Width           =   2055
               _Version        =   1572864
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPSD 
               Height          =   315
               Left            =   6600
               TabIndex        =   135
               Top             =   3840
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3625
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCompAdicional 
               Height          =   315
               Left            =   2280
               TabIndex        =   138
               Top             =   2760
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   550
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
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.PushButton btnIntereses 
               Height          =   315
               Left            =   8640
               TabIndex        =   152
               ToolTipText     =   "Cambio de Evaluador"
               Top             =   3480
               Width           =   315
               _Version        =   1572864
               _ExtentX        =   556
               _ExtentY        =   556
               _StockProps     =   79
               BackColor       =   -2147483633
               FlatStyle       =   -1  'True
               Appearance      =   16
               Picture         =   "frmPreaEstudiov2.frx":64BD
               ImageAlignment  =   6
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   20
               Left            =   0
               TabIndex        =   342
               Top             =   0
               Width           =   9255
               _Version        =   1572864
               _ExtentX        =   16325
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Cálculos del Estudio de Crédito"
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
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(+ %) Comp.Adicional"
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
               Index           =   49
               Left            =   240
               TabIndex        =   139
               Top             =   2760
               Width           =   2055
            End
            Begin VB.Label lblSalarioDevengado 
               BackStyle       =   0  'Transparent
               Caption         =   "Formalización:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   0
               Left            =   4560
               TabIndex        =   137
               Top             =   2760
               Width           =   1935
            End
            Begin VB.Label Label1 
               Caption         =   "P.S.D."
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
               Height          =   255
               Index           =   42
               Left            =   4920
               TabIndex        =   136
               Top             =   3840
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Intereses"
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
               Index           =   48
               Left            =   4920
               TabIndex        =   134
               Top             =   3480
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Comisiones"
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
               Index           =   47
               Left            =   4920
               TabIndex        =   132
               Top             =   3120
               Width           =   1575
            End
            Begin VB.Label lblMontoGirar 
               Caption         =   "Monto a Girar"
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
               Index           =   34
               Left            =   4920
               TabIndex        =   123
               Top             =   4320
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Fianzas"
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
               Index           =   27
               Left            =   4920
               TabIndex        =   122
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Label lblMontoGirar 
               Caption         =   "Diferencia Cuota.:"
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
               Left            =   4920
               TabIndex        =   121
               Top             =   4680
               Width           =   1575
            End
            Begin VB.Image imgCuotaDif 
               Height          =   240
               Left            =   8760
               Stretch         =   -1  'True
               ToolTipText     =   "Causas"
               Top             =   4680
               Width           =   240
            End
            Begin VB.Label Label1 
               Caption         =   "T. Liquido Persona"
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
               Height          =   255
               Index           =   26
               Left            =   4920
               TabIndex        =   120
               Top             =   1560
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "(+) Refundiciones"
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
               Index           =   25
               Left            =   4920
               TabIndex        =   119
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "(+) Desembolsos"
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
               Index           =   29
               Left            =   4920
               TabIndex        =   118
               Top             =   1200
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Salario Liquido"
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
               Index           =   24
               Left            =   4920
               TabIndex        =   117
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "T. Liquido Grupo"
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
               Height          =   255
               Index           =   0
               Left            =   4920
               TabIndex        =   116
               ToolTipText     =   "Es el Total Liquido del Deudor + Co Deudores"
               Top             =   1920
               Width           =   1575
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Corte de Colilla"
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
               Index           =   46
               Left            =   240
               TabIndex        =   105
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Salario"
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
               Index           =   45
               Left            =   240
               TabIndex        =   104
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(-) Créditos x Cobrar"
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
               Index           =   22
               Left            =   240
               TabIndex        =   103
               Top             =   4920
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(+) Créditos Cancelados"
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
               Index           =   21
               Left            =   240
               TabIndex        =   102
               Top             =   4560
               Width           =   2055
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(-) Deducciones"
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
               Index           =   20
               Left            =   240
               TabIndex        =   101
               Top             =   4200
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(-) Cargas"
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
               Index           =   19
               Left            =   240
               TabIndex        =   100
               Top             =   3480
               Width           =   1815
            End
            Begin VB.Label lblPorcentajeSalario 
               BackStyle       =   0  'Transparent
               Caption         =   "(%?) Sobre Salario"
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
               Height          =   255
               Left            =   240
               TabIndex        =   99
               Top             =   3840
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Devengado del Mes"
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
               Index           =   13
               Left            =   240
               TabIndex        =   98
               Top             =   3120
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(+) Comp.Adicionales"
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
               Index           =   12
               Left            =   240
               TabIndex        =   97
               Top             =   2400
               Width           =   2055
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Salario Real"
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
               Height          =   255
               Index           =   11
               Left            =   240
               TabIndex        =   96
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "(-) Rebajo de Extras"
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
               Index           =   10
               Left            =   240
               TabIndex        =   95
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label lblSalarioDevengado 
               BackStyle       =   0  'Transparent
               Caption         =   "Salario Devengado"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   9
               Left            =   240
               TabIndex        =   94
               Top             =   1320
               Width           =   1935
            End
         End
         Begin XtremeSuiteControls.GroupBox gbResumen 
            Height          =   5655
            Index           =   1
            Left            =   9240
            TabIndex        =   83
            Top             =   360
            Width           =   8895
            _Version        =   1572864
            _ExtentX        =   15690
            _ExtentY        =   9975
            _StockProps     =   79
            Caption         =   "Información de la Liquidez de la Persona"
            ForeColor       =   16711680
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
            Appearance      =   21
            Begin XtremeSuiteControls.PushButton btnResumen 
               Height          =   615
               Index           =   0
               Left            =   6000
               TabIndex        =   149
               Top             =   2400
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "Boleta de Estudio de Crédito"
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
               Appearance      =   21
               Picture         =   "frmPreaEstudiov2.frx":6BB8
            End
            Begin XtremeSuiteControls.FlatEdit txtLiquidezSinFianza 
               Height          =   315
               Left            =   1800
               TabIndex        =   124
               Top             =   960
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2350
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtLiquidezPorcSinFianza 
               Height          =   315
               Left            =   3120
               TabIndex        =   125
               Top             =   960
               Width           =   855
               _Version        =   1572864
               _ExtentX        =   1508
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtLiquidezConFianza 
               Height          =   315
               Left            =   1800
               TabIndex        =   126
               Top             =   1440
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2350
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtLiquidezPorcConFianza 
               Height          =   315
               Left            =   3120
               TabIndex        =   127
               Top             =   1440
               Width           =   855
               _Version        =   1572864
               _ExtentX        =   1508
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtLiquidezSinFianzaComp 
               Height          =   315
               Left            =   6000
               TabIndex        =   140
               Top             =   960
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2350
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtLiquidezPorcSinFianzaComp 
               Height          =   315
               Left            =   7320
               TabIndex        =   141
               Top             =   960
               Width           =   855
               _Version        =   1572864
               _ExtentX        =   1508
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtLiquidezConFianzaComp 
               Height          =   315
               Left            =   6000
               TabIndex        =   142
               Top             =   1440
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2350
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtLiquidezPorcConFianzaComp 
               Height          =   315
               Left            =   7320
               TabIndex        =   143
               Top             =   1440
               Width           =   855
               _Version        =   1572864
               _ExtentX        =   1508
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtSalarioMinInembargableEstudio 
               Height          =   315
               Left            =   2160
               TabIndex        =   148
               Top             =   2640
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.PushButton btnResumen 
               Height          =   615
               Index           =   1
               Left            =   6000
               TabIndex        =   150
               Top             =   3120
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "Deducciones de Planilla"
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
               Appearance      =   21
               Picture         =   "frmPreaEstudiov2.frx":72BF
            End
            Begin XtremeSuiteControls.PushButton btnResumen 
               Height          =   615
               Index           =   2
               Left            =   6000
               TabIndex        =   151
               Top             =   3840
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "Recuperación Mora (Ballon Payment)"
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
               Appearance      =   21
               Picture         =   "frmPreaEstudiov2.frx":79BA
            End
            Begin XtremeSuiteControls.FlatEdit txtSalarioNormativaEstudio 
               Height          =   315
               Left            =   2160
               TabIndex        =   339
               Top             =   3600
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
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
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   21
               Left            =   0
               TabIndex        =   343
               Top             =   0
               Width           =   9135
               _Version        =   1572864
               _ExtentX        =   16113
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Información de la Liquidez de la Persona"
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
            Begin VB.Label Label1 
               Caption         =   "Salario Normativa de este Estudio de Crédito:"
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
               Index           =   97
               Left            =   120
               TabIndex        =   340
               Top             =   3240
               Width           =   4695
            End
            Begin VB.Label Label1 
               Caption         =   "Salario Mínimo Inembargable de este Estudio de Crédito:"
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
               Index           =   51
               Left            =   120
               TabIndex        =   147
               Top             =   2280
               Width           =   4695
            End
            Begin VB.Label Label1 
               Caption         =   "Con Fianzas"
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
               Index           =   50
               Left            =   4560
               TabIndex        =   146
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "Sin Fianzas"
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
               Index           =   33
               Left            =   4560
               TabIndex        =   145
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Liquidez Considerando Componente Adicional (%)"
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
               Index           =   32
               Left            =   4320
               TabIndex        =   144
               Top             =   480
               Width           =   4215
            End
            Begin VB.Label Label1 
               Caption         =   "Con Fianzas"
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
               Left            =   360
               TabIndex        =   130
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "Sin Fianzas"
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
               Index           =   30
               Left            =   360
               TabIndex        =   129
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Liquidez sin considerar Componente Adicional (%)"
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
               Index           =   28
               Left            =   120
               TabIndex        =   128
               Top             =   480
               Width           =   4695
            End
         End
         Begin FPSpreadADO.fpSpread gDeducciones 
            Height          =   3495
            Left            =   -70000
            TabIndex        =   170
            Top             =   1920
            Visible         =   0   'False
            Width           =   10575
            _Version        =   524288
            _ExtentX        =   18653
            _ExtentY        =   6165
            _StockProps     =   64
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
            MaxCols         =   5
            ScrollBars      =   2
            SpreadDesigner  =   "frmPreaEstudiov2.frx":7E23
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtD_TotalMensual 
            Height          =   315
            Left            =   -61360
            TabIndex        =   172
            Top             =   5520
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtD_TotalColilla 
            Height          =   315
            Left            =   -63160
            TabIndex        =   173
            Top             =   5520
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboDeduccion 
            Height          =   330
            Left            =   -67960
            TabIndex        =   197
            Top             =   1080
            Visible         =   0   'False
            Width           =   5175
            _Version        =   1572864
            _ExtentX        =   9128
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
         Begin XtremeSuiteControls.FlatEdit txtD_Descripcion 
            Height          =   330
            Left            =   -67960
            TabIndex        =   199
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1440
            Visible         =   0   'False
            Width           =   5175
            _Version        =   1572864
            _ExtentX        =   9128
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
         Begin XtremeSuiteControls.PushButton btnDeduccion 
            Height          =   330
            Left            =   -60640
            TabIndex        =   201
            ToolTipText     =   "Guardar Deduccióm"
            Top             =   1440
            Visible         =   0   'False
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
            Picture         =   "frmPreaEstudiov2.frx":8536
         End
         Begin XtremeSuiteControls.FlatEdit txtD_Monto 
            Height          =   330
            Left            =   -62800
            TabIndex        =   202
            Top             =   1440
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
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
         Begin FPSpreadADO.fpSpread gCuotasCancela 
            Height          =   4455
            Left            =   -69880
            TabIndex        =   204
            Top             =   960
            Visible         =   0   'False
            Width           =   7935
            _Version        =   524288
            _ExtentX        =   13996
            _ExtentY        =   7858
            _StockProps     =   64
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
            ScrollBars      =   2
            SpreadDesigner  =   "frmPreaEstudiov2.frx":8C56
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtC_CuotaCancelaTotal 
            Height          =   315
            Left            =   -64000
            TabIndex        =   205
            Top             =   5520
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
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin FPSpreadADO.fpSpread gCuotasCobrar 
            Height          =   4455
            Left            =   -61120
            TabIndex        =   209
            Top             =   960
            Visible         =   0   'False
            Width           =   7935
            _Version        =   524288
            _ExtentX        =   13996
            _ExtentY        =   7858
            _StockProps     =   64
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
            ScrollBars      =   2
            SpreadDesigner  =   "frmPreaEstudiov2.frx":9234
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtC_CuotaPorCobrarTotal 
            Height          =   315
            Left            =   -55240
            TabIndex        =   210
            Top             =   5520
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
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCreditos 
            Height          =   330
            Index           =   0
            Left            =   -69400
            TabIndex        =   213
            ToolTipText     =   "Eliminar Todas!"
            Top             =   5520
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Eliminar Todas!"
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
            Picture         =   "frmPreaEstudiov2.frx":9812
         End
         Begin XtremeSuiteControls.PushButton btnCreditos 
            Height          =   330
            Index           =   1
            Left            =   -60640
            TabIndex        =   214
            ToolTipText     =   "Eliminar Todas!"
            Top             =   5520
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Eliminar Todas!"
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
            Picture         =   "frmPreaEstudiov2.frx":9DB6
         End
         Begin FPSpreadADO.fpSpread gDesembolsos 
            Height          =   4575
            Left            =   -61120
            TabIndex        =   215
            Top             =   960
            Visible         =   0   'False
            Width           =   9255
            _Version        =   524288
            _ExtentX        =   16325
            _ExtentY        =   8070
            _StockProps     =   64
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
            MaxCols         =   5
            ScrollBars      =   2
            SpreadDesigner  =   "frmPreaEstudiov2.frx":A35A
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtDS_TotalMonto 
            Height          =   315
            Left            =   -53680
            TabIndex        =   216
            Top             =   5640
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDS_TotalCuota 
            Height          =   315
            Left            =   -55120
            TabIndex        =   217
            Top             =   5640
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
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
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnDesembolso 
            Height          =   330
            Index           =   2
            Left            =   -60280
            TabIndex        =   221
            ToolTipText     =   "Eliminar Todas!"
            Top             =   5640
            Visible         =   0   'False
            Width           =   2535
            _Version        =   1572864
            _ExtentX        =   4471
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Eliminar Seleccionados!"
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
            Picture         =   "frmPreaEstudiov2.frx":AA37
         End
         Begin FPSpreadADO.fpSpread gFianzas 
            Height          =   3975
            Left            =   -69040
            TabIndex        =   248
            Top             =   960
            Visible         =   0   'False
            Width           =   14175
            _Version        =   524288
            _ExtentX        =   25003
            _ExtentY        =   7011
            _StockProps     =   64
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
            MaxCols         =   11
            ScrollBars      =   2
            SpreadDesigner  =   "frmPreaEstudiov2.frx":AFDB
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.PushButton btnFianzas_Actualiza 
            Height          =   495
            Left            =   -54040
            TabIndex        =   249
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Actualizar"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Picture         =   "frmPreaEstudiov2.frx":BA47
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtF_TotalSaldos 
            Height          =   315
            Left            =   -66880
            TabIndex        =   250
            Top             =   5160
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtF_TotalCuotas 
            Height          =   315
            Left            =   -65320
            TabIndex        =   251
            Top             =   5160
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
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
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin FPSpreadADO.fpSpread gRefunde 
            Height          =   4335
            Left            =   -69880
            TabIndex        =   255
            Top             =   1080
            Visible         =   0   'False
            Width           =   18015
            _Version        =   524288
            _ExtentX        =   31776
            _ExtentY        =   7646
            _StockProps     =   64
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
            MaxCols         =   18
            SpreadDesigner  =   "frmPreaEstudiov2.frx":C147
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtR_TotalCuotas 
            Height          =   315
            Left            =   -61840
            TabIndex        =   256
            Top             =   5520
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
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtR_TotalRefunde 
            Height          =   315
            Left            =   -65440
            TabIndex        =   257
            Top             =   5520
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
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtR_TotalMora 
            Height          =   315
            Left            =   -58120
            TabIndex        =   259
            Top             =   5520
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
            Text            =   "0.00"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnRefundiciones_Actualiza 
            Height          =   495
            Left            =   -53920
            TabIndex        =   262
            Top             =   5400
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Actualizar"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Picture         =   "frmPreaEstudiov2.frx":CDA0
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.DateTimePicker dtpR_Formaliza 
            Height          =   330
            Left            =   -53920
            TabIndex        =   265
            Top             =   620
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   582
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
            Format          =   3
         End
         Begin XtremeSuiteControls.PushButton btnDesembolso 
            Height          =   330
            Index           =   0
            Left            =   -61120
            TabIndex        =   269
            ToolTipText     =   "Eliminar Todas!"
            Top             =   5640
            Visible         =   0   'False
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
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
            Picture         =   "frmPreaEstudiov2.frx":D4A0
         End
         Begin XtremeSuiteControls.ComboBox cboEtiquetas 
            Height          =   330
            Left            =   -57040
            TabIndex        =   273
            Top             =   960
            Visible         =   0   'False
            Width           =   4935
            _Version        =   1572864
            _ExtentX        =   8705
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
         Begin XtremeSuiteControls.RadioButton optCausas 
            Height          =   495
            Index           =   0
            Left            =   -66520
            TabIndex        =   277
            Top             =   480
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4043
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Causas para Denegación"
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
            Appearance      =   21
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optCausas 
            Height          =   495
            Index           =   1
            Left            =   -63880
            TabIndex        =   278
            Top             =   480
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4043
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Pendientes para Estudio"
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.PushButton btnHipoteca 
            Height          =   615
            Index           =   1
            Left            =   -61840
            TabIndex        =   282
            Top             =   1800
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Sumar Avalúo CFIA"
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.PushButton btnHipoteca 
            Height          =   615
            Index           =   2
            Left            =   -61840
            TabIndex        =   284
            Top             =   2640
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Garantía Hipoteca"
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.PushButton btnHipoteca 
            Height          =   615
            Index           =   3
            Left            =   -61840
            TabIndex        =   286
            Top             =   3480
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Asignar Ingenieros"
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.PushButton btnHipoteca 
            Height          =   615
            Index           =   4
            Left            =   -61840
            TabIndex        =   288
            Top             =   4320
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Cambia el Estado del Estudio de Crédito"
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.FlatEdit txtCFIA_Avaluo 
            Height          =   315
            Left            =   -57160
            TabIndex        =   290
            Top             =   1920
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.GroupBox gbPrendaExamenes 
            Height          =   3375
            Left            =   -60880
            TabIndex        =   293
            Top             =   960
            Visible         =   0   'False
            Width           =   7335
            _Version        =   1572864
            _ExtentX        =   12938
            _ExtentY        =   5953
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Begin XtremeSuiteControls.ListView lswP_Examenes 
               Height          =   2055
               Left            =   240
               TabIndex        =   298
               Top             =   1080
               Width           =   6855
               _Version        =   1572864
               _ExtentX        =   12091
               _ExtentY        =   3625
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               View            =   3
               FullRowSelect   =   -1  'True
               Appearance      =   21
            End
            Begin XtremeSuiteControls.ComboBox cboP_Examenes 
               Height          =   330
               Left            =   2040
               TabIndex        =   295
               Top             =   600
               Width           =   2295
               _Version        =   1572864
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
            Begin XtremeSuiteControls.PushButton btnP_Examenes 
               Height          =   330
               Left            =   4440
               TabIndex        =   297
               ToolTipText     =   "Eliminar Todas!"
               Top             =   600
               Width           =   855
               _Version        =   1572864
               _ExtentX        =   1508
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
               Picture         =   "frmPreaEstudiov2.frx":DBC0
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Estado de Exámenes"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   92
               Left            =   240
               TabIndex        =   296
               Top             =   600
               Width           =   1575
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   375
               Index           =   15
               Left            =   0
               TabIndex        =   294
               Top             =   0
               Width           =   9615
               _Version        =   1572864
               _ExtentX        =   16960
               _ExtentY        =   661
               _StockProps     =   14
               Caption         =   "Exámenes Médicos:"
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
         Begin XtremeSuiteControls.FlatEdit txtArchivo 
            Height          =   435
            Left            =   -67600
            TabIndex        =   303
            Top             =   600
            Visible         =   0   'False
            Width           =   7215
            _Version        =   1572864
            _ExtentX        =   12726
            _ExtentY        =   767
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnArchivo 
            Height          =   375
            Left            =   -60280
            TabIndex        =   304
            Top             =   600
            Visible         =   0   'False
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   656
            _StockProps     =   79
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmPreaEstudiov2.frx":E2E0
         End
         Begin XtremeSuiteControls.PushButton btnAdjunto_Guardar 
            Height          =   375
            Left            =   -59800
            TabIndex        =   308
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Guardar"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmPreaEstudiov2.frx":E9E0
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton btnAdjunto_Elimina 
            Height          =   375
            Left            =   -57400
            TabIndex        =   310
            Top             =   600
            Visible         =   0   'False
            Width           =   2535
            _Version        =   1572864
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Eliminar Selecionados"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmPreaEstudiov2.frx":F0E8
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton cmdGuardaObservaciones 
            Height          =   495
            Left            =   -54160
            TabIndex        =   312
            Top             =   5640
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "&Guardar"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   21
            Picture         =   "frmPreaEstudiov2.frx":F68C
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtObservaciones 
            Height          =   3735
            Left            =   -57760
            TabIndex        =   313
            Top             =   1680
            Visible         =   0   'False
            Width           =   5415
            _Version        =   1572864
            _ExtentX        =   9551
            _ExtentY        =   6588
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
         Begin XtremeSuiteControls.PushButton btnEtiqueta 
            Height          =   495
            Left            =   -53680
            TabIndex        =   320
            Top             =   5400
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "&Guardar"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            Appearance      =   21
            Picture         =   "frmPreaEstudiov2.frx":FDBD
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtEtiqueta_Nota 
            Height          =   3975
            Left            =   -57040
            TabIndex        =   319
            Top             =   1320
            Visible         =   0   'False
            Width           =   4935
            _Version        =   1572864
            _ExtentX        =   8705
            _ExtentY        =   7011
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   19
            Left            =   -57040
            TabIndex        =   341
            Top             =   600
            Visible         =   0   'False
            Width           =   4935
            _Version        =   1572864
            _ExtentX        =   8705
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Tipos de Etiquetas u Observación"
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   17
            Left            =   -57760
            TabIndex        =   318
            Top             =   600
            Visible         =   0   'False
            Width           =   5415
            _Version        =   1572864
            _ExtentX        =   9551
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Observaciones Complementarias:"
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
         Begin XtremeSuiteControls.Label lblLoading 
            Height          =   735
            Left            =   -68560
            TabIndex        =   305
            Top             =   5280
            Visible         =   0   'False
            Width           =   13815
            _Version        =   1572864
            _ExtentX        =   24368
            _ExtentY        =   1296
            _StockProps     =   79
            Caption         =   "xxxxxxxxxx"
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   8
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   16
            Left            =   -68560
            TabIndex        =   309
            Top             =   1320
            Visible         =   0   'False
            Width           =   13815
            _Version        =   1572864
            _ExtentX        =   24368
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Documentos Registrados"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   96
            Left            =   -68560
            TabIndex        =   306
            Top             =   600
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Avalúo Factor CFIA: "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   91
            Left            =   -59560
            TabIndex        =   289
            Top             =   1920
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cambiar el Estado del Estudio de Crédito (Pre Análisis), Comité Revisión Avalúo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   90
            Left            =   -67240
            TabIndex        =   287
            Top             =   4440
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Asignar Ingenieros a la Garantía Hipotecaria"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   89
            Left            =   -67240
            TabIndex        =   285
            Top             =   3600
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Agregar la Garantía Hipotecaria al Estudio de Crédito"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   88
            Left            =   -67240
            TabIndex        =   283
            Top             =   2760
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sumar Avalúo Factor CFIA:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   87
            Left            =   -67240
            TabIndex        =   281
            Top             =   1920
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Agregar Montos por Hipotecas, Cancelación y Traspaso:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   86
            Left            =   -67240
            TabIndex        =   279
            Top             =   1080
            Visible         =   0   'False
            Width           =   4695
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   13
            Left            =   -70000
            TabIndex        =   274
            Top             =   600
            Visible         =   0   'False
            Width           =   12855
            _Version        =   1572864
            _ExtentX        =   22675
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Historial de Estudio de Crédito"
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   12
            Left            =   -55840
            TabIndex        =   264
            Top             =   600
            Visible         =   0   'False
            Width           =   3975
            _Version        =   1572864
            _ExtentX        =   7011
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Fecha Formaliza:"
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   11
            Left            =   -70000
            TabIndex        =   263
            Top             =   600
            Visible         =   0   'False
            Width           =   14175
            _Version        =   1572864
            _ExtentX        =   25003
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Listado de Operaciones para Refundir:"
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Mora"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   84
            Left            =   -59920
            TabIndex        =   261
            Top             =   5520
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Cuotas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   83
            Left            =   -63520
            TabIndex        =   260
            Top             =   5520
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Refundición"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   82
            Left            =   -67240
            TabIndex        =   258
            Top             =   5520
            Visible         =   0   'False
            Width           =   1575
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   10
            Left            =   -69040
            TabIndex        =   254
            Top             =   600
            Visible         =   0   'False
            Width           =   14055
            _Version        =   1572864
            _ExtentX        =   24791
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Listado de Fianzas de la Persona:"
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   81
            Left            =   -68200
            TabIndex        =   253
            Top             =   5160
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   69
            Left            =   -56440
            TabIndex        =   220
            Top             =   5640
            Visible         =   0   'False
            Width           =   1095
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   8
            Left            =   -61120
            TabIndex        =   219
            Top             =   600
            Visible         =   0   'False
            Width           =   9135
            _Version        =   1572864
            _ExtentX        =   16113
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Desembolsos:"
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   68
            Left            =   -56440
            TabIndex        =   212
            Top             =   5520
            Visible         =   0   'False
            Width           =   1095
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   7
            Left            =   -61120
            TabIndex        =   211
            Top             =   600
            Visible         =   0   'False
            Width           =   7935
            _Version        =   1572864
            _ExtentX        =   13996
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Cuota de Créditos por Cobrar:"
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   67
            Left            =   -65200
            TabIndex        =   208
            Top             =   5520
            Visible         =   0   'False
            Width           =   1095
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   6
            Left            =   -69880
            TabIndex        =   207
            Top             =   600
            Visible         =   0   'False
            Width           =   7935
            _Version        =   1572864
            _ExtentX        =   13996
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Cuota de Créditos Cancelados:"
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Monto"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   66
            Left            =   -61960
            TabIndex        =   203
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   65
            Left            =   -69040
            TabIndex        =   200
            Top             =   1440
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Deducción"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   64
            Left            =   -69040
            TabIndex        =   198
            Top             =   1080
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   62
            Left            =   -64000
            TabIndex        =   174
            Top             =   5520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   4
            Left            =   -70000
            TabIndex        =   171
            Top             =   600
            Visible         =   0   'False
            Width           =   10575
            _Version        =   1572864
            _ExtentX        =   18653
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Lista de Deducciones:"
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
      Begin XtremeSuiteControls.PushButton btnSolicitado 
         Height          =   330
         Left            =   8640
         TabIndex        =   3
         ToolTipText     =   "Poner en Estado de Solicitado"
         Top             =   30
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Volver a Solicitado"
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   21
         Picture         =   "frmPreaEstudiov2.frx":104EE
      End
      Begin XtremeSuiteControls.PushButton btnCopiar 
         Height          =   330
         Left            =   10440
         TabIndex        =   4
         ToolTipText     =   "Copiar Expediente"
         Top             =   30
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1714
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Copia"
         BackColor       =   16761024
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmPreaEstudiov2.frx":10BEE
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnGestion 
         Height          =   330
         Index           =   0
         Left            =   11880
         TabIndex        =   5
         Top             =   30
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Causas"
         BackColor       =   16761024
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
         TextAlignment   =   1
         Appearance      =   21
         Picture         =   "frmPreaEstudiov2.frx":112DE
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnGestion 
         Height          =   330
         Index           =   1
         Left            =   13080
         TabIndex        =   6
         Top             =   30
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Etiquetas"
         BackColor       =   16761024
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
         TextAlignment   =   1
         Appearance      =   21
         Picture         =   "frmPreaEstudiov2.frx":119E5
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnGestion 
         Height          =   330
         Index           =   2
         Left            =   14640
         TabIndex        =   7
         ToolTipText     =   "Resuelve"
         Top             =   30
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Resolución"
         BackColor       =   16761024
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
         TextAlignment   =   1
         Appearance      =   21
         Picture         =   "frmPreaEstudiov2.frx":120FE
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnGestion 
         Height          =   330
         Index           =   3
         Left            =   15960
         TabIndex        =   8
         Top             =   30
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Solicitud de Crédito"
         BackColor       =   16761024
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
         TextAlignment   =   1
         Appearance      =   21
         Picture         =   "frmPreaEstudiov2.frx":12825
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtExpediente 
         Height          =   405
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
         _ExtentY        =   714
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
      Begin XtremeSuiteControls.ComboBox cboSubExpediente 
         Height          =   390
         Left            =   4440
         TabIndex        =   11
         Top             =   600
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   714
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.GroupBox gbDatosPersonales 
         Height          =   1455
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   10335
         _Version        =   1572864
         _ExtentX        =   18230
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Datos Personales"
         ForeColor       =   8421504
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
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCedula 
            Height          =   330
            Left            =   1680
            TabIndex        =   14
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   360
            Width           =   1815
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   330
            Left            =   3480
            TabIndex        =   15
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   360
            Width           =   6615
            _Version        =   1572864
            _ExtentX        =   11668
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
         Begin XtremeSuiteControls.PushButton btnEdad 
            Height          =   330
            Left            =   3120
            TabIndex        =   29
            ToolTipText     =   "Justificación para la Edad"
            Top             =   1080
            Width           =   255
            _Version        =   1572864
            _ExtentX        =   450
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
            Picture         =   "frmPreaEstudiov2.frx":130F6
         End
         Begin XtremeSuiteControls.FlatEdit txtClasificacion 
            Height          =   315
            Left            =   8520
            TabIndex        =   344
            Top             =   720
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEstadoSocio 
            Height          =   315
            Left            =   3480
            TabIndex        =   346
            Top             =   720
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEdad 
            Height          =   315
            Left            =   3480
            TabIndex        =   347
            Top             =   1080
            Width           =   6615
            _Version        =   1572864
            _ExtentX        =   11668
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificación Crediticia"
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
            Index           =   40
            Left            =   6480
            TabIndex        =   345
            Top             =   720
            Width           =   2175
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   322
            Top             =   1080
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Edad"
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
            Index           =   0
            Left            =   1680
            TabIndex        =   28
            Top             =   720
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Estado de la Persona"
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
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Identificación"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   324
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   408
            Width           =   1212
         End
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   0
         Left            =   1800
         TabIndex        =   17
         ToolTipText     =   "Nuevo"
         Top             =   30
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Nuevo"
         BackColor       =   16761024
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
         Picture         =   "frmPreaEstudiov2.frx":137EA
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   1
         Left            =   960
         TabIndex        =   18
         ToolTipText     =   "Editar"
         Top             =   30
         Visible         =   0   'False
         Width           =   375
         _Version        =   1572864
         _ExtentX        =   656
         _ExtentY        =   582
         _StockProps     =   79
         BackColor       =   16761024
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
         Picture         =   "frmPreaEstudiov2.frx":13E1C
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   2
         Left            =   1320
         TabIndex        =   19
         ToolTipText     =   "Eliminar"
         Top             =   30
         Visible         =   0   'False
         Width           =   375
         _Version        =   1572864
         _ExtentX        =   656
         _ExtentY        =   582
         _StockProps     =   79
         BackColor       =   16761024
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
         Picture         =   "frmPreaEstudiov2.frx":14417
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   3
         Left            =   3000
         TabIndex        =   20
         ToolTipText     =   "Guardar"
         Top             =   30
         Width           =   375
         _Version        =   1572864
         _ExtentX        =   656
         _ExtentY        =   582
         _StockProps     =   79
         BackColor       =   16761024
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
         Picture         =   "frmPreaEstudiov2.frx":149BB
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   4
         Left            =   3360
         TabIndex        =   21
         ToolTipText     =   "Deshacer"
         Top             =   30
         Width           =   375
         _Version        =   1572864
         _ExtentX        =   656
         _ExtentY        =   582
         _StockProps     =   79
         BackColor       =   16761024
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
         Picture         =   "frmPreaEstudiov2.frx":150EC
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   330
         Index           =   5
         Left            =   3840
         TabIndex        =   22
         ToolTipText     =   "Reporte"
         Top             =   30
         Width           =   375
         _Version        =   1572864
         _ExtentX        =   656
         _ExtentY        =   582
         _StockProps     =   79
         BackColor       =   16761024
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
         Picture         =   "frmPreaEstudiov2.frx":157EC
      End
      Begin XtremeSuiteControls.GroupBox gbCredito 
         Height          =   2415
         Left            =   120
         TabIndex        =   32
         Top             =   2640
         Width           =   17775
         _Version        =   1572864
         _ExtentX        =   31353
         _ExtentY        =   4260
         _StockProps     =   79
         Caption         =   "Datos para el Crédito"
         ForeColor       =   8421504
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
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkPolizaVida 
            Height          =   255
            Left            =   10200
            TabIndex        =   33
            Top             =   360
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Póliza de Vida"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkPolizaIncendio 
            Height          =   255
            Left            =   10200
            TabIndex        =   34
            Top             =   720
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Póliza de Incendio"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkPrimerCuota 
            Height          =   615
            Left            =   10200
            TabIndex        =   35
            Top             =   1800
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Primera Cuota"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtLinea 
            Height          =   315
            Left            =   1680
            TabIndex        =   36
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   360
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
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
         Begin XtremeSuiteControls.FlatEdit txtDesLineaCredito 
            Height          =   315
            Left            =   3480
            TabIndex        =   37
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   360
            Width           =   6615
            _Version        =   1572864
            _ExtentX        =   11663
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
         Begin XtremeSuiteControls.FlatEdit txtPolizaVida 
            Height          =   315
            Left            =   12600
            TabIndex        =   38
            Top             =   360
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizaIncendio 
            Height          =   315
            Left            =   12600
            TabIndex        =   39
            Top             =   720
            Width           =   1815
            _Version        =   1572864
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAsignado 
            Height          =   315
            Left            =   8520
            TabIndex        =   40
            Top             =   1200
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuota 
            Height          =   315
            Left            =   15720
            TabIndex        =   41
            Top             =   1560
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtCompromiso 
            Height          =   315
            Left            =   15720
            TabIndex        =   42
            Top             =   1920
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3619
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   315
            Left            =   15720
            TabIndex        =   43
            Top             =   360
            Width           =   2055
            _Version        =   1572864
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTasa 
            Height          =   315
            Left            =   16680
            TabIndex        =   44
            Top             =   1080
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1926
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlazo 
            Height          =   315
            Left            =   16680
            TabIndex        =   45
            Top             =   720
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1926
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlMax 
            Height          =   315
            Left            =   15720
            TabIndex        =   46
            Top             =   720
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1714
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.CheckBox chkPolizaDesempleo 
            Height          =   255
            Left            =   10200
            TabIndex        =   47
            Top             =   1560
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Póliza de Desempleo"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizaDesempleo 
            Height          =   315
            Left            =   12600
            TabIndex        =   48
            Top             =   1560
            Width           =   1815
            _Version        =   1572864
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkPolizaVehiculo 
            Height          =   255
            Left            =   10200
            TabIndex        =   60
            Top             =   1080
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Póliza de Prendas"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   21
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizaPrenda 
            Height          =   315
            Left            =   12600
            TabIndex        =   61
            Top             =   1080
            Width           =   1815
            _Version        =   1572864
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboCPH 
            Height          =   330
            Left            =   5040
            TabIndex        =   64
            Top             =   1560
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
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
         Begin XtremeSuiteControls.FlatEdit txtCRM 
            Height          =   330
            Left            =   5040
            TabIndex        =   66
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1920
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
         Begin XtremeSuiteControls.FlatEdit txtMontoConstruccion 
            Height          =   315
            Left            =   8520
            TabIndex        =   67
            Top             =   1920
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboCantidadFiadores 
            Height          =   330
            Left            =   5040
            TabIndex        =   69
            Top             =   1200
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
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
         Begin XtremeSuiteControls.ComboBox cboFondo 
            Height          =   330
            Left            =   1680
            TabIndex        =   70
            Top             =   1560
            Width           =   2295
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboGarantia 
            Height          =   330
            Left            =   1680
            TabIndex        =   71
            Top             =   1200
            Width           =   2295
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboDestino 
            Height          =   330
            Left            =   1680
            TabIndex        =   72
            Top             =   720
            Width           =   8415
            _Version        =   1572864
            _ExtentX        =   14843
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
         Begin XtremeSuiteControls.ComboBox cboFondoContrato 
            Height          =   330
            Left            =   1680
            TabIndex        =   195
            Top             =   1920
            Width           =   2295
            _Version        =   1572864
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
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contrato"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   63
            Left            =   120
            TabIndex        =   196
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label lblMontoConstruccion 
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Construcción"
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
            Left            =   6480
            TabIndex        =   68
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "No.Op.CRM"
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
            Index           =   36
            Left            =   4080
            TabIndex        =   65
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CPH"
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
            Index           =   7
            Left            =   4080
            TabIndex        =   63
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Fiadores"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   43
            Left            =   4080
            TabIndex        =   62
            Top             =   1230
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Monto"
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
            Index           =   16
            Left            =   14520
            TabIndex        =   58
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label LblPlazo 
            BackStyle       =   0  'Transparent
            Caption         =   "Plazo"
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
            Left            =   14520
            TabIndex        =   57
            Top             =   720
            Width           =   975
         End
         Begin VB.Label LblTasa 
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa"
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
            Left            =   14520
            TabIndex        =   56
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Asignado a la Operación #"
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
            Index           =   39
            Left            =   6480
            TabIndex        =   55
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuota"
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
            Index           =   8
            Left            =   14520
            TabIndex        =   54
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Compromiso"
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
            Index           =   2
            Left            =   14520
            TabIndex        =   53
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Garantía"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   52
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Destino"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Respaldo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   34
            Left            =   120
            TabIndex        =   50
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Línea"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSComctlLib.ImageList ImageListtblGestion 
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
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":15EF3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":1688B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":16E6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":1766F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":17FD3
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":1876C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":1912F
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":1970E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":19F15
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":1A71A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":1AF09
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreaEstudiov2.frx":1B87D
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.PushButton btnAbandonar 
         Height          =   330
         Left            =   7320
         TabIndex        =   323
         ToolTipText     =   "Poner en Estado de Solicitado"
         Top             =   30
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Abandonar"
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   21
         Picture         =   "frmPreaEstudiov2.frx":1C013
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Expediente"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
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
         Height          =   405
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   18255
         _Version        =   1572864
         _ExtentX        =   32200
         _ExtentY        =   706
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   10
      End
      Begin XtremeShortcutBar.ShortcutCaption lblEstado 
         Height          =   405
         Left            =   7080
         TabIndex        =   2
         ToolTipText     =   "Estado del Estudio"
         Top             =   600
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   714
         _StockProps     =   14
         Caption         =   "Estado del Estudio"
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
         VisualTheme     =   9
         Alignment       =   1
      End
   End
End
Attribute VB_Name = "frmPreaEstudiov2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim mTipoExtra As String, mTipoExtraLista As String, vChanged As Boolean

Dim vPaso As Boolean, mFecha As Date

Dim m_curValor_Anterior As Currency ' Variable para comparar en los txt si vario el dato

Dim mFrecuenciaPago As String

Private clsMensajes As New ProGrX_EstudioCrd.clsEstudioMensajes
Private clsEntidad As New ProGrX_EstudioCrd.clsEntidad
Private clsNull As New ProGrX_EstudioCrd.clsNull

Public Item_Seleccionado As XtremeSuiteControls.ListViewItem
Public litem As XtremeSuiteControls.ListViewItem

Private m_ventanaEnModo As eVentanaEnModo
Private vCodExpediente As String
Private m_CambioDatos As Boolean
Private m_CambioCalculo As Boolean
Private m_CambioObservaciones As Boolean
Private m_valorComboExp As String
Private m_Cargando As Boolean
Private m_Paso As Boolean
Private m_CargoSalario As Boolean

Private m_MuestraMensaje As Boolean 'Control de mensajes  para el usuario
Private m_Expediente As String 'Numero de expediente actual consultado.
Private m_expedienteAnterior As String 'Almacena de forma temporal el expediente anterior consultado.
Private m_PreviousTab As Integer 'Mantiene el tab anterior.
Private m_FiadoresRegistrador As Integer ' Almacena la cantidad de fiadores por expedientes
Private m_estadoPreanalisis As String 'Mantiene el estado del preanalisis
Private m_DesplegoMensaje As Boolean 'Controla la despliegue de mensajes de información al usuario
Private m_CargoCombo As Boolean 'Para que no cargue el combo nuevamente

Private m_SoloVerSalarios As Boolean ' Variable que indica si en la lista de salarios solo es para consulta.

Dim m_SalarioDevengadoGrupo As Currency

Dim vPasoCarga As Boolean 'Para que Bancos Click cbo lo ignore

Dim m_NumPagos As Integer
Dim m_Editable As Boolean

Private m_FECHA_CREACION As String 'Para el calulo de colilla

Dim pTabBefore As Integer

'Leer el PortaPapeles
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long


Private m_Formulas As eFormulas

'Rutinas y funciones del forms

Private Sub DespliegaEdad(ByVal Anos As Integer, ByVal Meses As Integer, Optional ByVal Dias As Integer)
Dim tMeses As String, tAnos As String, tDias As String
' Para dar formato a las fechas
On Error GoTo vError
    
    tAnos = ""
    If (Anos = 1) Then
        tAnos = Anos & " Año "
    Else
        If (Anos > 1) Then
            tAnos = Anos & " Años "
        End If
    End If

    tMeses = ""
    If (Meses = 1) Then
        tMeses = Meses & " Mes "
    Else
        If (Meses > 1) Then
            tMeses = Meses & " Meses "
        End If
    End If

    tDias = Empty
    If Dias = 1 Then
        tDias = Dias & " Día"
    Else
        If (Dias > 1) Then
            tDias = Dias & " Días"
        End If
    End If

    txtEdad.Text = tAnos & tMeses & tDias
    

    Exit Sub
vError:
    MsgBox "Ocurrió un error al ejecutar el despliegue de la edad. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub


Public Function fxCalculaEdadAnos(ByVal pFechaNacimiento As String, ByVal ValorReturn As String) As Integer
Dim MesesOri As Integer
Dim FechaMasMeses As Date
Dim FechaActual As Date
Dim FechaNacimiento As Date
Dim fAnios As Integer
Dim fMeses As Integer
Dim Diasint As Integer

fxCalculaEdadAnos = 0

On Error GoTo vError
   
   FechaNacimiento = Format(pFechaNacimiento, "dd/mm/yyyy")
   
   If m_FECHA_CREACION = "-1" Or m_FECHA_CREACION = "" Then
       FechaActual = Format(mFecha, "dd/mm/yyyy")
   Else
       FechaActual = Format(m_FECHA_CREACION, "dd/mm/yyyy")
   End If
   
   MesesOri = DateDiff("m", FechaNacimiento, FechaActual)
   FechaMasMeses = DateAdd("m", MesesOri, FechaNacimiento)
   If Format(FechaMasMeses, "dd/mm/yyyy") > FechaActual Then
       MesesOri = MesesOri - 1
       FechaMasMeses = DateAdd("m", MesesOri, FechaNacimiento)
   End If
   
   fAnios = Int(MesesOri / 12)
   fMeses = MesesOri - (fAnios * 12)
   Diasint = DateDiff("d", FechaMasMeses, FechaActual)
  
    Call DespliegaEdad(fAnios, fMeses, Diasint)
    
    If ValorReturn = "D" Then
    fxCalculaEdadAnos = Diasint
    ElseIf ValorReturn = "M" Then
     fxCalculaEdadAnos = fMeses
    ElseIf ValorReturn = "A" Then
     fxCalculaEdadAnos = fAnios
    End If

    Exit Function
vError:
    MsgBox "Ocurrió un error al calcular la edad. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Function


Public Sub sbEstructuraActualiza(vNivel As Integer, MontoGirar As Boolean)
    '' Procedimiento para actualizar formulas y campos de la forma
    
    vNivel = vNivel + 1
    
    ''Nivel 1 compuesto por: Salario Devengado y Rebajo de Extras
    ''Cualquier cambio en N1 debe recalcular los niveles a partir de N2
    
    ''Nivel 2 compuesto por: Salario Real y Extras fijas
    ''Cualquier cambio en N2 debe recalcular los niveles a partir de N3

    
    'Si es Asociado Aplicar la deducción de la Asociacion
    If gPreAnalisis.Socio = "S" Then
       chkCargaAsociacion.Value = vbChecked
    End If
    
    If vNivel <= 2 Then
    
        Call sbAplicarFormulas(eFormulas.eSalarioReal)
        
    End If
    
    ''Nivel 3 compuesto por: devengado del mes, cargas sociales, porcentaje salario,
    '' deducciones, créditos cancelados, créditos por cobrar
    ''Cualquier cambio en N3 debe recalcular los niveles a partir de N4
    
    If vNivel <= 3 Then
    
        Call sbAplicarFormulas(eFormulas.eDevengadoDelMes)
        Call sbCalcula_Cargas
        Call sbAplicarFormulas(eFormulas.ePorcentajeSobreSalario)
        
    End If
    
    ''Nivel 4 compuesto por: salario liquido, refundiciones, desembolsos
    ''Cualquier cambio en N4 debe recalcular los niveles a partir de N5
    
    If vNivel <= 4 Then
        
        Call sbAplicarFormulas(eFormulas.eSalarioLiquido)
        
    End If
    
    ''Nivel 5 compuesto por: Total liquido, fianzas, nueva couta
    ''Cualquier cambio en N5 debe recalcular los niveles a partir de N6
    
    If vNivel <= 5 Then
        
        Call sbAplicarFormulas(eFormulas.eTotalLiquido)
        
    End If
    
    ''Nivel 6 compuesto por: liquidez sin fianza, liquidez con fianza
    ''este nivel no recalcula ningún siguiente nivel
    
    
    If vNivel <= 6 Then
    
        Call sbAplicarFormulas(eFormulas.eLiquidezConFianza)
        Call sbAplicarFormulas(eFormulas.eLiquidezSinFianzas)
        Call sbAplicarFormulas(eFormulas.eLiquidezPorcConFianza)
        Call sbAplicarFormulas(eFormulas.eLiquidezPorcSinFianzas)
 
        
    End If
    
    '' Nivel exclusivo para recalcular el monto a girar
    
    If MontoGirar = True Then '' Calcula Monto a Girar
        
        Call sbAplicarFormulas(ePolizaSD)
        Call sbAplicarFormulas(eFormulas.eMontoGirar)
        
    End If


End Sub



Public Sub sbCalcularPlazoMaximo()
Dim v_AnosEdad As Double
Dim v_MesesEdad As Double
Dim v_EdadMax As Integer
Dim v_Meses As Integer
Dim v_Dias As Integer

On Error GoTo vError

    v_Meses = 0
    v_Dias = 0
    
    'If InStr(1, TxtExpediente.Text, "-", vbTextCompare) > 0 Then Exit Sub
    
    v_AnosEdad = fxCalculaEdadAnos(dtpFecNac.Value, "A")
    
    v_MesesEdad = fxCalculaEdadAnos(dtpFecNac.Value, "M")
    
    v_AnosEdad = v_AnosEdad + (v_MesesEdad / 12)
    
    If fxSexoItemData(cboSexo.ListIndex) = "M" Then
    
        v_EdadMax = GlobalEdadMaximaPermitidaHombre
        
    ElseIf fxSexoItemData(cboSexo.ListIndex) = "F" Then
    
        v_EdadMax = GlobalEdadMaximaPermitidaMujeres
        
    End If
    
    txtPlMax.Text = CStr(CInt((v_EdadMax - v_AnosEdad) * 12))
    
    'Evaluacion de la Edad (Inicial)
    btnEdad.Visible = False
    txtEdad.Tag = 0
    txtEdad.ToolTipText = ""
    
    If (v_AnosEdad + (Val(txtPlazo.Text) / 12)) >= v_EdadMax Then
        txtEdad.Tag = 0
        txtEdad.Text = txtEdad.Text & ">> La edad supera el límite autorizado <<"
        txtEdad.BackColor = RGB(243, 185, 176)
    Else
        txtEdad.Tag = 1
        txtEdad.Text = txtEdad.Text & ">> La edad es satisfactoria <<"
        txtEdad.BackColor = RGB(187, 215, 247)
    End If
        
    Call sbEdad_Verifica
        
    Exit Sub
vError:
    MsgBox "Ocurrió un error al cacular plazo máximo. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub


Private Sub sbEdad_Verifica()
Dim strSQL As String, rs As New ADODB.Recordset

If txtExpediente.Text = "" Then Exit Sub

On Error GoTo vError

strSQL = "SELECT ISNULL(APL_JUSTIFICACION_EDAD,0) 'EDAD_APLICA', ISNULL(JUSTIFICACION_EDAD,'') 'EDAD_JUSTIFICACION'" _
       & " FROM CRD_PREA_PREANALISIS WHERE COD_PREANALISIS = '" & txtExpediente.Text & "'"
Call OpenRecordSet(rs, strSQL)

Dim Edad As String, Pos As Integer

Pos = InStr(txtEdad.Text, ">>")  ' Encontrar la posición de ">>"

If Pos > 0 Then
    Edad = Trim(Left(txtEdad.Text, Pos - 1)) ' Extraer el texto antes de ">>"
Else
    Edad = Trim(txtEdad.Text) ' Si no encuentra ">>", devuelve el texto completo
End If
 
 
'Evaluacion de la Edad (Revisada)
btnEdad.Visible = False
txtEdad.ToolTipText = ""

If rs!Edad_Aplica = 1 And Len(rs!Edad_Justificacion) > 0 And txtEdad.Tag = 0 Then
    txtEdad.Tag = rs!Edad_Aplica
    txtEdad.ToolTipText = rs!Edad_Justificacion
    btnEdad.Visible = True
End If
        
If txtEdad.Tag = 0 Then
    txtEdad.Text = Edad & ">> La edad supera el límite autorizado <<"
    txtEdad.BackColor = RGB(243, 185, 176)
    btnEdad.Visible = True
Else
    txtEdad.Text = Edad & ">> La edad es satisfactoria <<"
    txtEdad.BackColor = RGB(187, 215, 247)
End If
 
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub sbAplicarFormulas(vFormula As eFormulas)
Dim vPorc As Double
Dim v_Expediente As Boolean
Dim v_Compromiso As Integer
Dim vCodigo As String
Dim vMontoFianzas As Double, vTotalLiquido As Currency, vDevengado As Currency, vDevengadoComp As Currency

On Error GoTo vError

If vPaso Then Exit Sub



vMontoFianzas = 0
vTotalLiquido = 0
vDevengado = 0
vDevengadoComp = 0

vGrid.Col = 1
vGrid.Row = 2
vCodigo = vGrid.Text

If Val(cboCantidadFiadores.Text) = 0 Then
    v_Compromiso = 1
End If

If InStr(1, txtExpediente.Text, "-", vbTextCompare) = 0 Then
    v_Expediente = True
    v_Compromiso = 1
Else
    v_Compromiso = Val(cboCantidadFiadores.Text)
End If

If cboSubExpediente.Text = "Nuevo Expediente" Or cboSubExpediente.Text = "Nuevo SubExpediente" Then Exit Sub


Me.MousePointer = vbHourglass

'Totales

If txtT_Extras.Text = "" Then
    txtT_Extras.Text = Format(0, "Standard")
End If

If txtD_Monto.Text = "" Then
    txtD_Monto.Text = Format(0, "Standard")
End If

If txtD_TotalCargas.Text = "" Then
    txtD_TotalCargas.Text = Format(0, "Standard")
End If

If txtD_TotalColilla.Text = "" Then
    txtD_TotalColilla.Text = Format(0, "Standard")
End If

If txtD_TotalMensual.Text = "" Then
    txtD_TotalMensual.Text = Format(0, "Standard")
End If


If txtC_CuotaCancelaTotal.Text = "" Then
    txtC_CuotaCancelaTotal.Text = Format(0, "Standard")
End If

If txtC_CuotaPorCobrarTotal.Text = "" Then
    txtC_CuotaPorCobrarTotal.Text = Format(0, "Standard")
End If


If txtR_TotalCuotas.Text = "" Then
    txtR_TotalCuotas.Text = Format(0, "Standard")
End If

If txtR_TotalMora.Text = "" Then
    txtR_TotalMora.Text = Format(0, "Standard")
End If

If txtR_TotalRefunde.Text = "" Then
    txtR_TotalRefunde.Text = Format(0, "Standard")
End If


If txtDS_Monto.Text = "" Then
    txtDS_Monto.Text = Format(0, "Standard")
End If

If txtDS_TotalCuota.Text = "" Then
    txtDS_TotalCuota.Text = Format(0, "Standard")
End If

If txtDS_TotalMonto.Text = "" Then
    txtDS_TotalMonto.Text = Format(0, "Standard")
End If

If txtF_TotalCuotas.Text = "" Then
    txtF_TotalCuotas.Text = Format(0, "Standard")
End If

If txtF_TotalSaldos.Text = "" Then
    txtF_TotalSaldos.Text = Format(0, "Standard")
End If


If txtSalarioMinInembargableEstudio.Text = "" Then
    txtSalarioMinInembargableEstudio.Text = Format(0, "Standard")
End If


'----------------------------------------------------------------


If txtPolizaPrenda.Text = "" Then
    txtPolizaPrenda.Text = Format(0, "Standard")
End If

'Revisión e inicialización de campos vacíos
If Val(txtRefundiciones.ToolTipText) = 0 Then
    txtRefundiciones.ToolTipText = Format(0, "Standard")
End If

If Val(txtDesembolsos.ToolTipText) = 0 Then
    txtDesembolsos.ToolTipText = Format(0, "Standard")
End If

If txtSalarioLiquido.Text = "" Then
    txtSalarioLiquido.Text = Format(0, "Standard")
End If

If txtRefundiciones.ToolTipText = "" Then
    txtRefundiciones.ToolTipText = Format(0, "Standard")
End If

If txtDesembolsos.ToolTipText = "" Then
    txtDesembolsos.ToolTipText = Format(0, "Standard")
End If

If txtSalarioDevengado.Text = "" Then
    txtSalarioDevengado.Text = Format(0, "Standard")
End If

If txtRebajoExtras.Text = "" Then
    txtRebajoExtras.Text = Format(0, "Standard")
End If

If Val(txtCompAdicionalBase.Text) = 0 Then
    txtCompAdicionalBase.Text = Format(0, "Standard")
End If

If txtSalarioReal.Text = "" Then
    txtSalarioReal.Text = Format(0, "Standard")
End If

If txtCompAdicionalBase.Text = "" Then
     txtCompAdicionalBase.Text = Format(0, "Standard")
End If

If txtCompAdicionalBase.Text = "" Then
     txtCompAdicionalBase.Text = Format(0, "Standard")
End If

If txtCompAdicional.Text = "" Then
     txtCompAdicional.Text = Format(0, "Standard")
End If

If txtDevengadoMes.Text = "" Then
    txtDevengadoMes.Text = Format(0, "Standard")
End If


If txtIntereses.Text = "" Then
    txtIntereses.Text = Format(0, "Standard")
End If

If txtComisiones.Text = "" Then
    txtComisiones.Text = Format(0, "Standard")
End If


'--------------------------------------
'Bloque de Salarios

If Not IsNumeric(txtS_Devengado.Text) Then
    txtS_Devengado.Text = Format(0, "Standard")
End If

If Not IsNumeric(txtS_Mensual.Text) Then
    txtS_Mensual.Text = Format(0, "Standard")
End If

If Not IsNumeric(txtS_Constancia.Text) Then
    txtS_Constancia.Text = Format(0, "Standard")
End If

If Not IsNumeric(txtS_OrdenPatronal.Text) Then
    txtS_OrdenPatronal.Text = Format(0, "Standard")
End If

If Not IsNumeric(txtS_Privado.Text) Then
    txtS_Privado.Text = Format(0, "Standard")
End If

If Not IsNumeric(txtS_Privado_Porc.Text) Then
    txtS_Privado_Porc.Text = Format(100, "Standard")
End If

If Not IsNumeric(txtS_ComponenteAdicional.Text) Then
    txtS_ComponenteAdicional.Text = Format(0, "Standard")
End If

If Not IsNumeric(txtS_ComponenteAdicionalPorc.Text) Then
    txtS_ComponenteAdicionalPorc.Text = Format(0, "Standard")
End If

'--------------------------------------


If txtCrdTransitoCancelados.Text = "" Then
    txtCrdTransitoCancelados.Text = Format(0, "Standard")
End If
If txtTotal_Cargas_CCSS.Text = "" Then
    txtTotal_Cargas_CCSS.Text = Format(0, "Standard")
End If
If txtDeducciones.Text = "" Then
    txtDeducciones.Text = Format(0, "Standard")
End If
If txtCrdTransitoXCobrar.Text = "" Then
    txtCrdTransitoXCobrar.Text = Format(0, "Standard")
End If
If Val(txtCompromiso.Text) = 0 Then
    txtCompromiso.Text = Format(0, "Standard")
End If
If txtTotalLiquido.Text = "" Then
    txtTotalLiquido.Text = Format(0, "Standard")
End If
If txtCompromiso.Text = "" Then
    txtCompromiso.Text = Format(0, "Standard")
End If

If Val(txtFianzas.Text) = 0 Then
    txtFianzas.Text = Format(0, "Standard")
End If

'----------------------
If txtLiquidezSinFianza.Text = "" Then
    txtLiquidezSinFianza.Text = Format(0, "Standard")
End If

If txtLiquidezConFianza.Text = "" Then
    txtLiquidezConFianza.Text = Format(0, "Standard")
End If

If txtLiquidezPorcSinFianza.Text = "" Then
    txtLiquidezPorcSinFianza.Text = Format(0, "Standard")
End If

If txtLiquidezPorcConFianza.Text = "" Then
    txtLiquidezPorcConFianza.Text = Format(0, "Standard")
End If

'Liquidez Componente

If txtLiquidezSinFianzaComp.Text = "" Then
    txtLiquidezSinFianzaComp.Text = Format(0, "Standard")
End If

If txtLiquidezConFianzaComp.Text = "" Then
    txtLiquidezConFianzaComp.Text = Format(0, "Standard")
End If

If txtLiquidezPorcSinFianzaComp.Text = "" Then
    txtLiquidezPorcSinFianzaComp.Text = Format(0, "Standard")
End If

If txtLiquidezPorcConFianzaComp.Text = "" Then
    txtLiquidezPorcConFianzaComp.Text = Format(0, "Standard")
End If



'----------------------

If txtMonto.Text = "" Then
    txtMonto.Text = Format(0, "Standard")
End If

If txtCuota.Text = "" Then
    txtCuota.Text = Format(0, "Standard")
End If

If txtPSD.Text = "" Then
    txtPSD.Text = Format(0, "Standard")
End If



'Calcula Componente Adicional

txtCompAdicionalBase.Text = txtS_ComponenteAdicional.Text
txtCompAdicional.Text = Format(CCur(txtCompAdicionalBase.Text) * CCur(txtS_ComponenteAdicionalPorc.Text) / 100, "Standard")

'Calcular Deducciones
'Call sbCalcula_Cargas


'Total Liquido de Grupo
    txtTotalLiquido.Text = Format(CDbl(txtSalarioLiquido.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText), "Standard")
    
    vTotalLiquido = CCur(txtTotalLiquido.Text)
    vDevengado = CCur(txtDevengadoMes.Text)
    
    
    If IsNumeric(txtTotalLiquidoGrupo.Text) Then
        If (txtTotalLiquidoGrupo.Text) > vTotalLiquido Then
            vTotalLiquido = CCur(txtTotalLiquidoGrupo.Text)
        End If
    
        If m_SalarioDevengadoGrupo > vDevengado Then
            vDevengado = m_SalarioDevengadoGrupo
        End If
    End If
    
    'Devengado + % Componente Adicional
    vDevengadoComp = CCur(txtCompAdicional.Text)

'Aplicación de Formulas
Select Case vFormula

    Case eFormulas.eSalarioReal
        txtSalarioReal.Text = Format(CCur(txtSalarioDevengado.Text) - CCur(txtRebajoExtras.Text), "Standard")
        
        
        If cboSalario.Text <> "" Then
            If Left(Right(cboSalario.Text, 2), 1) = "e" Then
                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * m_NumPagos), "Standard")
            Else
                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1), "Standard")
            End If
        Else
            txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1), "Standard")
        End If
        
        txtS_Mensual.Text = txtDevengadoMes.Text
        
    Case eFormulas.eDevengadoDelMes
        
'        If cboSalario.Text <> "" Then
'            If Left(Right(cboSalario.Text, 2), 1) = "e" Then
'                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * m_NumPagos) + CDbl(txtCompAdicionalBase.Text), "Standard")
'            Else
'                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1) + CDbl(txtCompAdicionalBase.Text), "Standard")
'            End If
'        Else
'            txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1) + CDbl(txtCompAdicionalBase.Text), "Standard")
'        End If

        If cboSalario.Text <> "" Then
            If Left(Right(cboSalario.Text, 2), 1) = "e" Then
                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * m_NumPagos), "Standard")
            Else
                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1), "Standard")
            End If
        Else
            txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1), "Standard")
        End If
        
        txtS_Mensual.Text = txtDevengadoMes.Text
          
    Case eFormulas.ePorcentajeSobreSalario
    
        If IsNumeric(GlobalPorcLiquidezLibre) Then
            txtPorcSobreSalario.Text = Format((CDbl(txtDevengadoMes.Text) * GlobalPorcLiquidezLibre / 100), "Standard")
        Else
            MsgBox "El parámetro Porcentaje de Liquidez Libre no es un valor numérico.", vbExclamation, gMsgTitulo
        End If
    
    Case eFormulas.eSalarioLiquido
    
        If Right(cboSalario.Text, 3) <> "(g)" Then
            txtSalarioLiquido.Text = Format((CDbl(txtDevengadoMes.Text) + CDbl(txtCrdTransitoCancelados.Text)) - (CDbl(txtTotal_Cargas_CCSS.Text) + CDbl(txtDeducciones.Text) + CDbl(txtCrdTransitoXCobrar.Text)), "Standard")
        End If


       
    Case eFormulas.eTotalLiquido
                
        txtTotalLiquido.Text = Format(CDbl(txtSalarioLiquido.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText), "Standard")
        
        vTotalLiquido = CCur(txtTotalLiquido.Text)
        
        
        
        If IsNumeric(txtTotalLiquidoGrupo.Text) Then
            If v_Compromiso = 1 Then
                txtTotalLiquidoGrupo.Text = Format(vTotalLiquido, "Standard")
            End If
            
            If (txtTotalLiquidoGrupo.Text) > vTotalLiquido Then
                vTotalLiquido = CCur(txtTotalLiquidoGrupo.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText)
            End If
        End If
        
    
    Case eFormulas.eLiquidezSinFianzas
        
        txtLiquidezSinFianza.Text = Format(vTotalLiquido - (CDbl(txtCompromiso.Text) / v_Compromiso), "Standard")
        
        txtLiquidezSinFianzaComp.Text = Format((vTotalLiquido + vDevengadoComp) - (CDbl(txtCompromiso.Text) / v_Compromiso), "Standard")
        
        
    Case eFormulas.eLiquidezPorcSinFianzas
        
        If ((Val(txtDevengadoMes.Text) = 0) And (Val(txtLiquidezSinFianza.Text) > 0 Or Val(txtLiquidezSinFianza.Text) < 0)) Then
            txtLiquidezPorcSinFianza.Text = 0
        Else
           If vDevengado > 0 Then
                txtLiquidezPorcSinFianza.Text = Format((CDbl(txtLiquidezSinFianza.Text) / vDevengado) * 100, "Standard")
           Else
                txtLiquidezPorcSinFianza.Text = 0
           End If
        End If
        
        'Componente
        If ((Val(txtDevengadoMes.Text) = 0) And (Val(txtLiquidezSinFianzaComp.Text) > 0 Or Val(txtLiquidezSinFianzaComp.Text) < 0)) Then
            txtLiquidezPorcSinFianzaComp.Text = 0
        Else
           If vDevengado + vDevengadoComp > 0 Then
                txtLiquidezPorcSinFianzaComp.Text = Format((CDbl(txtLiquidezSinFianzaComp.Text) / (vDevengado + vDevengadoComp)) * 100, "Standard")
           Else
                txtLiquidezPorcSinFianzaComp.Text = 0
           End If
        End If
        
    Case eFormulas.eLiquidezConFianza
        
       vMontoFianzas = CDbl(txtFianzas.Text)
       
       txtLiquidezConFianza.Text = Format(vTotalLiquido - ((CDbl(txtCompromiso.Text) / v_Compromiso) + CDbl(vMontoFianzas)), "Standard")
       txtLiquidezConFianzaComp.Text = Format((vTotalLiquido + vDevengadoComp) - ((CDbl(txtCompromiso.Text) / v_Compromiso) + CDbl(vMontoFianzas)), "Standard")
       
    
    Case eFormulas.eLiquidezPorcConFianza
        'Cálculo del [%] Con Fianzas
        If ((Val(txtLiquidezConFianza.Text) = 0) And vDevengado > 0) Then
            txtLiquidezPorcConFianza.Text = 0
        Else
            
            If vDevengado > 0 Then
                txtLiquidezPorcConFianza.Text = Format((CDbl(txtLiquidezConFianza.Text) / vDevengado) * 100, "Standard")
            End If
        
            If vDevengado + vDevengadoComp > 0 Then
                txtLiquidezPorcConFianzaComp.Text = Format((CDbl(txtLiquidezConFianzaComp.Text) / (vDevengado + vDevengadoComp)) * 100, "Standard")
            End If
        
        End If
        
        Call sbClasificacion_CargaGrid ' Actualiza clasificacion en la ventana.
        
    Case eFormulas.eMontoGirar
       'Calcula monto a girar solo para deudores
        txtMontoGirar.Text = Format(0, "Standard")
        
        If v_Expediente Then
            If Val(txtPSD.Text) = 0 Then
           txtPSD.Text = 0
            End If
            
            If chkPrimerCuota.Value = vbChecked Then
                txtMontoGirar.Text = Format(CDbl(txtMonto.Text) - (0 + CDbl(txtRefundiciones.Text) + CDbl(txtDesembolsos.Text) _
                                + CDbl(txtCuota.Text) + CDbl(txtPSD.Text) + CCur(txtIntereses.Text) + CCur(txtComisiones.Text)), "Standard")
            Else
                txtMontoGirar.Text = Format(CDbl(txtMonto.Text) - (0 + CDbl(txtRefundiciones.Text) + CDbl(txtDesembolsos.Text) _
                                + CDbl(txtPSD.Text) + CCur(txtIntereses.Text) + CCur(txtComisiones.Text)), "Standard")
            End If
        End If
        
    Case eFormulas.ePolizaSD
        'Calcula solo para los expedientes la Póliza saldo deudor
        txtPSD.Text = 0
        
        If txtMonto.Text = "" Then
            txtMonto.Text = Format(0, "Standard")
        End If
        
        If v_Expediente Then
            txtPSD.Text = Format((CDbl(txtMonto.Text) * GlobalPorcPSD) / 100, "Standard")
        End If
        
'******************Todas las formulas****************************************************

    Case eFormulas.eAplicarTodas
        'Aplicada todas la formulas en orden
        txtSalarioReal.Text = Format(CCur(txtSalarioDevengado.Text) - CCur(txtRebajoExtras.Text), "Standard")
        
'        If cboSalario.Text <> "" Then
'            If Left(Right(cboSalario.Text, 2), 1) = "e" Then
'                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * m_NumPagos) + CDbl(txtCompAdicionalBase.Text), "Standard")
'            Else
'                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1) + CDbl(txtCompAdicionalBase.Text), "Standard")
'            End If
'        Else
'            txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1) + CDbl(txtCompAdicionalBase.Text), "Standard")
'        End If
        
        If cboSalario.Text <> "" Then
            If Left(Right(cboSalario.Text, 2), 1) = "e" Then
                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * m_NumPagos), "Standard")
            Else
                txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1), "Standard")
            End If
        Else
            txtDevengadoMes.Text = Format((CDbl(txtSalarioReal.Text) * 1), "Standard")
        End If
        
        txtS_Mensual.Text = txtDevengadoMes.Text
        
        'Fin Calcula devengado del mes
        If Right(cboSalario.Text, 3) <> "(g)" Then
            txtSalarioLiquido.Text = Format((CDbl(txtDevengadoMes.Text) + CDbl(txtCrdTransitoCancelados.Text)) - (CDbl(txtTotal_Cargas_CCSS.Text) + CDbl(txtDeducciones.Text) + CDbl(txtCrdTransitoXCobrar.Text)), "Standard")
        End If
        
        'Calculo del Total Liquido y Salario Devengado (Base)
        txtTotalLiquido.Text = Format(CDbl(txtSalarioLiquido.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText), "Standard")
        
        vTotalLiquido = CCur(txtTotalLiquido.Text)
        vDevengado = CCur(txtDevengadoMes.Text)
        vDevengadoComp = CCur(txtCompAdicional.Text)
        
        If IsNumeric(txtTotalLiquidoGrupo.Text) Then
            If (txtTotalLiquidoGrupo.Text) > vTotalLiquido Then
                vTotalLiquido = CCur(txtTotalLiquidoGrupo.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText)
            End If
        
            If m_SalarioDevengadoGrupo > vDevengado Then
                vDevengado = m_SalarioDevengadoGrupo
            End If
        
        End If
        
        
        txtLiquidezSinFianza.Text = Format(vTotalLiquido - (CDbl(txtCompromiso.Text) / v_Compromiso), "Standard")
        txtLiquidezSinFianzaComp.Text = Format((vTotalLiquido + vDevengadoComp) - (CDbl(txtCompromiso.Text) / v_Compromiso), "Standard")
        
        If Not ((vDevengado = 0) And (Val(txtLiquidezSinFianza.Text) > 0 Or Val(txtLiquidezSinFianza.Text) < 0)) Then
            txtLiquidezPorcSinFianza.Text = Format((CDbl(txtLiquidezSinFianza.Text) / vDevengado) * 100, "Standard")
        Else
            txtLiquidezPorcSinFianza.Text = 0
        End If
        
        If Not ((vDevengado = 0) And (Val(txtLiquidezSinFianzaComp.Text) > 0 Or Val(txtLiquidezSinFianzaComp.Text) < 0)) Then
            txtLiquidezPorcSinFianzaComp.Text = Format((CDbl(txtLiquidezSinFianzaComp.Text) / (vDevengado + vDevengadoComp)) * 100, "Standard")
        Else
            txtLiquidezPorcSinFianza.Text = 0
        End If
        
        
        
        'Calcula Liquidez Con Fianzas
        If Val(txtCompromiso.Text) = 0 Then
            txtCompromiso.Text = Format(0, "Standard")
        End If
        If Val(txtFianzas.Text) = 0 Then
            txtFianzas.Text = Format(0, "Standard")
        End If
        vMontoFianzas = CDbl(txtFianzas.Text)
        
       txtLiquidezConFianza.Text = Format(vTotalLiquido - ((CDbl(txtCompromiso.Text) / v_Compromiso) + CDbl(vMontoFianzas)), "Standard")
       txtLiquidezConFianzaComp.Text = Format((vTotalLiquido + vDevengadoComp) - ((CDbl(txtCompromiso.Text) / v_Compromiso) + CDbl(vMontoFianzas)), "Standard")

        'Calcula Liquiedez [%] Con Fianzas
        If Not ((Val(txtLiquidezConFianza.Text) = 0) And vDevengado > 0) Then
            If CDbl(txtDevengadoMes.Text) > 0 Then
                txtLiquidezPorcConFianza.Text = Format((CDbl(txtLiquidezConFianza.Text) / vDevengado) * 100, "Standard")
            Else
                txtLiquidezPorcConFianza.Text = 0
            End If
        Else
           txtLiquidezPorcConFianza.Text = 0
        End If
        
        'Componenete
        If Not ((Val(txtLiquidezConFianzaComp.Text) = 0) And vDevengado + vDevengadoComp > 0) Then
            If CDbl(txtDevengadoMes.Text) + vDevengadoComp > 0 Then
                txtLiquidezPorcConFianzaComp.Text = Format((CDbl(txtLiquidezConFianzaComp.Text) / (vDevengado + vDevengadoComp)) * 100, "Standard")
            Else
                txtLiquidezPorcConFianzaComp.Text = 0
            End If
        Else
           txtLiquidezPorcConFianzaComp.Text = 0
        End If
        
        
        
        Call sbClasificacion_CargaGrid ' Actualiza clasificacion en la ventana.
        
        'Calcula solo para los expedientes la Póliza saldo deudor
        txtPSD.Text = 0
        If v_Expediente Then
            txtPSD.Text = Format((CDbl(txtMonto.Text) * GlobalPorcPSD) / 100, "Standard")
        End If
        
        'Calcula monto a girar solo para deudores
        txtMontoGirar.Text = Format(0, "Standard")
        If v_Expediente Then
            If Val(txtPSD.Text) = 0 Then
               txtPSD.Text = 0
            End If
            
            If chkPrimerCuota.Value = vbChecked Then
                txtMontoGirar.Text = Format(CDbl(txtMonto.Text) - (0 + CDbl(txtRefundiciones.Text) + CDbl(txtDesembolsos.Text) _
                                + CDbl(txtCuota.Text) + CDbl(txtPSD.Text) + CCur(txtIntereses.Text) + CCur(txtComisiones.Text)), "Standard")
            Else
                txtMontoGirar.Text = Format(CDbl(txtMonto.Text) - (0 + CDbl(txtRefundiciones.Text) + CDbl(txtDesembolsos.Text) _
                                + CDbl(txtPSD.Text) + CCur(txtIntereses.Text) + CCur(txtComisiones.Text)), "Standard")
            End If
        
        End If
        
    End Select


'Calcula Cuota diferencia para la Persona
txtTotalLiquido.Text = Format(CDbl(txtSalarioLiquido.Text) + CDbl(txtRefundiciones.ToolTipText) + CDbl(txtDesembolsos.ToolTipText), "Standard")

txtCuotaDiferencia.Text = Format(CCur(txtCuota.Text) - (CCur(txtRefundiciones.ToolTipText) + CCur(txtDesembolsos.ToolTipText)), "Standard")

If CCur(txtCuotaDiferencia.Text) > 0 Then
   txtCuotaDiferencia.ForeColor = vbRed
   Set imgCuotaDif.Picture = ImageListtblGestion.ListImages.Item(5).Picture
Else
   txtCuotaDiferencia.ForeColor = vbBlack
   Set imgCuotaDif.Picture = ImageListtblGestion.ListImages.Item(4).Picture

End If

txtCuotaDiferencia.Text = Format(Abs(CCur(txtCuotaDiferencia.Text)), "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler("Ocurrió un error al aplicar las formulas. " & "-" & Err.Description), vbExclamation, gMsgTitulo

End Sub

Public Sub SbCalculaCargasCCSSPorDefecto()
Dim vTotalSuma As Double
Dim vCargaAsociacion As Double
Dim vCargaCCSS As Double
Dim vCargaImpSalario As Double

On Error GoTo vError


vCargaAsociacion = 0
vCargaCCSS = 0
vCargaImpSalario = 0

If Val(txtDevengadoMes.Text) = 0 Then
    txtDevengadoMes.Text = 0
End If

If gPreAnalisis.Socio = "S" Then
    If IsNumeric(GlobalPorcAsocSolidarista) Then
        vCargaAsociacion = Format((GlobalPorcAsocSolidarista * txtDevengadoMes.Text) / 100, "Standard")
    End If
End If

If IsNumeric(GlobalPorcCCSS) Then
        vCargaCCSS = Format((GlobalPorcCCSS * txtDevengadoMes.Text) / 100, "Standard")
End If

lblCargaImpSalario.Caption = 0
If Val(txtDevengadoMes.Text) > 0 Then
'    glogon.strSQL = "select dbo.fxCRDPreaCalculaRenta (" & fxFormatearValor(CDbl(txtDevengadoMes.Text), Numerico) & ")"
'    If (execSql(glogon.strSQL, True)) Then
'        If glogon.Recordset(0) & "" = "" Then Exit Sub
'        vCargaImpSalario = Format(glogon.Recordset(0), "Standard")
'    End If
    
    vCargaImpSalario = fxRentaCalculo(CCur(txtDevengadoMes.Text))
    
End If

txtTotal_Cargas_CCSS.Text = Format(vCargaAsociacion + vCargaCCSS + vCargaImpSalario, "Standard")
txtD_TotalCargas.Text = Format(vCargaAsociacion + vCargaCCSS + vCargaImpSalario, "Standard")
    
If Right(cboSalario.Text, 3) = "(g)" Then
    Call sbActCtlConstExternos(False, "(g)")
Else
    Call sbActCtlConstExternos(True, "(d)")
End If

Exit Sub

vError:
    MsgBox "Ocurrió un error al calcular las cargas sociales. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Public Sub sbCalcula_Cargas()

Dim vTotalSuma As Currency
Dim vCargaAsociacion As Currency
Dim vCargaCCSS As Currency
Dim vCargaImpSalario As Currency
Dim vCargaFrap As Currency

Dim FrapPorc As Currency, curSalario As Currency

On Error GoTo vError

If vPaso Then Exit Sub

vCargaAsociacion = 0
vCargaCCSS = 0
vCargaImpSalario = 0
vCargaFrap = 0


If IsNumeric(txtDevengadoMes.Text) Then
    curSalario = CCur(txtDevengadoMes.Text)
Else
    curSalario = 0
End If

If IsNumeric(txtFrapPorc.Text) Then
    FrapPorc = CCur(txtFrapPorc.Text)
Else
    FrapPorc = 0
End If

FrapPorc = FrapPorc + GlobalPorcFRAPFAP


If chkCargaAsociacion.Value = xtpChecked Then
     vCargaAsociacion = Format((curSalario * GlobalPorcAsocSolidarista) / 100, "Standard")
End If
    
If chkCargaFrap.Value = xtpChecked Then
    vCargaFrap = Format((curSalario * FrapPorc) / 100, "Standard")
End If
    
    
If IsNumeric(GlobalPorcCCSS) Then
    vCargaCCSS = Format((curSalario * GlobalPorcCCSS) / 100, "Standard")
End If


lblCargaImpSalario.Caption = 0

If curSalario > 0 Then
    vCargaImpSalario = fxRentaCalculo(curSalario)
'    glogon.strSQL = "select dbo.fxCRDPreaCalculaRenta (" & fxFormatearValor(CDbl(txtDevengadoMes.Text), Numerico) & ")"
'    If (execSql(glogon.strSQL, True)) Then
'        If glogon.Recordset(0) & "" = "" Then Exit Sub
'        vCargaImpSalario = Format(glogon.Recordset(0), "Standard")
'    End If
End If

If chkCargaAsociacion.Value = Checked Then
    vTotalSuma = vCargaAsociacion
End If

If chkCargaFrap.Value = Checked Then
    vTotalSuma = vTotalSuma + vCargaFrap
End If
    
    
lblCargaFrap.Caption = Format(vCargaFrap, "Standard")
lblCargaImpSalario.Caption = Format(vCargaImpSalario, "Standard")
lblCargaCCSS.Caption = Format(vCargaCCSS, "Standard")
lblCargaAsociacion.Caption = Format(vCargaAsociacion, "Standard")

    
vTotalSuma = vTotalSuma + vCargaImpSalario
vTotalSuma = vTotalSuma + vCargaCCSS
'vTotalSuma = vTotalSuma + vCargaAsociacion
'vTotalSuma = vTotalSuma + vCargaFrap

txtTotal_Cargas_CCSS.Text = Format(vTotalSuma, "Standard")
txtD_TotalCargas.Text = Format(vTotalSuma, "Standard")
    
    
If Right(cboSalario.Text, 3) = "(g)" Then
    Call sbActCtlConstExternos(False, "(g)")
Else
    Call sbActCtlConstExternos(True, "(d)")
    
End If
    
Exit Sub

vError:
    MsgBox "Ocurrió un error al calcular las cargas sociales. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub

Public Sub SbCalculaCargasCCSS()
Dim FrapPorc As Currency, curSalario As Currency

On Error GoTo vError

If IsNumeric(txtDevengadoMes.Text) Then
    curSalario = CCur(txtDevengadoMes.Text)
Else
    curSalario = 0
End If

If IsNumeric(txtFrapPorc.Text) Then
    FrapPorc = CCur(txtFrapPorc.Text)
Else
    FrapPorc = 0
End If
    
FrapPorc = FrapPorc + GlobalPorcFRAPFAP
    
If clsMensajes.Estado = "N" Then
    chkCargaAsociacion.Enabled = True
End If

If chkCargaFrap.Tag = "S" Then
    chkCargaFrap.Value = xtpChecked
    chkCargaFrap.Tag = "S"
Else
    chkCargaFrap.Value = 0
    chkCargaFrap.Tag = "N"
End If



If IsNumeric(GlobalPorcAsocSolidarista) Then
    If chkCargaAsociacion.Value = xtpChecked Then
        lblCargaAsociacion.Caption = Format((curSalario * GlobalPorcAsocSolidarista) / 100, "Standard")
    End If
End If

If chkCargaFrap.Value = xtpChecked Then
    lblCargaFrap.Caption = Format((curSalario * FrapPorc) / 100, "Standard")
Else
    lblCargaFrap.Caption = Format(0, "Standard")
End If

If IsNumeric(GlobalPorcCCSS) Then
    lblCargaCCSS.Caption = Format((curSalario * GlobalPorcCCSS) / 100, "Standard")
End If

lblCargaImpSalario.Caption = Format(0, "Standard")

If curSalario > 0 Then
    lblCargaImpSalario.Caption = Format(fxRentaCalculo(curSalario), "Standard")
End If

Exit Sub

vError:
    MsgBox "Ocurrió un error al calcular las cargas sociales. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Public Sub SbBloquearTxtSalario(ByVal pCodigo As String)
On Error GoTo vError

clsEntidad.tablaName = "spCRDPreaTIPO_SALARIO"

'txtSalarioDevengado.Locked = True
txtS_Devengado.Locked = True

txtCompAdicionalBase.Locked = True

'btnDetalle.Item(1).Visible = False
gExtras.Enabled = False

If clsEntidad.fxTraerUno(fxFormatearValor(pCodigo, caracter)) Then
    If glogon.Recordset!MODIFICA_DEVENGADO = 1 Then
       'txtSalarioDevengado.Locked = False
       txtS_Devengado.Locked = False
    End If
    
    If glogon.Recordset!MODIFICA_REBAJO_EXTRAS = 1 Then
       'btnDetalle.Item(1).Visible = True
       gExtras.Enabled = True
    End If
    
    If glogon.Recordset!MODIFICA_EXTRAS_FIJAS = 1 Then
       txtCompAdicionalBase.Locked = False
    End If


    gExtras.Enabled = False
    
    chkS_Constancia.Value = xtpUnchecked
    chkS_OrdenPatronal.Value = xtpUnchecked
    
    chkS_Constancia.Enabled = False
    chkS_OrdenPatronal.Enabled = False
    
    Select Case Mid(glogon.Recordset!Operacion, 1, 1)
        Case "a" 'Salario Minimo
                    
        Case "b" 'Salario Promedio
        
        Case "c" 'Ultimo Salario
        
        Case "d" 'Constancia Salarial
            gSalarios.MaxRows = 0
        
        Case "e" 'Colilla de pago
           gExtras.Enabled = True
           
        Case "f" 'Constante
        
        Case "g" 'Constancia Externos
            gSalarios.MaxRows = 0
            
            chkS_Constancia.Enabled = True
            chkS_OrdenPatronal.Enabled = True
    
    End Select
    
    
    Call chkS_Constancia_Click
    Call chkS_OrdenPatronal_Click


End If


Exit Sub

vError:
  MsgBox "Ocurrió un error desbloquear campos. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Public Sub sbBloquearTab()
Dim i As Integer

On Error GoTo vError

If tcMain.SelectedItem = 0 Then
    
    If Len(txtExpediente.Text) = 0 Then
        For i = 0 To tcMain.ItemCount - 1
            tcMain.Item(i).Enabled = False
        Next i
        tcMain.Item(0).Enabled = True
    
    Else
        For i = 0 To tcMain.ItemCount - 1
            tcMain.Item(i).Enabled = True
        Next i
    End If

End If


Exit Sub

vError:
   MsgBox "Ocurrió un error desbloquear campos. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub

Public Sub sbBloquearControles(ByRef forms As Form, ByVal Expediente As eTipoExpediente)
'Esta rutina se encarga de inicializar los valores de los controles que se encuentra pegados en la patalla
Dim i As Integer, vValor As Boolean

On Error GoTo vError

If (Expediente = SubExpediente) Then
    vValor = False
Else
    vValor = True
End If


gbComite.Enabled = vValor
gbCredito.Enabled = vValor


btnResumen(0).Enabled = vValor
btnResumen(1).Enabled = vValor
btnResumen(2).Enabled = vValor

For i = 9 To tcMain.ItemCount - 1
        tcMain.Item(i).Enabled = vValor
Next i


chkCargaAsociacion.Enabled = True

Exit Sub

vError:
   MsgBox "Ocurrió un error desbloquear campos. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Private Sub sbExpediente_Editable_Bloqueo()
'Bloquea las Opciones en Caso de se o no ser editable

gbCredito.Enabled = m_Editable
gbComite.Enabled = m_Editable


gbSalarios(0).Enabled = m_Editable
gDeducciones.Enabled = m_Editable
gExtras.Enabled = m_Editable

fraDCargas.Enabled = m_Editable
gbEstudioCIC.Enabled = m_Editable

gCuotasCancela.Enabled = m_Editable
gCuotasCobrar.Enabled = m_Editable
btnCreditos(0).Enabled = m_Editable
btnCreditos(1).Enabled = m_Editable


gRefunde.Enabled = m_Editable
btnRefundiciones_Actualiza.Enabled = m_Editable

gDesembolsos.Enabled = m_Editable
btnDesembolso(0).Enabled = m_Editable
'btnDesembolso(1).Enabled = m_Editable

gFianzas.Enabled = m_Editable
btnFianzas_Actualiza.Enabled = m_Editable

gbResumen(0).Enabled = m_Editable
gbResumen(1).Enabled = m_Editable

btnEtiqueta.Enabled = m_Editable
btnAdjunto_Guardar.Enabled = m_Editable
btnAdjunto_Elimina.Enabled = m_Editable

btnHipoteca(0).Enabled = m_Editable
btnHipoteca(1).Enabled = m_Editable
btnHipoteca(4).Enabled = m_Editable

gbPrendario.Enabled = m_Editable
gbPrendaExamenes.Enabled = m_Editable

cmdGuardaObservaciones.Enabled = m_Editable


If lblEstado.Tag = "A" Then

End If

'Bloquea Opciones de Calculo y Seguimiento con Expedientes Nuevos (Sin Asignación de Expediente)
If txtExpediente.Text = "" Then
   gbSalarios(0).Enabled = False
   
   Dim i As Integer
   For i = 1 To tcMain.ItemCount - 1
     tcMain.Item(i).Enabled = False
   Next i
End If

End Sub



Private Sub sbLigarDatos(ByVal rs As ADODB.Recordset)
Dim Codigo As String
Dim Item As String

On Error GoTo vError

With rs
'Información del Tab Datos


m_CargoSalario = True
txtFianzas.Text = Format(0, "Standard")

clsMensajes.TASA_PTS_BONO = 0

txtCedula.Text = !Cedula & ""

DoEvents

Call txtCedula_LostFocus

txtNombre.Text = !Nombre & ""
gPreAnalisis.Expediente = txtExpediente.Text


clsMensajes.Estado = Trim(!Estado & "")
clsMensajes.COD_CAPACIDAD = Trim(!COD_CAPACIDAD & "")
clsMensajes.COD_ENDEUDAMIENTO = Trim(!COD_ENDEUDAMIENTO & "")
clsMensajes.COD_GARANTIA = Trim(!COD_GARANTIA & "")
clsMensajes.COD_HISTORIAL = Trim(!COD_HISTORIAL & "")
clsMensajes.COD_MORA = Trim(!COD_MORA & "")
clsMensajes.CATEGORIA_PERSONA = Trim(!CATEGORIA_PERSONA & "")


gPreAnalisis.Estado = Trim(!Estado & "")
gPreAnalisis.EstadoDesc = Trim(!EstadoDesc)

gPreAnalisis.EstadoV2 = !COD_ESTADO_V2
gPreAnalisis.EstadoV2Desc = !EstadoV2Desc


If !INDICADOR_EDITABLE = 1 Or !INDICADOR_EDITABLE = True Then
    m_Editable = True
Else
    m_Editable = False
End If

gPreAnalisis.Editable = m_Editable

lblEstado.Caption = !EstadoDesc
lblEstado.Tag = !Estado & ""

If gPreAnalisis.EstadoV2Desc <> "" Then
    lblEstado.Caption = gPreAnalisis.EstadoV2Desc
End If

m_estadoPreanalisis = clsMensajes.Estado

Call sbBloquearTab

Call sbSeleccionaSexo(cboSexo, Trim(!sexo & ""))

lblRegistro(0).Caption = "R-US: " & !Usuario
lblRegistro(1).Caption = "R-FE: " & Format(!FECHA_CREACION, "dd-mm-yyyy")

txtEjecutivo.Text = "[" & !Id_Promotor & "] .. " & !PromotorDesc & ""
txtEjecutivo.Tag = !Id_Promotor & ""

txtOficina.Text = !OFICINA & ""
txtOficina.Tag = rs!COD_OFICINA & ""

txtCRM.Text = !NUM_OPORT_CRM & ""

Call sbCboAsignaDato(cboCPH, !COD_FORMULARIO_CPH, True, !COD_FORMULARIO_CPH)

m_FECHA_CREACION = Format(!FECHA_CREACION, "dd-mm-yyyy")

If Trim(!fecha_nacimiento & "") <> "" Then
    dtpFecNac.Value = !fecha_nacimiento
End If
If Trim(!NSUB_EXP & "") <> "" Then
    cboCantidadFiadores.Text = !NSUB_EXP
Else
    cboCantidadFiadores.Text = 0
End If

If Trim(!Cod_Linea & "") <> "" Then
    txtLinea.Text = !Cod_Linea
    Call txtLinea_LostFocus
End If

If Trim(!cod_destino & "") <> "" Then
    Call sbCboAsignaDato(cboDestino, Trim(!DestinoDesc), True, !cod_destino)
End If

If Len(txtDesLineaCredito.Text) > 0 Then
    Call sbSTCargaCboGarantiav2(cboGarantia, txtLinea.Text)
End If

If Trim(!GARANTIA & "") <> "" Then
    Call sbCboAsignaDato(cboGarantia, !GarantiaDesc, True, !GARANTIA)
End If


If Trim(!ComiteDesc & "") <> "" Then
    Call sbCboAsignaDato(cboComite, !ComiteDesc, True, !ID_COMITE)
Else
    cboComite.Text = " "
End If



chkPrimerCuota.Value = !apl_primer_cuota
chkPolizaVida.Value = !APL_POLIZA_VIDA
chkPolizaIncendio.Value = IIf(IsNull(!apl_poliza_incendio), 0, !apl_poliza_incendio)
chkPolizaDesempleo.Value = IIf(IsNull(!APL_POLIZA_DESEMPLEO), 0, !APL_POLIZA_DESEMPLEO)
chkPolizaVehiculo.Value = IIf(IsNull(!APL_POLIZA_VEHICULO), 0, !APL_POLIZA_VEHICULO)



txtPolizaVida.Text = IIf(Val(!MONTO_POLIZA_VIDA & "") = clsNull.NullNumerico, 0, Format(!MONTO_POLIZA_VIDA, "Standard"))
txtPolizaIncendio.Text = IIf(Val(!MONTO_POLIZA_INCENDIO) = clsNull.NullNumerico, 0, Format(!MONTO_POLIZA_INCENDIO, "Standard"))
txtPolizaDesempleo.Text = IIf(Val(!MONTO_POLIZA_DESEMPLEO) = clsNull.NullNumerico, 0, Format(!MONTO_POLIZA_DESEMPLEO, "Standard"))
txtPolizaPrenda.Text = IIf(IsNull(!MONTO_POLIZA_VEHICULO), 0, Format(!MONTO_POLIZA_VEHICULO, "Standard"))
    
   
txtPrendaValor.Text = IIf(IsNull(!MONTO_VALOR_VEHICULO), 0, Format(!MONTO_VALOR_VEHICULO, "Standard"))
txtMontoConstruccion.Text = IIf(IsNull(!MONTO_VALOR_VEHICULO), 0, Format(!MONTO_VALOR_VEHICULO, "Standard"))


txtMonto.Text = IIf(Val(!Monto & "") = clsNull.NullNumerico, 0, Format(!Monto, "Standard"))
txtPlazo.Text = IIf(Val(!Plazo & "") = clsNull.NullNumerico, 0, !Plazo)
txtTasa.Text = IIf(Val(!TASA & "") = clsNull.NullNumerico, 0, Format(!TASA, "Standard"))

txtCumplimientoNotas.Text = !CUMPLIMIENTO_NOTAS & ""

If !TASA_PTS_BONO > 0 Then
    txtTasa.ToolTipText = "Bono por Membresia de " & Format(!TASA_PTS_BONO, "Standard")
Else
    txtTasa.ToolTipText = Empty
End If

txtCuota.Text = IIf(Val(!Cuota & "") = clsNull.NullNumerico, 0, Format(!Cuota, "Standard"))
txtCompromiso.Text = IIf(Val(!COMPROMISO & "") = clsNull.NullNumerico, 0, Format(!COMPROMISO, "Standard"))
txtAsignado.Text = IIf(Val(Trim(!ID_SOLICITUD & "")) = clsNull.NullNumerico, 0, !ID_SOLICITUD)

Call sbCalcularPlazoMaximo

'Información del Tab Calculo

'Call sbSeleccionarItemCombo(cboSalario, Trim(!tipo_salario & ""))

If Trim(!tipo_salario & "") <> "" Then
    Call sbCboAsignaDato(cboSalario, Trim(!DescTipoSalario), True, rs!tipo_salario)
    
    txtTipoSalario.Text = !DescTipoSalario
End If

Codigo = fxDeCodificaPrimaryKey(cboSalario.Text, 1, "-")

Call SbBloquearTxtSalario(Trim(Codigo))

If !FECHA_CORTE_COLIILA & "" <> "" Then
    dtpCorte.Value = !FECHA_CORTE_COLIILA
    
    txtColillaCorte.Text = Format(!FECHA_CORTE_COLIILA, "dd-mm-yyyy")
    
End If

Item = Left(Right(cboSalario.Text, 2), 1)
    
    
'----------------------------------------------------
'Componente Adicional
Call sbCboAsignaDato(cboS_ComponenteAdicional, rs!compAdicionalDesc, True, rs!ID_COMPONENTE_AD)

txtS_ComponenteAdicional.Text = IIf(IsNull(!EXTRAS_FIJAS), 0, Format(!EXTRAS_FIJAS, "Standard"))
txtS_ComponenteAdicionalPorc.Text = IIf(IsNull(!PORCENTAJE_COMPONENTE_AD), 0, Format(!PORCENTAJE_COMPONENTE_AD, "Standard"))

txtCompAdicionalBase.Text = IIf(IsNull(!EXTRAS_FIJAS), 0, Format(!EXTRAS_FIJAS, "Standard"))
txtCompAdicional.Text = Format(CCur(txtCompAdicionalBase.Text) * CCur(txtS_ComponenteAdicionalPorc.Text) / 100, "Standard")



'----------------------------------------------------
'Salarios Nuevos

txtS_Devengado.Text = IIf(IsNull(!SALARIO_DEVENGADO_COLILLA), 0, Format(!SALARIO_DEVENGADO_COLILLA, "Standard"))
txtS_Mensual.Text = IIf(IsNull(!DEVENGADO_MES), 0, Format(!DEVENGADO_MES, "Standard"))

txtS_Constancia.Text = IIf(IsNull(!SALARIO_CONSTANCIA), 0, Format(!SALARIO_CONSTANCIA, "Standard"))
txtS_OrdenPatronal.Text = IIf(IsNull(!SALARIO_ORDEN_PATRONAL), 0, Format(!SALARIO_ORDEN_PATRONAL, "Standard"))
txtS_Privado.Text = IIf(IsNull(!MONTO_ACT_PRIVADAS), 0, Format(!MONTO_ACT_PRIVADAS, "Standard"))

chkS_Constancia.Value = IIf(CCur(txtS_Constancia.Text) > 0, xtpChecked, xtpUnchecked)
chkS_OrdenPatronal.Value = IIf(CCur(txtS_OrdenPatronal.Text) > 0, xtpChecked, xtpUnchecked)

Call chkS_OrdenPatronal_Click
Call chkS_Constancia_Click

txtSalarioDevengado.Text = IIf(Val(!SALARIO_DEVENGADO_COLILLA & "") = clsNull.NullNumerico, 0, Format(!SALARIO_DEVENGADO_COLILLA, "Standard"))

txtSalarioDevengado.ToolTipText = IIf(Val(!SALARIO_DEVENGADO_GRUPO & "") = clsNull.NullNumerico, 0, Format(!SALARIO_DEVENGADO_GRUPO, "Standard"))
m_SalarioDevengadoGrupo = IIf(Val(!SALARIO_DEVENGADO_GRUPO & "") = clsNull.NullNumerico, 0, Format(!SALARIO_DEVENGADO_GRUPO, "Standard"))

txtRebajoExtras.Text = IIf(Val(!REBAJO_EXTRAS & "") = clsNull.NullNumerico, 0, Format(!REBAJO_EXTRAS, "Standard"))

txtSalarioReal.Text = IIf(Val(!SALARIO_REAL & "") = clsNull.NullNumerico, 0, Format(!SALARIO_REAL, "Standard"))

txtDevengadoMes.Text = IIf(Val(!DEVENGADO_MES & "") = clsNull.NullNumerico, 0, Format(!DEVENGADO_MES, "Standard"))



'Totales
txtT_Extras.Text = IIf(Val(!REBAJO_EXTRAS & "") = clsNull.NullNumerico, 0, Format(!REBAJO_EXTRAS, "Standard"))


'----------------------------------------------------
'Cargas Sociales e Impositivas

chkCargaAsociacion.Tag = "N"
clsMensajes.CARGA_ASOCIACION = 0

If !CARGA_ASOCIACION & "" <> "" Then
    clsMensajes.CARGA_ASOCIACION = !CARGA_ASOCIACION
    If clsMensajes.CARGA_ASOCIACION > 0 Then
        chkCargaAsociacion.Tag = "S"
        chkCargaAsociacion.Value = xtpChecked
    End If
End If
If chkCargaAsociacion.Tag = "N" Then
    chkCargaAsociacion.Value = xtpUnchecked
Else
    chkCargaAsociacion.Value = xtpChecked
End If

chkCargaFrap.Tag = "N"
If !CARGA_FRAP & "" <> "" Then
    clsMensajes.CARGA_FRAP = !CARGA_FRAP
    If clsMensajes.CARGA_FRAP > 0 Then
        chkCargaFrap.Tag = "S"
        chkCargaFrap.Value = xtpChecked
    End If
End If

lblCargaCCSS.Caption = IIf(Val(!CARGA_CCSS & "") = clsNull.NullNumerico, 0, Format(!CARGA_CCSS, "Standard"))
lblCargaImpSalario.Caption = IIf(Val(!CARGA_IMPUESTO_SALARIO & "") = clsNull.NullNumerico, 0, Format(!CARGA_IMPUESTO_SALARIO, "Standard"))
lblCargaAsociacion.Caption = IIf(Val(!CARGA_ASOCIACION & "") = clsNull.NullNumerico, 0, Format(!CARGA_ASOCIACION, "Standard"))
lblCargaFrap.Caption = IIf(Val(!CARGA_FRAP & "") = clsNull.NullNumerico, 0, Format(!CARGA_FRAP, "Standard"))

txtFrapPorc.Text = IIf(IsNull(!PTS_EXTRA_FRAP), 0, Format(!PTS_EXTRA_FRAP, "Standard"))

txtTotal_Cargas_CCSS.Text = IIf(Val(!TOTAL_CARGA_CCSS & "") = clsNull.NullNumerico, 0, Format(!TOTAL_CARGA_CCSS, "Standard"))
txtD_TotalCargas.Text = IIf(Val(!TOTAL_CARGA_CCSS & "") = clsNull.NullNumerico, 0, Format(!TOTAL_CARGA_CCSS, "Standard"))


'----------------------------------------------------
'Resumen:

txtPorcSobreSalario.Text = IIf(Val(!PORCENTAJE_LIBRE & "") = clsNull.NullNumerico, 0, Format(!PORCENTAJE_LIBRE, "Standard"))

txtDeducciones.Text = IIf(Val(!DEDUCCIONES & "") = clsNull.NullNumerico, 0, Format(!DEDUCCIONES, "Standard"))

txtCrdTransitoCancelados.Text = IIf(Val(!CRD_TRANSITO_CANCELADOS & "") = clsNull.NullNumerico, 0, Format(!CRD_TRANSITO_CANCELADOS, "Standard"))

txtCrdTransitoXCobrar.Text = IIf(Val(!CRD_TRANSITO_XCOBRAR & "") = clsNull.NullNumerico, 0, Format(!CRD_TRANSITO_XCOBRAR, "Standard"))

txtSalarioLiquido.Text = IIf(Val(!SALARIO_LIQUIDO & "") = clsNull.NullNumerico, 0, Format(!SALARIO_LIQUIDO, "Standard"))

txtRefundiciones.Text = IIf(Val(!REFUNDICIONES & "") = clsNull.NullNumerico, 0, Format(!REFUNDICIONES, "Standard"))

If Val(txtRefundiciones.Text) = 0 Then
    txtRefundiciones.ToolTipText = Format(0, "Standard")

Else
    txtRefundiciones.ToolTipText = IIf(Val(!REFUNDICIONES_CUOTA & "") = Format(clsNull.NullNumerico, "Standard"), 0, Format(!REFUNDICIONES_CUOTA, "Standard"))
End If

txtDesembolsos.ToolTipText = Format(0, "Standard")
txtDesembolsos.Text = IIf(Val(!DESEMBOLSOS & "") = clsNull.NullNumerico, 0, Format(!DESEMBOLSOS, "Standard"))
If Val(txtDesembolsos.Text) = 0 Then

    txtSalarioLiquido.Text = IIf(Val(!SALARIO_LIQUIDO & "") = clsNull.NullNumerico, 0, Format(!SALARIO_LIQUIDO, "Standard"))
Else
    txtDesembolsos.ToolTipText = IIf(Val(!DESEMBOLSOS_CUOTA & "") = Format(clsNull.NullNumerico, "Standard"), 0, Format(!DESEMBOLSOS_CUOTA, "Standard"))
End If


txtTotalLiquido.Text = IIf(Val(!LIQUIDO_TOTAL & "") = clsNull.NullNumerico, 0, Format(!LIQUIDO_TOTAL, "Standard"))
txtTotalLiquidoGrupo.Text = IIf(Val(!LIQUIDO_TOTAL_GRUPO & "") = clsNull.NullNumerico, 0, Format(!LIQUIDO_TOTAL_GRUPO, "Standard"))


txtFianzas.Text = IIf(Val(!FIANZAS & "") = clsNull.NullNumerico, 0, Format(!FIANZAS, "Standard"))


'Resumen Nuevos
txtSalarioMinInembargableEstudio.Text = IIf(IsNull(!SALARIO_USURA), 0, Format(!SALARIO_USURA, "Standard"))
txtSalarioNormativaEstudio.Text = IIf(IsNull(!SALARIO_NORMATIVA), 0, Format(!SALARIO_NORMATIVA, "Standard"))


txtLiquidezSinFianza.Text = IIf(IsNull(!LIQUIDEZ_SIMPLE), 0, Format(!LIQUIDEZ_SIMPLE, "Standard"))
txtLiquidezConFianza.Text = IIf(IsNull(!LIQUIDEZ_CFIANZAS), 0, Format(!LIQUIDEZ_CFIANZAS, "Standard"))

txtLiquidezPorcSinFianza.Text = IIf(IsNull(!PORC_LIQ_SIN_FIANZA), 0, Format(!PORC_LIQ_SIN_FIANZA, "Standard"))
txtLiquidezPorcConFianza.Text = IIf(IsNull(!PORC_LIQ_CON_FIANZA), 0, Format(!PORC_LIQ_CON_FIANZA, "Standard"))

txtLiquidezSinFianzaComp.Text = IIf(IsNull(!LIQUIDEZ_SFIANZAS_CA), 0, Format(!LIQUIDEZ_SFIANZAS_CA, "Standard"))
txtLiquidezConFianzaComp.Text = IIf(IsNull(!LIQUIDEZ_CFIANZAS_CA), 0, Format(!LIQUIDEZ_CFIANZAS_CA, "Standard"))

txtLiquidezPorcSinFianzaComp.Text = IIf(IsNull(!PORC_LIQ_SIN_FIANZA_CA), 0, Format(!PORC_LIQ_SIN_FIANZA_CA, "Standard"))
txtLiquidezPorcConFianzaComp.Text = IIf(IsNull(!PORC_LIQ_CON_FIANZA_CA), 0, Format(!PORC_LIQ_CON_FIANZA_CA, "Standard"))


txtComisiones.Text = IIf(IsNull(!MONTO_COMISION), 0, Format(!MONTO_COMISION, "Standard"))
txtIntereses.Text = IIf(IsNull(!Monto_Interes), 0, Format(!Monto_Interes, "Standard"))

txtDiasIntereses.Text = IIf(IsNull(!DIAS_INTERES_GASTOS_OP), 0, CStr(!DIAS_INTERES_GASTOS_OP))

'Totales Tabs Nuevos
txtT_Extras.Text = IIf(IsNull(!REBAJO_EXTRAS), 0, Format(!REBAJO_EXTRAS, "Standard"))

txtD_TotalColilla.Text = IIf(IsNull(!T_DEDUC_CUOTA_COLILLA), 0, Format(!T_DEDUC_CUOTA_COLILLA, "Standard"))
txtD_TotalMensual.Text = IIf(IsNull(!T_DEDUC_CUOTA_MENSUAL), 0, Format(!T_DEDUC_CUOTA_MENSUAL, "Standard"))

txtCIC_Puntaje.Text = !PUNTOS_CIC_DEUDOR & ""
txtCIC_NivelHistorico = !NIVEL_COMPORTAMIENTO_HIST & ""


txtC_CuotaCancelaTotal.Text = IIf(IsNull(!CRD_TRANSITO_CANCELADOS), 0, Format(!CRD_TRANSITO_CANCELADOS, "Standard"))
txtC_CuotaPorCobrarTotal.Text = IIf(IsNull(!CRD_TRANSITO_XCOBRAR), 0, Format(!CRD_TRANSITO_XCOBRAR, "Standard"))

txtR_TotalRefunde.Text = IIf(IsNull(!REFUNDICIONES), 0, Format(!REFUNDICIONES, "Standard"))
txtR_TotalCuotas.Text = IIf(IsNull(!REFUNDICIONES_CUOTA), 0, Format(!REFUNDICIONES_CUOTA, "Standard"))
txtR_TotalMora.Text = IIf(IsNull(!REFUNDICIONES_MORA), 0, Format(!REFUNDICIONES_MORA, "Standard"))

txtDS_TotalMonto.Text = IIf(IsNull(!DESEMBOLSOS), 0, Format(!DESEMBOLSOS, "Standard"))
txtDS_TotalCuota.Text = IIf(IsNull(!DESEMBOLSOS_CUOTA), 0, Format(!DESEMBOLSOS_CUOTA, "Standard"))

txtF_TotalSaldos.Text = IIf(IsNull(!FIANZAS), 0, Format(!FIANZAS, "Standard"))
txtF_TotalCuotas.Text = IIf(IsNull(!T_FIANZA_CUOTA), 0, Format(!T_FIANZA_CUOTA, "Standard"))


txtCFIA_Avaluo.Text = IIf(IsNull(!MONTO_AVALUO_CFIA), 0, Format(!MONTO_AVALUO_CFIA, "Standard"))

'Obtienen las obsevaciones
'& "" esto me valida si tiene nulos
vObservacion(0) = Trim(!OBSERVACION_ANALISTA & "") 'Observaciones de Analisis de crédito
vObservacion(1) = Trim(!OBSERVACION_COMITE & "") 'Observaciones de Resolución del comité
vObservacion(2) = Trim(!OBSERVACION_JD & "") 'Observaciones de Junta directiva

txtCumplimientoNotas.Text = Trim(!CUMPLIMIENTO_NOTAS & "")

If !Estado = "R" Or !Estado = "P" Then
    'Call SbCalculaCargasCCSS
    Call sbCalcula_Cargas
    Call sbCalculaPolizaDeVida
    Call sbCalculaPolizaDeIncendio
    Call sbCalculaPolizaDesempleo
End If


End With

'Call sbToolBar(Me.tlb, "edicion")
Call sbAccionVentana(ModificarRegistro)


Call sbTraerNumFiadores

 m_CambioDatos = False



Exit Sub

vError:
    MsgBox "Ocurrió un error al mostrar la información consultada. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Public Sub sbAsignaTipoSalario(pDato As String)
On Error GoTo vError

        cboSalario.Text = Trim(pDato)

Exit Sub

vError:
  cboSalario.AddItem pDato
  cboSalario.Text = pDato
  
End Sub


Private Function fxExistenFiadores() As Boolean
Dim m_Valor As String
Dim Indicador As String
Dim sql As String
Dim ExpedientePadre As String
Dim m_codGarantia As String

On Error GoTo vError


m_FiadoresRegistrador = 0
fxExistenFiadores = True

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    MsgBox "Debe selecionar un expediente maestro.", vbInformation, gMsgTitulo
    fxExistenFiadores = False
    Exit Function
Else
    ExpedientePadre = txtExpediente.Text
End If

sql = "Select count(*) as NumFiadores from CRD_PREA_PREANALISIS where COD_PREANALISIS_REF = " & fxFormatearValor(ExpedientePadre, caracter)

If clsEntidad.fxEjecutaSQL(sql) Then
    m_FiadoresRegistrador = glogon.Recordset!NumFiadores
End If
m_codGarantia = cboGarantia.ItemData(cboGarantia.ListIndex)

If (m_codGarantia = "F") And Val(cboCantidadFiadores.Text) > m_FiadoresRegistrador Then
    MsgBox "No se han registrado todos fiadores indicados en el expediente maestro.", vbInformation, gMsgTitulo
    fxExistenFiadores = False
    
End If


    Exit Function
vError:
    MsgBox "Ocurrió un error traer los sub expediente registrados. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Function


Private Function fxValidaNumFiadoresRegistrados(Optional MuestreMensaje As Boolean) As Boolean
Dim m_Valor As String
Dim Indicador As String
Dim sql As String
Dim ExpedientePadre As String

On Error GoTo vError
fxValidaNumFiadoresRegistrados = True
m_FiadoresRegistrador = 0

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    ExpedientePadre = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
Else
    ExpedientePadre = txtExpediente.Text
End If

sql = "Select count(*) as NumFiadores from CRD_PREA_PREANALISIS where COD_PREANALISIS_REF = " & fxFormatearValor(ExpedientePadre, caracter)

If clsEntidad.fxEjecutaSQL(sql) Then
    m_FiadoresRegistrador = glogon.Recordset!NumFiadores
    clsMensajes.NSUB_EXP = m_FiadoresRegistrador
End If
If (m_FiadoresRegistrador + 1) > Val(cboCantidadFiadores.Text) Then
    If MuestreMensaje Then
        MsgBox "Para agregar un sub expediente debe aumentar la cantidad de fiadores/co-deudores en el expediente maestro.", vbInformation, gMsgTitulo
    End If
    fxValidaNumFiadoresRegistrados = False
End If

    Exit Function
vError:
    MsgBox "Ocurrió un error traer los sub expediente registrados. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    fxValidaNumFiadoresRegistrados = False

End Function

Private Sub sbTraerNumFiadores()
Dim m_Valor As String
Dim Indicador As String
Dim sql As String
Dim ExpedientePadre As String

On Error GoTo vError


If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    ExpedientePadre = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
Else
    ExpedientePadre = txtExpediente.Text
End If

sql = "Select NSUB_EXP from CRD_PREA_PREANALISIS where   COD_PREANALISIS = " & fxFormatearValor(ExpedientePadre, caracter)

If clsEntidad.fxEjecutaSQL(sql) Then
cboCantidadFiadores.Text = glogon.Recordset!NSUB_EXP
End If

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto digitado. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    

End Sub
    
Private Sub sbTraerDatosExpediente()
    
On Error GoTo vError

Dim strSQL As String, rs As New ADODB.Recordset
Dim VcboCantidadFiadores As Integer
Dim m_Valor As String

Dim Indicador As String

Me.MousePointer = vbDefault

If Len(txtExpediente.Text) = 0 Then Exit Sub

    'carga combo de salarios por si se agregó algún tipo de salario no activo
    cboSalario.Clear
    Call sbLlenarComboTodosV2(cboSalario, "spCRDPreaTIPO_SALARIO", "TIPO_SALARIO", "DescTipoSalario")

clsEntidad.tablaName = "spCRDPreaPREANALISIS"
m_expedienteAnterior = txtExpediente.Text

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    m_Valor = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
    Indicador = "S"
    btnComiteCambio.Enabled = False

Else
    m_Valor = txtExpediente.Text
    Indicador = "E"
    btnComiteCambio.Enabled = True

End If
        
If clsEntidad.fxTraerUno(fxFormatearValor(txtExpediente.Text, caracter)) Then

    Set rs = glogon.Recordset
    m_Cargando = True
    txtExpediente.Locked = False
    If m_CargoCombo = False Then
        cboSubExpediente.Clear
        cboSubExpediente.AddItem m_Valor
        Call sbLlenarComboFiltrado(cboSubExpediente, "spCRDPreaPREANALISIS", "COD_PREANALISIS", "COD_PREANALISIS", "SubExpediente", "", fxFormatearValor(m_Valor, caracter))
        cboSubExpediente.AddItem "Nuevo Expediente"
        cboSubExpediente.ItemData(cboSubExpediente.NewIndex) = -1
        VcboCantidadFiadores = IIf(Val(cboCantidadFiadores.Text) = 0, 1, Val(cboCantidadFiadores.Text))
        
'        (cboGarantia.ItemData(cbogarantia.ListIndex) = "F" And
        If (m_FiadoresRegistrador + 1) <= VcboCantidadFiadores Then
        
            cboSubExpediente.AddItem "Nuevo SubExpediente"
            cboSubExpediente.ItemData(cboSubExpediente.NewIndex) = -2
            
        End If

    End If
    
    Call sbLigarDatos(rs)

    Call sbExtras_Load
    Call sbSalarios_Load
    Call sbIncapacidades_Load

    If Indicador = "S" Then
        Call sbBloquearControles(Me, SubExpediente)
    Else
       Call sbBloquearControles(Me, Expediente)
    End If
    
    If lblEstado.Tag <> "R" Then
      btnNotificacion.Visible = True
      btnComiteCambio.Enabled = False

    End If
    
    If lblEstado.Tag <> "A" And txtAsignado.Text = "0" Then
        btnGestion(2).Enabled = False 'Gestion Estudio
        btnGestion(3).Enabled = True 'Gestion Solicitud de Credito
    End If
    
Else
'    Call sbToolBar(Me.tlb, "edicion")
    Call sbInicializaComboExpediente
'    Call tlb_ButtonClick(tlb.Buttons("nuevo"))
    txtExpediente.Locked = False
End If


m_Cargando = False

Call dtpCorte_Change

DoEvents
Call sbSeleccionarItemComboExp(cboSubExpediente, m_expedienteAnterior)


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto digitado. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Private Sub sbTraerMaxExpediente()

Dim vpadre As String
Dim vRecordset As New ADODB.Recordset

On Error GoTo vError

Set vRecordset = Nothing

If m_valorComboExp = "Nuevo SubExpediente" Then
        If InStr(1, txtExpediente.Text, "-", vbTextCompare) <> 0 Then
            vpadre = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
        Else
            vpadre = txtExpediente.Text
        End If
         clsEntidad.tablaName = "spCRDPreaMaxSubExpediente"
        If clsEntidad.fxTraerUno(fxFormatearValor(vpadre, caracter)) Then
        Set vRecordset = glogon.Recordset
            If Trim(vRecordset(0) & "") <> "" Then
                txtExpediente.Text = Trim(vRecordset(0) & "")
                m_Expediente = Trim(vRecordset(0) & "")
            End If
         
        End If
ElseIf InStr(1, txtExpediente.Text, "-", vbTextCompare) = 0 Then
    clsEntidad.tablaName = "spCRDPreaPARAMETROS"
    m_Expediente = clsEntidad.fxTraerValor("VALOR", "'11'")
    If m_Expediente <> -1 Then
        txtExpediente.Text = m_Expediente
    End If
    If m_valorComboExp = "Nuevo SubExpediente" Then
         clsEntidad.tablaName = "spCRDPreaMaxSubExpediente"
        If clsEntidad.fxTraerUno(fxFormatearValor(m_Expediente, caracter)) Then
        Set vRecordset = glogon.Recordset
            If Trim(vRecordset(0) & "") <> "" Then
                txtExpediente.Text = Trim(vRecordset(0) & "")
                m_Expediente = Trim(vRecordset(0) & "")
            End If
         
        End If
    End If
Else
    vpadre = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
    clsEntidad.tablaName = "spCRDPreaMaxSubExpediente"
    If clsEntidad.fxTraerUno(fxFormatearValor(vpadre, caracter)) Then
    Set vRecordset = glogon.Recordset
        If Trim(vRecordset(0) & "") <> "" Then
            txtExpediente.Text = Trim(vRecordset(0) & "")
            m_Expediente = Trim(vRecordset(0) & "")
        End If
     
    End If
    
End If


    Exit Sub
vError:
    MsgBox "Ocurrió un error consultar número de experiente registrado. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Private Sub sbAccionVentana(ByVal Tipo As eVentanaEnModo)

m_CambioDatos = False
m_CambioCalculo = False
m_CambioObservaciones = False

Select Case True
  Case Tipo = eVentanaEnModo.NuevoRegistro
        m_ventanaEnModo = eVentanaEnModo.NuevoRegistro
        tcMain.Item(0).Selected = True
        m_FiadoresRegistrador = 0
  
  Case Tipo = ModificarRegistro
        m_ventanaEnModo = eVentanaEnModo.ModificarRegistro

End Select

Call sbExpediente_Editable_Bloqueo

End Sub
    
Private Sub sbCargarCombos()
'Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

'Call sbSTCargaCboGarantia(cboGarantia, FormatearValor("ADB2", Caracter))
'Call sbLlenarComboTodos(cboDestino, "spCRDPreaDestinos", "cod_destino", "DescDestino", "Seleccione un destino")


dtpFecNac.Value = dtpCorte.Value

cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("TS")
cboTipoDocumento.AddItem fxTipoDocumento("ND")
cboTipoDocumento.Text = fxTipoDocumento("TE")

vPaso = True
strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
vPaso = False

strSQL = "select COD_DIVISA as 'IdX', DESCRIPCION as 'ItmX'   From vSys_Divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)


cboBanco.Clear
strSQL = "exec spCrd_SGT_Bancos_Desembolso '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  MsgBox "No existen Bancos [Creados/Asignados], verifique en Tesoreria...", vbCritical

Else
 Do While Not rs.EOF
    cboBanco.AddItem IIf(IsNull(rs!Descripcion), "SIN DESCRIPCION", rs!Descripcion)
    cboBanco.ItemData(cboBanco.ListCount - 1) = CStr(rs!Id_Banco)
   
   rs.MoveNext
 Loop
 rs.MoveFirst
 Call sbCboAsignaDato(cboBanco, IIf(IsNull(rs!Descripcion), "SIN DESCRIPCION", rs!Descripcion), True, rs!Id_Banco)
End If
rs.Close

'Etiquetas de Seguimiento v2
strSQL = "SELECT COD_ETIQUETA as 'IdX',DESCRIPCION as 'ItmX'" _
       & " FROM CRD_PREA_V2_ETIQUETAS WHERE SISTEMA = 0 AND MANEJO_ERRORES = 0"
Call sbCbo_Llena_New(cboEtiquetas, strSQL, False, True)

'Carga la Lista de Extras para Combos en Grids
strSQL = "select rtrim(cod_Extras) + ' - ' + rtrim(descripcion) as 'TipoExtra'" _
       & " from CRD_PREA_TIPOS_EXTRAS order by cod_Extras"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And mTipoExtra = "" Then mTipoExtra = rs!TipoExtra

Do While Not rs.EOF
  If Len(mTipoExtraLista) = 0 Then
    mTipoExtraLista = Chr$(9) & rs!TipoExtra
  Else
    mTipoExtraLista = mTipoExtraLista & Chr$(9) & rs!TipoExtra
  End If
  rs.MoveNext
Loop
rs.Close


Call sbSeleccionaSexo(cboSexo, "F")

vPaso = True
strSQL = "select ID_DEDUCCION AS 'IdX', rtrim(DESCRIPCION) as 'ItmX'" _
       & "  From CRD_PREA_V2_DEDUCCIONES_CONFIG" _
       & " Where AUTOMATICA = 0 ORDER BY PRIORIDAD"
Call sbCbo_Llena_New(cboDeduccion, strSQL, False, True)


strSQL = "SELECT COD_PARAMETRO as 'IdX',RTRIM(DESCRIPCION) + ' [ ' + VALOR + ' % ]' as 'ItmX'" _
       & " From CRD_PREA_PARAMETROS" _
       & " WHERE COD_PARAMETRO IN('18','19','20')"
Call sbCbo_Llena_New(cboS_ComponenteAdicional, strSQL, False, True)

cboS_ComponenteAdicional.AddItem ""
cboS_ComponenteAdicional.Text = ""

vPaso = False

txtD_Descripcion.Text = ""
Call cboDeduccion_Click
Call cboS_ComponenteAdicional_Click

Call sbSTCargaCboGarantiav2(cboGarantia, "-1")
Call sbLlenarComboTodosV2(cboSalario, "spCRDPreaTIPO_SALARIO", "TIPO_SALARIO", "DescTipoSalario")

Call sbCargaCboComites

 
dtpCorte.Value = fxFechaServidor

'Carga Garantias de Fondos
strSQL = "exec spCRDGarantiaFND"
Call sbCbo_Llena_New(cboFondo, strSQL, False, True)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbInicializaComboExpediente()
    m_Cargando = True
    cboSubExpediente.Clear
    cboSubExpediente.AddItem "Nuevo Expediente"
    cboSubExpediente.ItemData(cboSubExpediente.NewIndex) = -1
    cboSubExpediente.ListIndex = 0
    clsMensajes.Estado = "P"
     m_Cargando = False
     
End Sub

Function fxDescLineaCredito(ByVal strCodigo As String) As String
On Error GoTo vError

glogon.strSQL = "select descripcion from catalogo where codigo = '" & Trim(strCodigo) & "'"

If execSql(glogon.strSQL, True) Then
    fxDescLineaCredito = IIf(IsNull(glogon.Recordset!Descripcion), "", glogon.Recordset!Descripcion)
Else
    MsgBox "No se encontró la descripción del código de la linea de crédito digitada. - " & strCodigo, vbCritical
End If


    Exit Function
vError:
    MsgBox "Ocurrió un error validar información digitada. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Function

Private Sub sbBusqueda(ByVal Control As String)
'Set GLOBALES.gfrmFormulario = Me
gBusquedas.Resultado = ""
gBusquedas.Convertir = "N"

Select Case Control
  Case "txtLinea" 'Codigo de linea de credito
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
        gBusquedas.Filtro = " and Activo = 1 and Retencion = 'N'"
        frmBusquedas.Show vbModal
        txtLinea.Text = gBusquedas.Resultado
        If Len(Trim(txtLinea.Text)) > 0 Then
          txtDesLineaCredito.Text = fxDescLineaCredito(Trim(txtLinea.Text))
        End If
   
  Case "txtDesLineaCredito" 'Descripcion Linea Credito
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Columna = "descripcion"
        frmBusquedas.Show vbModal
        txtLinea.Text = gBusquedas.Resultado
        txtLinea.Text = gBusquedas.Resultado
        If Len(Trim(txtLinea.Text)) > 0 Then
          txtDesLineaCredito.Text = fxDescLineaCredito(Trim(txtLinea.Text))
        End If

    Case "txtCedula"
        
        gBusquedas.Convertir = "N"
        gBusquedas.Col1Name = "Cédula Colilla"
        gBusquedas.Col2Name = "Cédula Real"
        gBusquedas.Col3Name = "Nombre"
        gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
        gBusquedas.Orden = "cedula"
        gBusquedas.Columna = "cedula"
        frmBusquedas.Show vbModal
        txtCedula.Text = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          Call txtCedula_LostFocus
        End If
        
        
        
    Case "txtNombre"
        gBusquedas.Convertir = "N"
        gBusquedas.Col1Name = "Cédula Colilla"
        gBusquedas.Col2Name = "Cédula Real"
        gBusquedas.Col3Name = "Nombre"
        gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
        gBusquedas.Orden = "nombre"
        gBusquedas.Columna = "nombre"
        frmBusquedas.Show vbModal
        txtCedula.Text = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          Call txtCedula_LostFocus
        End If
    Case "txtExpediente"
      Call sbMostraVentanBusqueda
End Select

End Sub

Private Function fxSelectItemSubExpediente(ByVal ListIndex As String) As String
Select Case Trim(ListIndex)
    Case "-1"
        fxSelectItemSubExpediente = "E"
    Case "-2"
        fxSelectItemSubExpediente = "S"
End Select
End Function

Public Function fxSexoItemData(ByVal ListIndex As Integer) As String
Select Case Trim(ListIndex)
    Case "0"
        fxSexoItemData = "M"
    Case "1"
        fxSexoItemData = "F"
End Select
End Function

Private Sub sbSeleccionaSexo(ByVal Combo As Object, ByVal ItemData As String)
Select Case ItemData
    Case "M"
        Combo.ListIndex = 0
    Case "F"
        Combo.ListIndex = 1
        
End Select
End Sub

 
Private Function fxValidaDatos(ByVal TabValidar As Integer) As Boolean
Dim m_Valor As String

On Error GoTo vError

fxValidaDatos = True

If m_ventanaEnModo = ModificarRegistro Then
   If Trim(txtExpediente.Text) <> Trim(cboSubExpediente.Text) Then
    MsgBox "No es posible realizar cambios al expediente seleccionado.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    Exit Function
   End If
End If

If ((clsMensajes.Estado = "A") Or (clsMensajes.Estado = "D")) Then
    m_estadoPreanalisis = clsMensajes.Estado
    MsgBox "No es posible realizar cambios al expediente seleccionado.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    Exit Function
Else
    clsMensajes.Estado = "R"
End If

If Val(cboCantidadFiadores.Text) < m_FiadoresRegistrador Then
    MsgBox "No es posible disminuir la cantidad de sub expedientes."
    cboCantidadFiadores.Text = m_FiadoresRegistrador
        fxValidaDatos = False
    Exit Function
'ElseIf Val(cboCantidadFiadores.Text) > m_FiadoresRegistrador Then
'    MsgBox "No es posible aumentar la cantidad de sub expedientes."
'    cboCantidadFiadores.Text = m_FiadoresRegistrador
'        fxValidaDatos = False
'    Exit Function
End If

m_estadoPreanalisis = clsMensajes.Estado
clsMensajes.Usuario = glogon.Usuario

'Validar de Cantidad Fiadores
clsMensajes.NSUB_EXP = cboCantidadFiadores.ListIndex

'Validar número de Salario
If cboSalario.ListCount = 0 Then
    clsMensajes.tipo_salario = clsNull.SetNull
Else
    clsMensajes.tipo_salario = SIFGlobal.fxCodText(cboSalario.Text)
End If
    
If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    m_Valor = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
Else
    m_Valor = txtExpediente.Text
End If

clsMensajes.cod_preanalisis = IIf(Len(txtExpediente.Text) = 0, clsNull.SetNull, txtExpediente.Text)

If fxSelectItemSubExpediente(cboSubExpediente.ItemData(cboSubExpediente.ListIndex)) = "E" Then

    clsMensajes.tipo_preanalisis = "E"
    clsMensajes.cod_preanalisis_ref = clsNull.SetNull
    
ElseIf fxSelectItemSubExpediente(cboSubExpediente.ItemData(cboSubExpediente.ListIndex)) = "S" Then

    If m_Valor = "" Then
        clsMensajes.cod_preanalisis_ref = vCodExpediente
        clsMensajes.cod_preanalisis = "" 'vCodExpediente
    Else
        clsMensajes.cod_preanalisis_ref = m_Valor
        clsMensajes.cod_preanalisis = txtExpediente.Text
    End If
    
    clsMensajes.tipo_preanalisis = "S"
    
    
Else
    If InStr(1, txtExpediente.Text, "-", vbTextCompare) = 0 Then
        clsMensajes.tipo_preanalisis = "E"
        clsMensajes.cod_preanalisis_ref = clsNull.SetNull
    Else
        clsMensajes.tipo_preanalisis = "S"
        clsMensajes.cod_preanalisis_ref = m_Valor
        clsMensajes.cod_preanalisis = txtExpediente.Text
    End If
End If

'If TabValidar = Datos Then 'Valida datos del tab del tab de datos
    'Validar número de cédula
clsMensajes.Cedula = Trim(txtCedula.Text)
If Len(txtCedula.Text) = 0 Then
    MsgBox "El Número de cédula es requerido.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    clsMensajes.Cedula = ""
    txtCedula.SetFocus
    Exit Function
End If
'Validar nombre
clsMensajes.Nombre = Trim(txtNombre.Text)
If Len(txtNombre.Text) = 0 Then
    MsgBox "El Nombre es requerido.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    clsMensajes.Nombre = ""
    txtNombre.SetFocus
    Exit Function
End If

clsMensajes.sexo = fxSexoItemData(cboSexo.ListIndex)
'Validar fecha de nacimiento
clsMensajes.fecha_nacimiento = dtpFecNac.Value
If Not (IsDate(dtpFecNac.Value)) Then
    MsgBox "Fecha de nacimiento no es válida.", vbInformation, gMsgTitulo
    tcMain.Item(0).Selected = True
    fxValidaDatos = False
    clsMensajes.fecha_nacimiento = Date
    dtpFecNac.SetFocus
    Exit Function
End If

'Validar Linea de credito
If Len(txtDesLineaCredito.Text) = 0 Then
    MsgBox "Línea de crédito es requerida.", vbInformation, gMsgTitulo
  '  tcmain.Item(0).Selected =True
    fxValidaDatos = False
    txtLinea.SetFocus
    Exit Function
Else
 clsMensajes.Cod_Linea = txtLinea.Text
End If

'Validar Linea de garantias
clsMensajes.GARANTIA = cboGarantia.ItemData(cboGarantia.ListIndex) 'cboGarantia.ItemData(cboGarantia.ListIndex)
If cboGarantia.ListCount = 0 Then
    MsgBox "Debe seleccionar una garantía.", vbInformation, gMsgTitulo
    'tcmain.Item(0).Selected =True
    fxValidaDatos = False
    clsMensajes.GARANTIA = clsNull.SetNull
    cboGarantia.SetFocus
    Exit Function
End If


'Validar Linea de Comite
'If cboComite.ItemData(cboComite.ListIndex) = 0 Then
'    MsgBox "Debe seleccionar un comite de aprobación.", vbInformation, gMsgTitulo
'    'tcmain.Item(0).Selected =True
'    fxValidaDatos = False
'    Exit Function
'Else
'    clsMensajes.ID_COMITE = cboComite.ItemData(cboComite.ListIndex)
'End If

clsMensajes.ID_COMITE = cboComite.ItemData(cboComite.ListIndex)



'Validar  respaldo segun garantía

clsMensajes.GARANTIA_FND = clsNull.SetNull

If cboGarantia.ItemData(cboGarantia.ListIndex) = "Y" Then
    clsMensajes.GARANTIA_FND = cboFondo.ItemData(cboFondo.ListIndex)
    If cboFondo.ListCount = 0 Then
        MsgBox "Debe seleccionar un respaldo de la lista.", vbInformation, gMsgTitulo
        'tcmain.Item(0).Selected =True
        fxValidaDatos = False
        clsMensajes.GARANTIA_FND = clsNull.SetNull
        cboFondo.SetFocus
        Exit Function
    End If
End If


'Validar Destino credito


If (cboDestino.ListCount = 0) Or (cboDestino.ItemData(cboDestino.ListIndex) = "") Then
    MsgBox "Debe seleccionar un destino de crédito.", vbInformation, gMsgTitulo
    
    fxValidaDatos = False
    clsMensajes.cod_destino = ""
    cboDestino.SetFocus
    Exit Function
Else
    clsMensajes.cod_destino = cboDestino.ItemData(cboDestino.ListIndex)
End If

clsMensajes.APL_POLIZA_VIDA = IIf(chkPolizaVida.Value = 0, 0, 1)
clsMensajes.MONTO_POLIZA_VIDA = IIf(Len(txtPolizaVida.Text) = 0, clsNull.NullNumerico, CDbl(Val(txtPolizaVida.Text)))
clsMensajes.apl_poliza_incendio = IIf(chkPolizaIncendio.Value = 0, 0, 1)
clsMensajes.MONTO_POLIZA_INCENDIO = IIf(Len(txtPolizaIncendio.Text) = 0, clsNull.NullNumerico, CDbl(Val(txtPolizaIncendio.Text)))
clsMensajes.apl_primer_cuota = chkPrimerCuota.Value
'Validar Monto  credito

If Val(txtMonto.Text) = 0 Then
    MsgBox "Debe digitar el monto del crédito.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    txtMonto.SetFocus
    Exit Function
Else
    clsMensajes.Monto = CDbl(txtMonto.Text)
End If
'Validar plazo Monto  credito

If Val(txtPlazo.Text) = 0 Then
    MsgBox "Debe digitar el plazo para el monto del crédito.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    txtPlazo.SetFocus
    Exit Function
Else
    clsMensajes.Plazo = CInt(Val(txtPlazo.Text))
End If
'Validar tasa del Monto  credito

If Val(txtTasa.Text) = 0 Then
    clsMensajes.TASA = 0
Else
    clsMensajes.TASA = CDbl(Val(txtTasa.Text))
End If

'Validar cuota
If Val(txtCuota.Text) = 0 Then
    MsgBox "La cuota no fue calculada correctamente.", vbInformation, gMsgTitulo
    fxValidaDatos = False
    txtTasa.SetFocus
    Exit Function
Else
    clsMensajes.Cuota = CDbl(txtCuota.Text)
End If

If Val(txtCompromiso.Text) = 0 Then
    clsMensajes.COMPROMISO = clsNull.NullNumerico
Else
   clsMensajes.COMPROMISO = CDbl(txtCompromiso.Text)
End If

If TabValidar = 1 Then 'Valida datos en el tab de calculo

If cboSalario.ListCount = 0 Then
    clsMensajes.tipo_salario = clsNull.SetNull
Else
    clsMensajes.tipo_salario = SIFGlobal.fxCodText(cboSalario.Text)
End If

End If

clsMensajes.FECHA_CORTE_COLIILA = dtpCorte.Value
 
clsMensajes.SALARIO_DEVENGADO_COLILLA = Val(txtSalarioDevengado.Text)
If Val(txtSalarioDevengado.Text) <> 0 Then
    clsMensajes.SALARIO_DEVENGADO_COLILLA = CDbl(txtSalarioDevengado.Text)
    clsMensajes.SALARIO_DEVENGADO_GRUPO = CDbl(txtSalarioDevengado.Text)
End If

clsMensajes.REBAJO_EXTRAS = Val(txtRebajoExtras.Text)
If Val(txtRebajoExtras.Text) <> 0 Then
    clsMensajes.REBAJO_EXTRAS = CCur(txtRebajoExtras.Text)
End If
clsMensajes.REBAJO_EXTRAS = Val(txtRebajoExtras.Text)
If Val(txtRebajoExtras.Text) <> 0 Then
    clsMensajes.REBAJO_EXTRAS = CDbl(txtRebajoExtras.Text)
End If
clsMensajes.REBAJO_EXTRAS = Val(txtRebajoExtras.Text)
If Val(txtRebajoExtras.Text) <> 0 Then
    clsMensajes.REBAJO_EXTRAS = CDbl(txtRebajoExtras.Text)
End If
clsMensajes.SALARIO_REAL = Val(txtSalarioReal.Text)
If Val(txtSalarioReal.Text) <> 0 Then
    clsMensajes.SALARIO_REAL = CDbl(txtSalarioReal.Text)
End If
clsMensajes.EXTRAS_FIJAS = Val(txtCompAdicionalBase.Text)
If Val(txtCompAdicionalBase.Text) <> 0 Then
    clsMensajes.EXTRAS_FIJAS = CDbl(txtCompAdicionalBase.Text)
End If

clsMensajes.DEVENGADO_MES = Val(txtDevengadoMes.Text)
If Val(txtDevengadoMes.Text) <> 0 Then
    clsMensajes.DEVENGADO_MES = CDbl(txtDevengadoMes.Text)
End If

clsMensajes.PORCENTAJE_LIBRE = 0
If Val(txtPorcSobreSalario.Text) <> 0 Then
    clsMensajes.PORCENTAJE_LIBRE = CDbl(txtPorcSobreSalario.Text)
End If
clsMensajes.DEDUCCIONES = 0
If Val(txtDeducciones.Text) <> 0 Then
    clsMensajes.DEDUCCIONES = CDbl(txtDeducciones.Text)
End If
clsMensajes.CRD_TRANSITO_CANCELADOS = 0
If Val(txtCrdTransitoCancelados.Text) <> 0 Then
    clsMensajes.CRD_TRANSITO_CANCELADOS = CDbl(txtCrdTransitoCancelados.Text)
End If

clsMensajes.CRD_TRANSITO_XCOBRAR = 0
If Val(txtCrdTransitoXCobrar.Text) <> 0 Then
    clsMensajes.CRD_TRANSITO_XCOBRAR = CDbl(txtCrdTransitoXCobrar.Text)
End If
clsMensajes.SALARIO_LIQUIDO = Val(txtSalarioLiquido.Text)
If Val(txtSalarioLiquido.Text) <> 0 Then
    clsMensajes.SALARIO_LIQUIDO = CDbl(txtSalarioLiquido.Text)
End If
clsMensajes.REFUNDICIONES = Val(txtRefundiciones.Text)
If Val(txtRefundiciones.Text) <> 0 Then
    clsMensajes.REFUNDICIONES = CDbl(txtRefundiciones.Text)
End If

clsMensajes.REFUNDICIONES_CUOTA = Val(txtRefundiciones.ToolTipText)
If Val(txtRefundiciones.ToolTipText) <> 0 Then
    clsMensajes.REFUNDICIONES_CUOTA = CDbl(txtRefundiciones.ToolTipText)
End If


clsMensajes.DESEMBOLSOS = Val(txtDesembolsos.Text)
If Val(txtDesembolsos.Text) <> 0 Then
    clsMensajes.DESEMBOLSOS = CDbl(txtDesembolsos.Text)
End If

clsMensajes.DESEMBOLSOS_CUOTA = Val(txtDesembolsos.ToolTipText)
If Val(txtDesembolsos.ToolTipText) <> 0 Then
    clsMensajes.DESEMBOLSOS_CUOTA = CDbl(txtDesembolsos.ToolTipText)
End If

clsMensajes.LIQUIDO_TOTAL = Val(txtTotalLiquido.Text)
If Val(txtTotalLiquido.Text) <> 0 Then
    clsMensajes.LIQUIDO_TOTAL = CDbl(txtTotalLiquido.Text)
End If


clsMensajes.FIANZAS = 0
If Val(txtFianzas.Text) <> 0 Then
    clsMensajes.FIANZAS = CDbl(txtFianzas.Text)
End If
clsMensajes.LIQUIDEZ_SIMPLE = 0
If Val(txtLiquidezSinFianza.Text) <> 0 Then
    clsMensajes.LIQUIDEZ_SIMPLE = CDbl(txtLiquidezSinFianza.Text)
End If

clsMensajes.LIQUIDEZ_CFIANZAS = 0
If Val(txtLiquidezConFianza.Text) <> 0 Then
    clsMensajes.LIQUIDEZ_CFIANZAS = CDbl(txtLiquidezConFianza.Text)
End If


clsMensajes.TOTAL_CARGA_CCSS = 0
clsMensajes.CARGA_ASOCIACION = 0
chkCargaAsociacion.Tag = "N"
If chkCargaAsociacion.Value = 1 Then
    If Val(lblCargaAsociacion.Caption) > 0 Then
        clsMensajes.CARGA_ASOCIACION = CDbl(lblCargaAsociacion.Caption)
    End If
    chkCargaAsociacion.Tag = "S"
End If

chkCargaFrap.Tag = "N"
clsMensajes.CARGA_FRAP = 0
If chkCargaFrap.Value = 1 Then
    If Val(lblCargaFrap.Caption) > 0 Then
        clsMensajes.CARGA_FRAP = CDbl(lblCargaFrap.Caption)
        chkCargaFrap.Tag = "S"
    End If
End If

clsMensajes.CARGA_CCSS = Val(lblCargaCCSS.Caption)
If chkCargaAsociacion.Value = 1 Then
    If clsMensajes.CARGA_CCSS > 0 Then
        clsMensajes.CARGA_CCSS = CDbl(lblCargaCCSS.Caption)
    End If
End If
clsMensajes.TOTAL_CARGA_CCSS = Val(txtTotal_Cargas_CCSS.Text)
If Val(txtTotal_Cargas_CCSS.Text) <> 0 Then
    clsMensajes.TOTAL_CARGA_CCSS = CDbl(txtTotal_Cargas_CCSS.Text)
End If
   
    
clsMensajes.CARGA_IMPUESTO_SALARIO = Val(lblCargaImpSalario.Caption)
If clsMensajes.CARGA_IMPUESTO_SALARIO > 0 Then
    clsMensajes.CARGA_IMPUESTO_SALARIO = CDbl(lblCargaImpSalario.Caption)
End If


'ElseIf TabValidar = Observaciones Then 'Valida datos en el tab de calculo
clsMensajes.OBSERVACION_ANALISTA = vObservacion(0)
clsMensajes.OBSERVACION_COMITE = vObservacion(1)
clsMensajes.OBSERVACION_JD = vObservacion(2)

clsMensajes.COD_OFICINA = GLOBALES.gOficinaTitular
clsMensajes.CUMPLIMIENTO_NOTAS = txtCumplimientoNotas.Text

Call LigarDatosClasificacion

Exit Function

vError:
    MsgBox "Ocurrió un error validar la información digitada. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    fxValidaDatos = False

End Function




Private Function fxValidaDatosBorrar() As Boolean
    Dim m_Valor As String
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
fxValidaDatosBorrar = True

If ((clsMensajes.Estado = "A") Or (clsMensajes.Estado = "D")) Then
    m_estadoPreanalisis = clsMensajes.Estado
    MsgBox "No es posible realizar cambios al expediente seleccionado.", vbInformation, gMsgTitulo
    fxValidaDatosBorrar = False
    Exit Function
End If

If Len(txtExpediente.Text) = 0 Then
   MsgBox "Debe seleccionar un expediente.", vbInformation, gMsgTitulo
    fxValidaDatosBorrar = False
    Exit Function
End If


    strSQL = "select count(*) from CRD_PREA_PREANALISIS where COD_PREANALISIS_REF = '" & Trim(txtExpediente.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        If rs.Fields(0) > 0 Then
            MsgBox "El expediente tiene subexpedientes asociados y debe de borrados antes de borrar el principal"
            fxValidaDatosBorrar = False
            Exit Function
        End If
    End If

clsMensajes.cod_preanalisis = txtExpediente.Text

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    clsMensajes.cod_preanalisis_ref = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
Else
    clsMensajes.cod_preanalisis_ref = txtExpediente.Text
    
End If


Exit Function
    
vError:
    MsgBox "Ocurrió un error validar la información para eliminar los datos seleccionados. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    fxValidaDatosBorrar = False

End Function


Private Sub LigarDatosClasificacion()
Dim i As Integer

vGrid.Col = 1
For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.Col = 2
    Select Case True
        Case InStr(1, vGrid.Text, "CAPACIDAD")
            vGrid.Col = 1
            clsMensajes.COD_CAPACIDAD = vGrid.Text
        
        Case InStr(1, vGrid.Text, "ENDEUDAMIENTO")
            vGrid.Col = 1
            clsMensajes.COD_ENDEUDAMIENTO = vGrid.Text
        
        Case InStr(1, vGrid.Text, "HISTORIAL")
            vGrid.Col = 1
            clsMensajes.COD_HISTORIAL = vGrid.Text
        
        Case InStr(1, vGrid.Text, "MOROSIDAD")
            vGrid.Col = 1
            clsMensajes.COD_MORA = vGrid.Text
        
        Case InStr(1, vGrid.Text, "GARANTIA")
            vGrid.Col = 1
            clsMensajes.COD_GARANTIA = vGrid.Text
        
    End Select
Next i
End Sub
Public Function fxAgregaColleccionBorrar(ByVal cod_preanalisis As String, ByVal cod_preanalisis_ref As String) As String
On Error GoTo error
Dim Vcoleccion As New Collection
With Vcoleccion
    .Add fxFormatearValor(cod_preanalisis, caracter)
    .Add fxFormatearValor(cod_preanalisis_ref, caracter)

End With
fxAgregaColleccionBorrar = fxFormatearValuesCollection(Vcoleccion)

Exit Function
error:
    MsgBox fxSys_Error_Handler(Err.Description)
End Function

Private Sub sbBorrar()

On Error GoTo vError

If Not fxValidaDatosBorrar Then Exit Sub

clsEntidad.tablaName = "spCRDPreaPREANALISIS"

If m_ventanaEnModo = ModificarRegistro Then
    If (MsgBox("¿ Desea borrar la información seleccionada?", vbQuestion + vbYesNo, gMsgTitulo) = vbYes) Then
        Call clsEntidad.fxRemover(fxAgregaColleccionBorrar(clsMensajes.cod_preanalisis, clsMensajes.cod_preanalisis_ref))
            MsgBox "La información fue borrada correctamente.", vbInformation, gMsgTitulo
            cboSubExpediente.ListIndex = 0
            
'            Call tlb_ButtonClick(tlb.Buttons("deshacer"))
            
        End If
Else
    MsgBox "La información no se encuentra almacenada.", vbInformation, gMsgTitulo
End If


    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)

End Sub


Private Sub sbEstudio_Guarda_Nuevo()
Dim strSQL As String

Dim pTipoExpediente As String, pExpedienteRef As String, pGarFondo As String, pGarFondoContrato As String
Dim pOficina  As String, pEdad As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass


pOficina = GLOBALES.gOficinaTitular
pEdad = DateDiff("yyyy", dtpFecNac.Value, Date)


If cboSubExpediente.Text = "Nuevo Expediente" Or InStr(txtExpediente.Text, "-") = 0 Then
  pTipoExpediente = "E"
  pExpedienteRef = "Null"
Else
  pTipoExpediente = "S"
  pExpedienteRef = "'" & txtExpediente.Text & "'"
End If

If cboGarantia.ItemData(cboGarantia.ListIndex) = "F" Then
    pGarFondo = "'" & cboFondo.ItemData(cboFondo.ListIndex) & "'"
    pGarFondoContrato = cboFondoContrato.ItemData(cboFondoContrato.ListIndex)
Else
    pGarFondo = "Null"
    pGarFondoContrato = "Null"
End If

If Not IsNumeric(txtMontoConstruccion.Text) Then
    txtMontoConstruccion.Text = 0
End If

If Not IsNumeric(txtPrendaValor.Text) Then
    txtPrendaValor.Text = 0
End If


strSQL = "exec spCrdPreaPreanalisisNuevo '" & cboSalario.ItemData(cboSalario.ListIndex) & "', '" & pTipoExpediente & "', " & pExpedienteRef _
       & ", '" & glogon.Usuario & "', '" & txtCedula.Text & "', '" & txtLinea.Text & "', '" & cboDestino.ItemData(cboDestino.ListIndex) _
       & "', '" & txtNombre.Text & "', '" & Mid(cboSexo.Text, 1, 1) & "', '" & Format(dtpFecNac.Value, "yyyy-mm-dd") _
       & "', " & chkPolizaVida.Value & ", " & chkPolizaIncendio.Value & ", " & chkPrimerCuota.Value _
       & ", " & CCur(txtMonto.Text) & ", " & CCur(txtTasa.Text) & ", " & CLng(txtPlazo.Text) & ", " & CCur(txtCuota.Text) _
       & ", " & CCur(txtPolizaVida.Text) & ", " & CCur(txtPolizaIncendio.Text) & ", " & CCur(txtCompromiso.Text) & ", Null" _
       & ", '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "', '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "', " & cboCantidadFiadores.Text _
       & ", " & pGarFondo & ", '" & pOficina & "', " & clsMensajes.TASA_PTS_BONO & ", " & pGarFondoContrato & ", " & pEdad _
       & ", " & txtEdad.Tag & ", '" & txtEdad.ToolTipText & "', " & txtPlazo.Text _
       & ", 0, 0, 0, " & IIf(cboCPH.Text = "", 0, cboCPH.Text) & ", 1, " & CCur(txtMontoConstruccion.Text) _
       & ", " & chkPolizaVehiculo.Value & ", " & CCur(txtPolizaPrenda.Text) & ", " & CCur(txtPrendaValor.Text) & ", '" _
       & txtClasificacion.Text & "', '" & txtCRM.Text & "'"
Call OpenRecordSet(rs, strSQL)
    txtExpediente.Text = rs!cod_preanalisis
rs.Close

Me.MousePointer = vbDefault



Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbEstudio_Guarda_Modifica()
Dim strSQL As String

Dim pTipoExpediente As String, pExpedienteRef As String, pGarFondo As String, pGarFondoContrato As String
Dim pOficina  As String, pEdad As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass


pOficina = GLOBALES.gOficinaTitular
pEdad = DateDiff("yyyy", dtpFecNac.Value, Date)


If cboSubExpediente.Text = "Nuevo Expediente" Or InStr(txtExpediente.Text, "-") = 0 Then
  pTipoExpediente = "E"
  pExpedienteRef = "Null"
Else
  pTipoExpediente = "S"
  pExpedienteRef = "'" & txtExpediente.Text & "'"
End If

If cboGarantia.ItemData(cboGarantia.ListIndex) = "F" Then
    pGarFondo = "'" & cboFondo.ItemData(cboFondo.ListIndex) & "'"
    pGarFondoContrato = cboFondoContrato.ItemData(cboFondoContrato.ListIndex)
Else
    pGarFondo = "Null"
    pGarFondoContrato = "Null"
End If

Dim pCA_Aplica As Integer, pEditable As Integer, pSalarioExterno As Integer, pEstadoV2 As String
Dim pMontoLiq_Fiador_Ext As Currency, pMejoraCuota As Currency, pMejoraCuotaInd As String

Dim pCal_Garantia As String, pCal_Endeudamiento As String, pCal_Historial As String, pCal_Mora As String, pCal_Capacidad As String

pCal_Capacidad = clsMensajes.COD_CAPACIDAD
pCal_Endeudamiento = clsMensajes.COD_ENDEUDAMIENTO
pCal_Garantia = clsMensajes.COD_GARANTIA
pCal_Historial = clsMensajes.COD_HISTORIAL
pCal_Mora = clsMensajes.COD_MORA


pMejoraCuota = CCur(txtCuotaDiferencia.Text)

If pMejoraCuota < 0 Then
    pMejoraCuotaInd = "S"
Else
    pMejoraCuotaInd = "N"
End If

pMontoLiq_Fiador_Ext = 0
pEstadoV2 = "RECI"

Select Case lblEstado.Tag
    Case "R", "P"
        pEditable = 1
    Case Else
        pEditable = 0
End Select

pSalarioExterno = 0

If chkS_Constancia.Value = xtpChecked Then
    pSalarioExterno = 1
End If
If chkS_OrdenPatronal.Value = xtpChecked Then
    pSalarioExterno = 2
End If


If CCur(txtS_ComponenteAdicionalPorc.Text) = 0 Then
    pCA_Aplica = 0
Else
    pCA_Aplica = 1
End If

Dim pCA_ID As String
If cboS_ComponenteAdicional.ItemData(cboS_ComponenteAdicional.ListIndex) = "" Then
    pCA_ID = "Null"
Else
    pCA_ID = cboS_ComponenteAdicional.ItemData(cboS_ComponenteAdicional.ListIndex)
End If

strSQL = "exec spCrdPreaPreanalisisModifica '" & txtExpediente.Text & "', '" & cboSalario.ItemData(cboSalario.ListIndex) & "', '" & pTipoExpediente & "', " & pExpedienteRef _
       & ", '" & glogon.Usuario & "', Null, '" & RTrim(txtCedula.Text) & "', '" & txtLinea.Text & "', '" & cboDestino.ItemData(cboDestino.ListIndex) _
       & "', '" & txtNombre.Text & "', '" & Mid(cboSexo.Text, 1, 1) & "', '" & Format(dtpFecNac.Value, "yyyy-mm-dd") _
       & "', " & chkPolizaVida.Value & ", " & chkPolizaIncendio.Value & ", " & chkPrimerCuota.Value & ", '" & Format(dtpCorte.Value, "yyyy-mm-dd") _
       & "', " & CCur(txtMonto.Text) & ", " & CCur(txtTasa.Text) & ", " & CLng(txtPlazo.Text) & ", " & CCur(txtCuota.Text) _
       & ", " & CCur(txtPolizaVida.Text) & ", " & CCur(txtPolizaIncendio.Text) & ", " & CCur(txtCompromiso.Text) & ", '" & lblEstado.Tag _
       & "', " & txtAsignado.Text & ", " & CCur(txtSalarioDevengado.Text) & ", " & CCur(txtT_Extras) & ", " & CCur(txtCompAdicionalBase.Text) _
       & ", " & CCur(txtD_TotalCargas.Text) & ", " & CCur(lblCargaCCSS.Caption) & ", " & CCur(lblCargaAsociacion.Caption) & ", " & CCur(lblCargaFrap.Caption) _
       & ", " & CCur(lblCargaImpSalario.Caption) & ", " & CCur(txtS_Privado_Porc.Text) & ", " & CCur(txtDeducciones.Text) _
       & ", " & CCur(txtC_CuotaCancelaTotal.Text) & ", " & CCur(txtC_CuotaPorCobrarTotal.Text) & ", " & CCur(txtSalarioLiquido.Text) _
       & ", " & CCur(txtRefundiciones.Text) & ", " & CCur(txtR_TotalCuotas.Text) _
       & ", " & CCur(txtDesembolsos.Text) & ", " & CCur(txtDS_TotalCuota.Text) & ", " & CCur(txtTotalLiquido.Text) _
       & ", " & CCur(txtLiquidezSinFianza.Text) & ", " & CCur(txtFianzas.Text) & ", " & CCur(txtLiquidezConFianza.Text) _
       & ", '" & vObservacion(0) & "', '" & vObservacion(1) & "', '" & vObservacion(2) & "', '" & txtClasificacion.Text & "', '" & cboGarantia.ItemData(cboGarantia.ListIndex) _
       & "', '" & pCal_Garantia & "', '" & pCal_Endeudamiento & "', '" & pCal_Historial & "', '" & pCal_Mora & "', '" & pCal_Capacidad _
       & "', " & CCur(txtSalarioReal.Text) & ", " & CCur(txtDevengadoMes.Text) & ", " & cboCantidadFiadores.Text _
       & ", " & pGarFondo & ", '" & pOficina & "', " & clsMensajes.TASA_PTS_BONO & ", " & pGarFondoContrato & ", " & pEdad _
       & ", " & CCur(txtFrapPorc.Text) & ", " & CCur(txtS_Constancia.Text) & ", " & CCur(txtS_OrdenPatronal.Text) _
       & ", 0, " & CCur(txtLiquidezPorcConFianza.Text) & ", " & CCur(txtLiquidezPorcSinFianza.Text) & ", 1, 0, 0" _
       & ", " & CCur(txtCompAdicional.Text) & ", " & pCA_Aplica & ", " & CCur(txtS_ComponenteAdicionalPorc.Text) & ", " & pCA_ID _
       & ", '" & txtCIC_Puntaje.Text & "', '" & txtCIC_NivelHistorico.Text & "', " & cboCPH.Text & ", " & txtDiasIntereses.Text _
       & ", '" & txtCRM.Text & "', " & pEditable & ", " & pSalarioExterno & ", " & CCur(txtMontoConstruccion.Text) _
       & ", '" & pEstadoV2 & "', " & CCur(txtLiquidezPorcSinFianzaComp.Text) & ", " & CCur(txtLiquidezPorcConFianzaComp.Text) _
       & ", " & CCur(txtLiquidezSinFianzaComp.Text) & ", " & CCur(txtLiquidezConFianzaComp.Text) & ", " & pMontoLiq_Fiador_Ext _
       & ", " & chkPolizaVehiculo.Value & ", " & CCur(txtPolizaPrenda.Text) & ", " & CCur(txtPrendaValor.Text)
       
strSQL = strSQL & ", '" & txtClasificacion.Text & "', '" & txtCRM.Text & "', " & IIf(txtEjecutivo.Tag = "", "Null", txtEjecutivo.Tag) _
       & ", " & CCur(txtS_Privado.Text) & ", " & pMejoraCuota & ", " & CCur(txtIntereses.Text) & ", " & CCur(txtComisiones.Text) _
       & ", '" & pMejoraCuotaInd & "', " & CCur(txtR_TotalMora.Text) & ", " & CCur(txtSalarioMinimoInembargable.Text) _
       & ", " & CCur(txtSalarioNormativa.Text) & ", '" & txtCumplimientoNotas.Text & "', " & chkPolizaDesempleo.Value & ", " & CCur(txtPolizaDesempleo.Text)
       
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxGuardar() As Boolean
Dim strSQL As String, rsX As New ADODB.Recordset
Dim vMensaje As String


On Error GoTo vError

vMensaje = ""

fxGuardar = fxValidaDatos(m_PreviousTab)

If Not fxGuardar Then Exit Function


Screen.MousePointer = vbHourglass

clsEntidad.tablaName = "spCRDPreaPREANALISIS"
  
'Nuevo
If Len(vMensaje) = 0 Then
  strSQL = "exec spCrdFormaliza_Valida_Rangos '" & txtCedula.Text & "','" & txtLinea.Text & "'," _
         & CCur(txtMonto.Text) & "," & CCur(txtTasa.Text) & "," & CInt(txtPlazo.Text) _
         & ",'" & cboDestino.ItemData(cboDestino.ListIndex) & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) _
         & "',0"
  Call OpenRecordSet(rsX, strSQL)
  If Len(rsX!Mensaje) > 0 Then
      vMensaje = vMensaje & vbCrLf & rsX!Mensaje
  End If
  rsX.Close
End If

'Fix
If Not IsNumeric(txtPolizaDesempleo.Text) Then
    Call sbCalculaPolizaDesempleo
End If


txtCumplimientoNotas.Text = vMensaje
clsMensajes.CUMPLIMIENTO_NOTAS = txtCumplimientoNotas.Text
clsMensajes.APL_POLIZA_DESEMPLEO = chkPolizaDesempleo.Value
clsMensajes.MONTO_POLIZA_DESEMPLEO = txtPolizaDesempleo.Text
  
Select Case True
  Case m_ventanaEnModo = NuevoRegistro
    
    'If clsEntidad.fxAgregar(clsMensajes.fxConcatenaColleccion) Then
            
        Call sbEstudio_Guarda_Nuevo
            
'        Call sbTraerMaxExpediente
        
        Call txtExpediente_LostFocus
        Call sbAccionVentana(ModificarRegistro)
        
        If m_MuestraMensaje = True Then
            MsgBox "La información fue registrada correctamente.", vbInformation, gMsgTitulo
        End If


        'Inicializa datos vinculados
        If gPreAnalisis.Expediente = "" Then
            MsgBox "El número de expediente no fue cargado en las variables globales.", vbInformation, gMsgTitulo
        Else
        
            'Refundiciones (Lista)
            glogon.strSQL = "spCRDPreaRefundiciones " & fxFormatearValor(gPreAnalisis.Expediente, caracter) & "," & "'I'"
            If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
                MsgBox "Ocurrió un error al inicializar refundiciones.", vbInformation, gMsgTitulo
            End If
            
            'Fianzas (Lista)
            glogon.strSQL = "spCRDPreaFianzas " & fxFormatearValor(gPreAnalisis.Expediente, caracter) & "," & "'I'"
            If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
                MsgBox "Ocurrió un error al inicializar fianzas.", vbInformation, gMsgTitulo
            End If
        
            'Cancelacion y Operaciones por Cobrar (Lista)
            glogon.strSQL = "spCRDPreaCreditosTransito " & fxFormatearValor(gPreAnalisis.Expediente, caracter) & "," & "'I', 0"
            If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
                MsgBox "Ocurrió un error al inicializar Operaciones por Cobrar.", vbInformation, gMsgTitulo
            End If
            
            'Deducciones
            strSQL = "exec spCRDPreaImportCreditosVigentes '" & txtExpediente.Text & "', " & m_NumPagos
            Call ConectionExecute(strSQL)
        
        
        End If 'Else
    
    
    'End If
    
 
 Case m_ventanaEnModo = ModificarRegistro
    'Call clsEntidad.fxModificar(clsMensajes.fxConcatenaColleccion)
    
    Call sbEstudio_Guarda_Modifica
    
    Call sbAccionVentana(ModificarRegistro)
    Call txtExpediente_LostFocus
    
    If m_MuestraMensaje = True Then
        MsgBox "La información fue actualizada correctamente.", vbInformation, gMsgTitulo
    End If
    
    
End Select

'FIX TEMPORAL DE COLUMNAS NUEVAS
With glogon
   If Not IsNumeric(txtPolizaDesempleo.Text) Then
    txtPolizaDesempleo.Text = "0"
   End If

    .strSQL = "UPDATE CRD_PREA_PREANALISIS SET CUMPLIMIENTO_NOTAS = '" & txtCumplimientoNotas.Text _
            & "', MONTO_POLIZA_DESEMPLEO = " & CCur(txtPolizaDesempleo.Text) _
            & " , APL_POLIZA_DESEMPLEO = " & chkPolizaDesempleo.Value _
            & " where cod_Preanalisis = '" & txtExpediente.Text & "'"
   Call ConectionExecute(.strSQL)
End With


m_CargoSalario = False
m_MuestraMensaje = False


Screen.MousePointer = vbDefault

Exit Function

vError:
    Screen.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbMostraVentanBusqueda()

On Error GoTo vError

tcMain.Item(0).Selected = True

frmPreaConsultaExpeditentes.Show vbModal

If frmPreaConsultaExpeditentes.m_Expediente <> "" Then
    txtExpediente.SetFocus
    txtExpediente.Text = frmPreaConsultaExpeditentes.m_Expediente
    Call txtExpediente_LostFocus
    txtCedula.SetFocus
    
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description)
    
End Sub


Private Sub btnComite_Click(Index As Integer)

Select Case Index
    Case 0 'Asigna comite
    
    Case 1 'Resolucion
        Call btnGestion_Click(2)
        
    Case 2 '

End Select

End Sub


Private Function fxValidaFormalizacion() As Boolean



On Error GoTo vError

strSQL = "select dbo.fxCrdPreaFormalizacionValida('" & txtExpediente.Text & "') as 'Valida'"
Call OpenRecordSet(rs, strSQL)

fxValidaFormalizacion = IIf(rs!Valida = 1, True, False)


Exit Function

vError:
    fxValidaFormalizacion = False

End Function

Private Sub btnAbandonar_Click()
Dim i As Integer

On Error GoTo vError

If txtExpediente.Text = "" Then
    MsgBox "Consulte un Expediente!", vbExclamation
    Exit Sub
End If

If lblEstado.Tag = "D" Then
    MsgBox "No se puede ABANDONAR un expediente que ya ha sido DESCARTADO. ", vbExclamation
    Exit Sub
End If

If lblEstado.Tag = "B" Then
    MsgBox "Ya este estudio ha sido ABANDONADO anteriormente, no se puede realizar la acción nuevamente. ", vbExclamation
    Exit Sub
End If

If lblEstado.Tag = "A" And fxValidaFormalizacion Then
    MsgBox "No se puede ABANDONAR un expediente que ya ha sido FORMALIZADO. ", vbExclamation
    Exit Sub
End If

If fxSelectItemSubExpediente(cboSubExpediente.ItemData(cboSubExpediente.ListIndex)) = "S" Then
    MsgBox "No se puede ABANDONAR un expediente secundario, por favor seleccione el expediente principal e intente de nuevo. ", vbExclamation
    Exit Sub
End If


i = MsgBox("¿Desea ABANDONAR el estudio seleccionado?", vbYesNo)
If i = vbYes Then
    strSQL = "exec spCrdPreaCambiaEstadoPreanalisis '" & txtExpediente.Text & "', 'B'"
    Call ConectionExecute(strSQL)
    
    MsgBox "Se ha ABANDONADO el expediente correctamente.", vbInformation
    
    
    gPreAnalisis.Expediente = txtExpediente.Text
    Call sbFormsCall("frmPrea_Abandona_Motivos", vbModal, , , False, Me, True)
    
    
    lblEstado.Tag = "B"
    lblEstado.Caption = "Abandonado"
    
    
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnAdjunto_Elimina_Click()
Dim i As Integer, pArchivos As String

On Error GoTo vError

pArchivos = ""

If lblEstado.Tag = "A" Or lblEstado.Tag = "D" Or lblEstado.Tag = "B" Then
    MsgBox "Este Expediente no puede ser modificado!", vbExclamation, "Error"
    Exit Sub
End If


With lswArchivos.ListItems

    For i = 1 To .Count
        If .Item(i).Checked Then
          If Len(pArchivos) > 0 Then
            pArchivos = .Item(i).Text & ", " & pArchivos
          Else
            pArchivos = .Item(i).Text
          End If
        
        End If
    Next i

If Len(pArchivos) = 0 Then
  MsgBox "Seleccione los Adjuntos que desea eliminar!", vbExclamation
Else

  strSQL = "delete CRD_PREA_V2_ADJUNTOS  Where ID_EXPEDIENTE = '" & txtExpediente.Text _
         & "' and ID_ADJUNTO in(" & pArchivos & ")"
  Call ConectionExecute(strSQL)
  
  MsgBox "Adjuntos Eliminados Satisfactoriamente!", vbInformation
    
  Call sbAdjuntos_List
End If

End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnAdjunto_Guardar_Click()

If txtArchivo.Text = "" Then
    MsgBox "Por favor, selecciona un archivo primero.", vbExclamation, "Error"
    Exit Sub
End If

If lblEstado.Tag = "A" Or lblEstado.Tag = "D" Or lblEstado.Tag = "B" Then
    MsgBox "Este Expediente no puede ser modificado!", vbExclamation, "Error"
    Exit Sub
End If


On Error GoTo vError

Dim rs As New ADODB.Recordset

lblLoading.Caption = "Subiendo Archivo..Espere!"
DoEvents

Me.MousePointer = vbHourglass



'Version 2
Dim Cmd As ADODB.Command
Dim fileData() As Byte

Dim pFecha As Date, pFRegistro As String

pFecha = fxFechaServidor
pFRegistro = Format(pFecha, "yyyy-mm-dd") & " " & Format(pFecha, "hh:mm:ss")

' Leer el contenido del archivo en un arreglo de bytes
Open txtArchivo.Text For Binary Access Read As #1
ReDim fileData(LOF(1) - 1)
Get #1, , fileData
Close #1


' Preparar la consulta SQL para insertar el archivo
Set Cmd = New ADODB.Command
Cmd.ActiveConnection = glogon.Conection

Cmd.CommandText = "INSERT INTO CRD_PREA_V2_ADJUNTOS (ID_EXPEDIENTE, DOC_ADJUNTO, NOM_ADJUNTO, USUARIO_REG, FECHA_REG)" _
                & " VALUES (?, ?, ?, ?, ?)"

Cmd.Parameters.Append Cmd.CreateParameter("@ID_EXPEDIENTE", adVarChar, adParamInput, 20, txtExpediente.Text)
Cmd.Parameters.Append Cmd.CreateParameter("@DOC_ADJUNTO", adLongVarBinary, adParamInput, UBound(fileData) + 1, fileData)
Cmd.Parameters.Append Cmd.CreateParameter("@NOM_ADJUNTO", adVarChar, adParamInput, 1000, fxFileName_Valido(Dir(txtArchivo.Text, vbArchive)))
Cmd.Parameters.Append Cmd.CreateParameter("@USUARIO_REG", adVarChar, adParamInput, 30, glogon.Usuario)
Cmd.Parameters.Append Cmd.CreateParameter("@FECHA_REG", adVarChar, adParamInput, 20, pFRegistro)

Cmd.Execute


'-------Fin Version 2


    
Me.MousePointer = vbDefault

MsgBox "Archivo subido exitosamente.", vbInformation, "Éxito"
    
txtArchivo.Text = ""
lblLoading.Caption = ""

Call sbAdjuntos_List

    
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    lblLoading.Caption = "Error en la carga!"

End Sub

Private Sub btnArchivo_Click()
With frmContenedor.CD
        
        .InitDir = "C:\"
        .DialogTitle = "Localice el Archivo.."
        .Filter = "*.*"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        txtArchivo.Text = .FileName
End With

End Sub

Public Sub btnBarra_Click(Index As Integer)

m_MuestraMensaje = False

Select Case Index

    Case 0 '"INSERTAR", "NUEVO"
      Call sbAccionVentana(NuevoRegistro)
      Call sbLimipiaControles(Me, True)
'      Call sbToolBar(Me.tlb, "edicion")
      Call sbInicializaComboExpediente
      
      txtExpediente.Locked = False
      
      txtFrapPorc.Text = Format(0, "Standard")
      txtS_Privado_Porc.Text = "100"
      
      tcMain.Item(0).Selected = True
      gbSalarios(0).Enabled = False
      
      txtSalarioMinimoInembargable.Text = Format(GlobalSalarioMinimoInembargable, "Standard")
      txtSalarioNormativa.Text = Format(GlobalSalarioNormativo, "Standard")

      m_Editable = True
      
      lblEstado.Tag = "R"
      lblEstado.Caption = "(Recibiendo)"
      
      
      btnNotificacion.Visible = False
      
      
      gExtras.MaxRows = 1
      gExtras.Row = 1
      gExtras.Col = 3
      gExtras.Text = "0"
        
      gExtras.Col = 2
      gExtras.CellType = CellTypeComboBox
      gExtras.TypeComboBoxList = mTipoExtraLista
      gExtras.TypeComboBoxEditable = False
        
        
      
      gSalarios.MaxRows = 0
      cboS_ComponenteAdicional.Text = ""
      cboComite.Text = ""
      
      lswIncapacidades.ListItems.Clear
      
      
      
    Case 1 ' "MODIFICAR", "EDITAR"
'      Call sbToolBar(Me.tlb, "edicion")
      txtCedula.SetFocus
      Call sbBloquearTab
    
    Case 2 '"BORRAR"
         Call sbBorrar
         
    Case 3 '"GUARDAR", "SALVAR"
        m_MuestraMensaje = True
        If tcMain.Item(0).Selected = True Then
            If m_CambioDatos = False Then Exit Sub
        ElseIf tcMain.Item(1).Selected = True Then
            If m_CambioCalculo = False Then Exit Sub
        ElseIf tcMain.Item(2).Selected = True Then
            If m_CambioObservaciones = False Then Exit Sub
        End If
        
        If Not m_Editable Then Exit Sub
          
        Call fxGuardar
        
    Case 4 '"DESHACER"
      
      Call sbAccionVentana(NuevoRegistro)
      
      m_Editable = True
      'Call sbToolBar(Me.tlb, "nuevo")
      m_DesplegoMensaje = True
      Call sbLimipiaControles(Me, True)
      Call sbInicializaComboExpediente
      Call sbBloquearControles(Me, Expediente)
      tcMain.Item(0).Selected = True
      txtCedula.SetFocus
      m_CambioDatos = False
      m_CambioCalculo = False
      m_CambioObservaciones = False
      'TxtExpediente.Locked = True
      
'      Call sbToolBar(Me.tlb, "nuevo")
         
        
    Case 5 '"REPORTES"
        If m_ventanaEnModo = ModificarRegistro Then
           
           If Trim(txtExpediente.Text) <> Trim(cboSubExpediente.Text) Then
            MsgBox "Debe selecionar un expediente o sub expediente válido.", vbInformation, gMsgTitulo
            Exit Sub
           End If
           
           If Len(txtExpediente.Text) = 0 Then Exit Sub
            
            
            If m_Editable Then
                Call fxGuardar 'Actualiza Datos
            End If
            
            'Verifica y Guarda en caso de cambios
            'TODO: Call ssTab_Click(ssTab.Tab)
            Call sbTabChange(tcMain.SelectedItem)

            gPreAnalisis.Expediente = txtExpediente.Text
            frmPreaSubReporte.Show vbModal
        End If
         
    Case 6 'CONSULTAR
        Call sbMostraVentanBusqueda

End Select

Call RefrescaTags(Me)


End Sub

Private Sub btnComiteCambio_Click()
Dim strSQL As String, vNuevoEstado As String, IndicadorEditable As Integer

If txtExpediente.Text = "" Then Exit Sub

On Error GoTo vError



'Poner Validacion de Adjuntos
'
' Para Asignar un comité
' - Valida la liquidez
' - Valida el Maximo permitido del Comite
' - Valida una Etiqueta (que al menos tenga una)
' - Valida de la Justificacion de la Edad
' - Traslado de Salarios >> Linea: CRRM


strSQL = "exec spCrdPrea_Comite_Asigna_Valida '" & txtExpediente.Text & "', " & cboComite.ItemData(cboComite.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Len(rs!Mensaje) > 0 Then
    MsgBox rs!Mensaje, vbExclamation
    Exit Sub
End If

If cboComite.ItemData(cboComite.ListIndex) = 7 Or cboComite.ItemData(cboComite.ListIndex) = 8 Then
                vNuevoEstado = "RECI"
            ElseIf cboComite.ItemData(cboComite.ListIndex) = 1 Or cboComite.ItemData(cboComite.ListIndex) = 4 _
                    Or cboComite.ItemData(cboComite.ListIndex) = 10 Or cboComite.ItemData(cboComite.ListIndex) = 16 Then
                vNuevoEstado = "PRCO"
Else
    vNuevoEstado = ""
End If

If vNuevoEstado <> "" Then
    If vNuevoEstado = "RECI" Then
        IndicadorEditable = 1
    Else
        IndicadorEditable = 0
    End If
End If

    
strSQL = "exec spCrdPreaGestionaComiteResolutivo " & cboComite.ItemData(cboComite.ListIndex) & ", '" & txtExpediente.Text & "', '" & vNuevoEstado & "', " & IndicadorEditable
Call ConectionExecute(strSQL)
    


If cboComite.ItemData(cboComite.ListIndex) <> 0 And cboComite.ItemData(cboComite.ListIndex) <> 8 Then '35795 mchaves
'                    cargaDatosBLL.RegistroBitacoraCambioComite(clsV2Preanalisis.expedienteSeleccionado, glogon.Usuario, cboComite.ItemData(cboComite.ListIndex), "ELEVADO AL COMITÉ: " + Convert.ToString(cboComite.ItemData(cboComite.ListIndex)) + "-" + cboComite.Text + " Fecha envío: " + Convert.ToString(Date.Now))
End If



If cboComite.ItemData(cboComite.ListIndex) = 7 Or cboComite.ItemData(cboComite.ListIndex) = 8 Then
'    btnGestionarEstudio.Enabled = True
'    btnGestionarSolicitud.Enabled = False
    
    btnGestion(2).Enabled = True 'Gestion Estudio
    btnGestion(3).Enabled = False 'Gestion Solicitud de Credito

Else
'    btnGestionarEstudio.Enabled = False
'    btnGestionarSolicitud.Enabled = False
    
    btnGestion(2).Enabled = False 'Gestion Estudio
    btnGestion(3).Enabled = False 'Gestion Solicitud de Credito
    
    btnComiteCambio.Enabled = False
End If

'
'strSQL = "exec spCrd_Prea_Expediente_Comite_Cambio '" & txtExpediente.Text & "','" & glogon.Usuario & "'," & cboComite.ItemData(cboComite.ListIndex)
'Call ConectionExecute(strSQL)

If glogon.error Then
    MsgBox "No fue posible realizar el cambio de estado, verifique!", vbExclamation
End If

'Actualiza Expediente
Call txtExpediente_LostFocus

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnCopiar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If txtExpediente.Text = "" Then Exit Sub

On Error GoTo vError


strSQL = "exec spCrd_Prea_Expediente_Copia '" & txtExpediente.Text & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If glogon.error Then
    MsgBox "No fue posible realizar la copia del Expediente, verifique!", vbExclamation
    Exit Sub
End If

txtExpediente.Text = rs!Expediente

rs.Close

'Actualiza Expediente
Call txtExpediente_LostFocus

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnCreditos_Click(Index As Integer)


Dim strSQL As String, pTipo As String

On Error GoTo vError

If Index = 0 Then
    pTipo = "C"
Else
    pTipo = "A"
End If

strSQL = "spCrdPreaEliminarCreditosCuotasCxC '" & txtExpediente.Text & "', '" & pTipo & "'"
Call ConectionExecute(strSQL)

Call sbCreditosTransito_Load(pTipo)


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub

Private Sub btnCuenta_Click()
Dim strSQL As String

On Error GoTo vError


GLOBALES.gTag = Trim(txtIdentificación.Text)
GLOBALES.gTag2 = "CRD"

frmCC_Cuentas_Bancarias.Show vbModal

txtIdentificación_LostFocus

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnDeduccion_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spCrdPrea_Deducciones_Add '" & txtExpediente.Text & "', 0, " & cboDeduccion.ItemData(cboDeduccion.ListIndex) _
       & ", '" & txtD_Descripcion.Text & "', " & CCur(txtD_Monto.Text) & ", " & CCur(txtD_Monto.Text) * m_NumPagos _
       & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
    m_CambioCalculo = True
    
    txtD_Monto.Text = "0"
    txtD_Descripcion.Text = ""
    Call cboDeduccion_Click
    Call sbDeducciones_Load
Else
    MsgBox rs!Mensaje, vbExclamation
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbDesembolso_Guardar()

On Error GoTo vError

Me.MousePointer = vbHourglass

'spCrdPreaGuardaDesembolsos](
'                              @CodPreanalisis varchar(30),
'                              @CodAcreedor varchar(10),
'                              @Ordinario smallint,
'                              @Descripcion varchar(60),
'                              @Cuota decimal(16,2),--S.42712 C.8482 APSOLANO
'                              @Monto decimal(16,2),--S.42712 C.8482 APSOLANO
'                              @TipoGiro char(2),
'                              @CedulaDestino char(15),
'                              @TipoCedula smallint,
'                              @CtaIbanDestino varchar(34),
'                              @CodDivisa varchar(10),
'                              @IdBancoDestino varchar(10),
'                              @CorreoNotifica varchar(50),
'                              @Detalle varchar(255),
'                              @GradoHipotec varchar(50)='',
'                              @CodBanco int

Dim vMensaje As String, vTipo As String

vMensaje = ""
vTipo = fxTipoDocumento(cboTipoDocumento.Text)

If cboBanco.ListCount = 0 Then
    vMensaje = vMensaje & "No se ha indicado un Banco para Transferencia" & vbCrLf
End If
If cboCuenta.ListCount = 0 And (vTipo = "TE" Or vTipo = "TS") Then
    vMensaje = vMensaje & "No se ha indicado una Cuenta Bancaria para Transferencia" & vbCrLf
End If

If Len(txtDS_Descripcion.Text) < 10 And (vTipo = "TE" Or vTipo = "TS") Then
    vMensaje = vMensaje & "La Descripción/Beneficiario no es válido!" & vbCrLf
End If

If Len(txtDetalle.Text) < 10 And (vTipo = "TE" Or vTipo = "TS") Then
    vMensaje = vMensaje & "El Detalle de la Transferencia no es válido!" & vbCrLf
End If

If Not fxEmail_Valida(txtCorreo.Text) And (vTipo = "TE" Or vTipo = "TS") Then
    vMensaje = vMensaje & "El Correo Electrónico no es válido!" & vbCrLf
End If


If Len(vMensaje) > 0 Then
    Me.MousePointer = vbDefault
    MsgBox vMensaje, vbExclamation
    Exit Sub
End If



strSQL = "exec spCrdPreaGuardaDesembolsos '" & txtExpediente.Text & "', '" & txtDS_Descripcion.Tag & "', " & IIf(Mid(cboD_Ordinario.Text, 1, 1) = "S", 1, 0) _
       & ", '" & txtDS_Descripcion.Text & "', " & CCur(txtDS_Cuota.Text) & ", " & CCur(txtDS_Monto.Text) & ", '" & vTipo _
       & "', '" & RTrim(txtIdentificación.Text) & "', " & cboTipoId.ItemData(cboTipoId.ListIndex) & ", '" & cboCuenta.ItemData(cboCuenta.ListIndex) _
       & "', '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "', '', '" & txtCorreo.Text & "', '" & txtDetalle.Text & "', '',  " & cboBanco.ItemData(cboBanco.ListIndex)
Call ConectionExecute(strSQL)

Call sbDesembolsos_Nuevo

Me.MousePointer = vbDefault

MsgBox "Desembolso Registrado Satisfactoriamente!", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbDesembolso_Borrar()

On Error GoTo vError

Dim i As Long, vPass As Boolean


Me.MousePointer = vbHourglass

strSQL = ""
vPass = False

With gDesembolsos
    For i = 1 To .MaxRows
       .Row = i
       .Col = 2
       If .Value = vbChecked Then
          .Col = 1
           strSQL = strSQL & Space(10) & "exec spCrdPreaEliminarDesembolsos '" & txtExpediente.Text & "', " & .Text
           vPass = True
       End If
    
    Next i
    
If vPass Then
        Call ConectionExecute(strSQL)
        
        Me.MousePointer = vbDefault
        MsgBox "Desembolsos Eliminados Satifactoriamente!", vbInformation
End If

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnDesembolso_Click(Index As Integer)


If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
    Exit Sub
End If
    
Select Case Index
    Case 0 'Guardar
        Call sbDesembolso_Guardar
        
    Case 2 'Eliminar Sel
        Call sbDesembolso_Borrar
End Select

m_CambioDatos = True

Call sbDesembolsos_Load


End Sub

Private Sub btnEdad_Click()

If txtExpediente.Text = "" Then Exit Sub

GLOBALES.gTag = txtExpediente.Text
GLOBALES.gCedulaActual = txtCedula.Text

frmPrea_Edad_Justificacion.Show vbModal

Call sbEdad_Verifica

End Sub

Private Sub btnEtiqueta_Click()
If Len(txtEtiqueta_Nota.Text) < 50 Then
  MsgBox "Indique una Observación válida, tiene que ser de almenos 50 caracteres!", vbExclamation
  Exit Sub
End If

On Error GoTo vError

strSQL = "exec spCrdPreaAgregaEtiqueta '" & txtExpediente.Text & "', '" & cboEtiquetas.ItemData(cboEtiquetas.ListIndex) _
       & "', '" & txtEtiqueta_Nota.Text & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

MsgBox "Observación aplicada satisfactoriamente!", vbInformation

txtEtiqueta_Nota.Text = ""

Call sbHistorial_Load("E")

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnFianzas_Actualiza_Click()
    If (MsgBox("Está seguro que desea actualizar las fianzas, deberá volver a selecionar los créditos que desea aplicar", vbQuestion + vbYesNo)) = vbYes Then
            
            m_CambioCalculo = True
    
            'Actualizar Fianzas
            glogon.strSQL = "spCRDPreaFianzas '" & txtExpediente.Text & "', 'I'"
            
            If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
                MsgBox "Ocurrió un error al inicializar fianzas.", vbInformation, gMsgTitulo
            End If
    
            Call sbFianzas_Load

    End If
    
End Sub

Private Sub btnGestion_Click(Index As Integer)

If Trim(txtExpediente.Text) = "" Then Exit Sub

'Carga Variables Globales para Otras Ventanas
GLOBALES.gTag = txtExpediente.Text
GLOBALES.gCedulaActual = txtCedula.Text

Call sbTabChange(tcMain.SelectedItem)

Select Case Index

    Case 0 'Causas
        If Len(txtExpediente) > 0 Then
            frmPreaSeguimientoCausas.mCod_linea = Trim(txtLinea)
            frmPreaSeguimientoCausas.Show
       End If
       
    Case 1 'Tags
        If Len(txtExpediente) > 0 Then
            frmPrea_SeguimientoEtiquetas.mId_Solicitud = Trim(txtAsignado)
            frmPrea_SeguimientoEtiquetas.Show
        End If


    Case 2 'gestion
            If fxExistenFiadores Then
                
                frmPreaEstadoPreanalisis.m_estadoPreanalisis = m_estadoPreanalisis
                
                gPreAnalisis.Expediente = txtExpediente.Text
                
                frmPreaEstadoPreanalisis.Show vbModal
                
                
                m_estadoPreanalisis = frmPreaEstadoPreanalisis.m_estadoPreanalisis
                
                tcMain.Item(0).Selected = True
                
                Select Case m_estadoPreanalisis
                Case "P"
                    lblEstado.Caption = "Pendiente"
                Case "R"
                    lblEstado.Caption = "Recibido"
                Case "A"
                    lblEstado.Caption = "Aprobado"
                Case "D"
                    lblEstado.Caption = "Denegado"
                Case "B"
                    lblEstado.Caption = "Abandonado"
                    
                End Select

            Else
            
                tcMain.Item(0).Selected = True
            End If
                      
    Case 3 'Solicitud
            If txtAsignado.Text = "0" Or txtAsignado.Text = "" Then
                If fxExistenFiadores Then
                    gPreAnalisis.Expediente = txtExpediente.Text
                    gPreAnalisis.Tag1 = txtCedula
                    frmPreaSubCredito.Show vbModal
                    tcMain.Item(0).Selected = True
                    
                    txtAsignado = Trim(frmPreaSubCredito.m_Id_Solicitud)
                Else
                    tcMain.Item(0).Selected = True
                End If
            End If

End Select

Me.MousePointer = vbDefault

End Sub


Private Sub btnHipoteca_Click(Index As Integer)
Dim Modulo_Hipotecario As clsHipotecario


On Error GoTo vError

GLOBALES.gTag = txtExpediente.Text

Select Case Index
    Case 0 'Montos Hipoteca
         'Call sbFormsCall("frmPrea_HipotecaMonto", vbModal, , , False, Me, True)
    
         frmPrea_HipotecaMonto.Show vbModal
         
         Call txtExpediente_LostFocus
        
        If txtLinea.Text <> "CRRM" Then 'Sol 53487 MChaves
           Call sbCalcularCuota("txtMonto")
           
           m_CambioCalculo = True
           Call fxGuardar
        End If
     
    Case 1 'Avaluos CFIA
      
      strSQL = "exec spCrdPreaSumarAvaluoCFIA '" & txtExpediente.Text & "', '" & glogon.Usuario & "'"
      Call ConectionExecute(strSQL)
     
      Call txtExpediente_LostFocus
     
     If txtLinea.Text <> "CRRM" Then 'Sol 53487 MChaves
        Call sbCalcularCuota("txtMonto")
        
        m_CambioCalculo = True
        Call fxGuardar
     End If

            
    Case 2 'Garantia
                    Set Modulo_Hipotecario = New clsHipotecario
                    Set Modulo_Hipotecario.vCon = glogon.Conection
                    
                    Modulo_Hipotecario.xOperacion = 0
                    Modulo_Hipotecario.xKey = glogon.ConectRPT
'                    Modulo_Hipotecario.xToolBar = gToolBar
                    
                    Modulo_Hipotecario.xPreaCodigo = txtExpediente.Text
                    Modulo_Hipotecario.xPreaTipo = "E"
                    Modulo_Hipotecario.vCedula = txtCedula.Text
                    
                    Call Modulo_Hipotecario.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                        , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                        , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
        
        
    Case 3 'Asigna Ingenieros
                    Set Modulo_Hipotecario = New clsHipotecario
                    Set Modulo_Hipotecario.vCon = glogon.Conection
                    
                    Modulo_Hipotecario.xOperacion = 0
                    Modulo_Hipotecario.xKey = glogon.ConectRPT
'                    Modulo_Hipotecario.xToolBar = gToolBar
                    
                    Modulo_Hipotecario.xPreaCodigo = txtExpediente.Text
                    Modulo_Hipotecario.xPreaTipo = "E"
                    Modulo_Hipotecario.vCedula = txtCedula.Text
                    
                    Call Modulo_Hipotecario.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                        , App.Path, glogon.ConectRPT, 3, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                        , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
    
    Case 4 'Cambio de Estado
        strSQL = "Select dbo.fxValidaAsignacionComite('" & txtExpediente.Text & "') as Estado"
        Call OpenRecordSet(rs, strSQL)
        If rs!Estado = 0 Then
                MsgBox "Debe seleccionar un comité para poder continuar, favor validar.", vbExclamation, "Asignar Comité."
                Exit Sub
        End If
        
'                If validaComite = 0 Then 'Sol.44794 MChaves
'                    mMensajes.fxMensajeInformacion("Debe seleccionar un comité para poder continuar, favor validar.", "Asignar Comité.")
'                    Exit Sub
        

            If cboComite.Text = "" Then 'Sol.44794 MChaves
                MsgBox "Debe seleccionar un comité para poder continuar, favor validar.", vbExclamation, "Asignar Comité."
                Exit Sub
            End If


            strSQL = "exec spCrdPrea_Comite_Asigna_Valida '" & txtExpediente.Text & "', " & cboComite.ItemData(cboComite.ListIndex)
            Call OpenRecordSet(rs, strSQL)
            If Len(rs!Mensaje) > 0 Then
                MsgBox rs!Mensaje, vbExclamation
                Exit Sub
            End If


            If cboGarantia.ItemData(cboGarantia.ListIndex) = "H" Then

                strSQL = "exec spCRDPreaEstadoHipotecarioAprob '" & txtExpediente.Text & "', '" & glogon.Usuario & "', 0, ''"
                
                Call OpenRecordSet(rs, strSQL)
                If Len(rs!Mensaje) > 0 Then
                    MsgBox rs!Mensaje, vbExclamation
                    Exit Sub
                End If

                Call txtExpediente_LostFocus

            End If

End Select

Exit Sub

vError:

End Sub



Private Sub btnIntereses_Click()
GLOBALES.gTag = txtExpediente.Text

'Call sbFormsCall("frmPrea_FechaFormaliza", vbModal, , , False, Me, True)

frmPrea_FechaFormaliza.Show vbModal

Call txtExpediente_LostFocus

End Sub

Private Sub btnNotificacion_Click()

GLOBALES.gTag = txtExpediente.Text
GLOBALES.gCedulaActual = txtCedula.Text

If lblEstado.Tag <> "R" Then
    Call sbFormsCall("frmPrea_Notificacion", vbModal, , , False, Me, True)
End If

End Sub

Private Sub btnOficinaCambia_Click()
Dim i As Integer

On Error GoTo vError

If txtExpediente.Text = "" Then
    MsgBox "Consulte un Expediente!", vbExclamation
    Exit Sub
End If


If lblEstado.Tag = "A" Then
    MsgBox "No se puede cambiar la Oficina de un expediente que ya ha sido APROBADO. ", vbExclamation
    Exit Sub
End If

If fxSelectItemSubExpediente(cboSubExpediente.ItemData(cboSubExpediente.ListIndex)) = "S" Then
    MsgBox "No se puede CAMBIAR la Oficina de un expediente secundario, por favor seleccione el expediente principal e intente de nuevo. ", vbExclamation
    Exit Sub
End If

Dim pPromotor As String

pPromotor = "Null"
If IsNumeric(txtEjecutivo.Tag) Then
    pPromotor = txtEjecutivo.Tag
End If

i = MsgBox("¿Desea Cambiar la Oficina y el Ejecutivo Colocador del estudio seleccionado?", vbYesNo)
If i = vbYes Then
    strSQL = "exec spCrdPreaAsignaOficina '" & txtExpediente.Text & "', '" & txtOficina.Tag & "', " & pPromotor
    Call ConectionExecute(strSQL)
    
    MsgBox "Se ha actualizado la Oficina y el Ejecutivo del expediente correctamente.", vbInformation
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnP_Examenes_Click()
'Select Case Mid(cboEstadoExamenes.Text, 1, 1)
'                    Case "E"
'                        nota = "Exámenes médicos enviados"
'                    Case "R"
'                        nota = "Exámenes médicos recibidos"
'                    Case "A"
'                        nota = "Exámenes médicos aprobados"
'                    Case Else
'                        nota = ""
'                End Select
                
'spCRD_PreaAplicaEstadoExamenes


'SELECT NOTA, USUARIO, FECHA FROM CRD_PREA_HISTORIALEXAMENES " _
'& "WHERE COD_PREANALISIS = @EXPEDIENTE ORDER BY ID_NOTA


End Sub

Private Sub btnPrenda_Click(Index As Integer)

GLOBALES.gTag = txtExpediente.Text
GLOBALES.gCedulaActual = txtCedula.Text

Select Case Index
    Case 0 'Prenda
    
        Operacion.GarantiaTipo = "P" 'Prenda
        Operacion.GarantiaId = 0
        
        Operacion.Expendiente = txtExpediente.Text
        Operacion.GarantiaParam = "E" 'Estudio
        
        Operacion.Operacion = 0
        Operacion.Cedula = Trim(txtCedula.Text)
    
        'Call sbFormsCall("frmCR_Prendas", vbModal, , , False, Me, True)
    
        frmCR_Prendas.Show vbModal
    
    Case 1 'Gastos
        ' Call sbFormsCall("frmPrea_PrendaMonto", vbModal, , , False, Me, True)
        
        
        frmPrea_PrendaMonto.Show vbModal
        
         Call txtExpediente_LostFocus
        
        If txtLinea.Text <> "CRRM" Then 'Sol 53487 MChaves
           Call sbCalcularCuota("txtMonto")
           
           m_CambioCalculo = True
           Call fxGuardar
        End If

End Select

End Sub

Private Sub btnRefundiciones_Actualiza_Click()
Dim strSQL As String

On Error GoTo vError
    
If (MsgBox("Está seguro que desea actualizar las refundiciones, deberá volver a selecionar los créditos que desea refundir", vbQuestion + vbYesNo)) = vbYes Then
          
    m_CambioCalculo = True
                
    
    Me.MousePointer = vbHourglass
    
    strSQL = "exec spCrdPreaRefundicionesActualiza '" & gPreAnalisis.Expediente & "'"
    Call ConectionExecute(strSQL)
    
    Me.MousePointer = vbDefault

    MsgBox "Estado de las Operaciones a Refinanciar o Abonar actualizado!", vbInformation

    Call sbRefundiciones_Load

 
'            If fxValidaEstado(gPreAnalisis.Expediente) = True Then
'                glogon.strSQL = "exec spCRDPreaRefundiciones '" & gPreAnalisis.Expediente & "','A'"
'                If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
'                    MsgBox "Ocurrió un error al inicializar fianzas.", vbInformation, gMsgTitulo
'                End If
'            End If
'
'            Call sbClasificacion_CargaGrid
End If 'Yes/No

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub btnResumen_Click(Index As Integer)
 
GLOBALES.gTag = txtExpediente.Text
 
 Select Case Index
    Case 0 'Reportes
        Call btnBarra_Click(5)
        
    Case 1 'Deducciones
        GLOBALES.gCedulaActual = txtCedula.Text
        Call sbFormsCall("frmCR_EnCobroCuotas", , , , False, Me, True)
        
    Case 2 'Ballon
    
     If txtLinea.Text <> "CRRM" Then
        MsgBox "Esta Linea no Está autorizada para Balloon Payment!", vbExclamation
        Exit Sub
     End If
        'Call sbFormsCall("frmPrea_BallonPayment", vbModal, , , False, Me, True)
 
      frmPrea_BallonPayment.Show vbModal
 
 End Select
 
End Sub



Private Sub sbSalarios_Guardar()
Dim pSQL_Sistema As String, pSQL_Otros As String
Dim i As Integer

On Error GoTo vError

If txtExpediente.Text = "" Or Not gbSalarios(0).Enabled Then
   Exit Sub
End If


Dim pSalario As Currency, pOrden As Integer, pFecha As String, pCA As Currency, pMesRH As Integer
Dim pSalarioRH As Currency

Me.MousePointer = vbDefault


pSQL_Sistema = ""
pSQL_Otros = ""

strSQL = "exec spCrdPreaEliminarSalarios '" & txtExpediente.Text & "'"
Call ConectionExecute(strSQL)

With gSalarios
 For i = 1 To .MaxRows
    .Row = i
    .Col = 1
    pFecha = .Text
    .Col = 2
    pSalario = .Text
    .Col = 3
    pMesRH = .Text
    .Col = 4
    pSalarioRH = .Text
    .Col = 5
    pCA = .Text
    
    pOrden = i
    
    pSQL_Sistema = pSQL_Sistema & Space(10) & "exec spCrdPreaGeneraSalariosSistema '" & txtExpediente.Text & "', " & pSalario & ", '" & pFecha & "', " _
                & pOrden & ", " & pCA & ", " & pMesRH
    
    pSQL_Otros = pSQL_Otros & Space(10) & "exec spCrdPreaGeneraSalariosConsultaLinea '" & txtExpediente.Text & "', " & pSalario & ", '" & pFecha & "', " _
                & pOrden & ", " & pCA & ", " & pMesRH

 
 Next i
End With

If Len(pSQL_Sistema) > 0 Then
   Call ConectionExecute(pSQL_Sistema)
End If

If Len(pSQL_Otros) > 0 Then
   Call ConectionExecute(pSQL_Otros)
End If


Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbIncapacidades_Guardar()
Dim i As Integer

On Error GoTo vError

Dim pDias As Integer, pDesde As String, pHasta As String


Me.MousePointer = vbDefault
                                               
strSQL = "exec spCrdPreaEliminarIncapacidades '" & txtExpediente.Text & "'"
Call ConectionExecute(strSQL)

strSQL = ""

With lswIncapacidades.ListItems
 For i = 1 To .Count
  
    pDesde = .Item(i).Text
    pHasta = .Item(i).SubItems(1)
    pDias = .Item(i).SubItems(2)
    
    strSQL = strSQL & Space(10) & "exec spCrdPreaGeneraIncapacidades '" & txtExpediente.Text & "', " & pDias & ", '" & pDesde & "', '" & pHasta & "', " & i
 
 Next i
End With

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnS_Copy_Click(Index As Integer)

If gSalarios.MaxRows = 0 Or txtExpediente.Text = "" Then
    Exit Sub
End If

If Not m_Editable Then Exit Sub

Dim TextoClipboard As String
Dim Filas() As String
Dim Columnas() As String
Dim i As Integer, j As Integer

Dim Col As Integer

Dim Valor As String


On Error GoTo vError

Select Case Index
    Case 0 'Mes
      Col = 3
    Case 1 'Salario RH
      Col = 4
    Case 2 'CA
      Col = 5
    Case 3 'Incapacidades
      Col = 6
    Case 4 'Elimina Incapacidades
      Col = 6
      
      strSQL = "exec spCrdPreaEliminarIncapacidades '" & txtExpediente.Text & "'"
      Call ConectionExecute(strSQL)
      
      lswIncapacidades.ListItems.Clear
      
      Exit Sub
End Select


' Verificar si hay datos en el portapapeles
If Clipboard.GetFormat(vbCFText) = False Then
    MsgBox "No hay datos de texto en el portapapeles.", vbExclamation, "Aviso"
    Exit Sub
End If

' Obtener el texto del portapapeles
TextoClipboard = Clipboard.GetText
If Trim(TextoClipboard) = "" Then
    MsgBox "El portapapeles está vacío.", vbExclamation, "Aviso"
    Exit Sub
End If


' Dividir en filas (cada línea representa una fila copiada de Excel)
Filas = Split(TextoClipboard, vbCrLf)

' Limpiar el grid antes de cargar datos
'gSalarios.ClearRange 0, 0, fpsGrid.MaxRows, fpsGrid.MaxCols, True
If Col < 6 Then
    'Salarios
    With gSalarios
        ' Recorrer filas y columnas
        For i = 0 To UBound(Filas)
            If Trim(Filas(i)) <> "" Then
                Columnas = Split(Filas(i), vbTab)  ' Separar columnas (Excel usa TAB)
                Valor = Columnas(0)
                
                If .MaxRows >= i + 1 Then
                   .Row = i + 1
                   .Col = Col
                    If IsNumeric(Valor) Then
                        .Text = Format(Valor, "Standard")
                    Else
                        .Text = Format(0, "Standard")
                    End If
                End If
                
                'Registra Datos
                'TODO:
                
            End If
        Next i
    
    End With
   
    'Guardar Salarios
    'sbSalarios_Guardar
End If 'Col < 6

If Col = 6 Then
    'Incapacidades
    With lswIncapacidades
        For i = 0 To UBound(Filas)
            If Trim(Filas(i)) <> "" Then
                Columnas = Split(Filas(i), vbTab) ' Separar columnas (Excel usa TAB)
                        
                Set itmX = .ListItems.Add(, , Columnas(0))
                    itmX.SubItems(1) = Columnas(1)
                    itmX.SubItems(2) = Columnas(2)
                    
            End If
        Next i
    
       If .ListItems.Count > 0 Then
            Call sbIncapacidades_Guardar
       End If
    
    End With

End If
    

Exit Sub


vError:
    MsgBox "Error en la carga del Portapapeles al Sistema!", vbExclamation


End Sub

Private Sub btnSolicitado_Click()
Dim strSQL As String

If txtExpediente.Text = "" Then Exit Sub

On Error GoTo vError


strSQL = "exec spCrd_Prea_Estado_Solicitado '" & txtExpediente.Text & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

If glogon.error Then
    MsgBox "No fue posible realizar el cambio de estado, verifique!", vbExclamation
End If

'Actualiza Expediente
Call txtExpediente_LostFocus

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

Private Sub cboCantidadFiadores_Change()
    m_CambioDatos = True
End Sub

Private Sub cboCantidadFiadores_Click()
m_CambioDatos = True
If cboSubExpediente.ListCount = 1 Then Exit Sub
If m_DesplegoMensaje Then Exit Sub

m_DesplegoMensaje = True
If m_FiadoresRegistrador > cboCantidadFiadores.Text Then
    MsgBox "No es posible disminuir la cantidad de sub expedientes."
    cboCantidadFiadores.Text = m_FiadoresRegistrador
End If
m_DesplegoMensaje = False
End Sub
     
Private Sub cboComite_Change()
        m_CambioDatos = True
End Sub

Private Sub cboComite_Click()
    m_CambioDatos = True
End Sub

Private Sub cboD_Ordinario_Click()
If vPaso Then Exit Sub
If txtExpediente.Text = "" Then Exit Sub

Call sbDesembolsos_Externos_Lista

End Sub

Private Sub cboDeduccion_Click()
If vPaso Then Exit Sub

txtD_Descripcion.Text = cboDeduccion.Text
txtD_Descripcion.Locked = True

txtD_Monto.Text = Format(0, "Standard")

strSQL = "select EDITAR_DESCRIPCION, FRAP" _
       & "  From CRD_PREA_V2_DEDUCCIONES_CONFIG" _
       & " Where ID_DEDUCCION = " & cboDeduccion.ItemData(cboDeduccion.ListIndex)
Call OpenRecordSet(rs, strSQL)

If rs!EDITAR_DESCRIPCION = 1 Then
    txtD_Descripcion.Locked = False
End If



End Sub

Private Sub cboDestino_Click()
   m_CambioDatos = True
   Call sbAplicaPrimeraCta
End Sub
Private Sub sbAplicaPrimeraCta()

On Error GoTo vError

chkPrimerCuota.Value = 0

If (cboDestino.ListCount = 0) Or (cboDestino.ItemData(cboDestino.ListIndex) = "") Then Exit Sub
If Len(txtDesLineaCredito.Text) = 0 Then Exit Sub

clsEntidad.tablaName = "spCRDPreaDestinos"
If clsEntidad.fxTraerFiltrado("AplicaPrimCta", "'" & Trim(txtLinea.Text) & "','" & cboDestino.ItemData(cboDestino.ListIndex) & "'") Then
    chkPrimerCuota.Value = glogon.Recordset.Fields!PRIMER_CUOTA
Else
    chkPrimerCuota.Value = 0
End If

 
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Private Sub cboDestino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboGarantia.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "D.cod_destino"
   gBusquedas.Orden = "D.Cod_Destino"
   gBusquedas.Consulta = "select D.cod_Destino,D.descripcion" _
                        & " from catalogo_destinos D inner join catalogo_destinosASG C on D.cod_destino = C.cod_destino"
   gBusquedas.Filtro = " and C.codigo = '" & txtLinea.Text & "' "
   frmBusquedas.Show vbModal
   cboDestino.Text = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub cboFondo_Click()
    Dim strSQL As String, rs As New ADODB.Recordset
    
    If vPasoCarga Then Exit Sub
    If cboFondo.ListCount <= 0 Then Exit Sub
    If cboFondo.Text = "" Then Exit Sub
    
    If cboGarantia.ItemData(cboGarantia.ListIndex) <> "Y" Then Exit Sub
    
    strSQL = "exec spCRDGarantiaFNDCalculo '" & Trim(txtCedula.Text) & "','" & cboFondo.ItemData(cboFondo.ListIndex) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    txtMonto.Text = Format(rs!Disponible, "Standard")
    If rs!AplicaTasa = 1 Then
        txtTasa.Text = rs!TASA
    End If
    
    If rs!AplicaPlazo = 1 Then
        txtPlazo.Text = rs!Plazo
    End If
    
    LblTasa.Tag = rs!AplicaTasa
    LblPlazo.Tag = rs!AplicaPlazo
    
    rs.Close
    
    m_CambioDatos = True
    
    

strSQL = "select cod_contrato,Tasa_Referencia,Aportes, isnull(FECHA_CORTE, getdate()) as 'FECHA_CORTE'" _
       & " from fnd_contratos" _
       & " where cod_plan = '" & cboFondo.ItemData(cboFondo.ListIndex) _
       & "' and estado = 'A' and cedula = '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)
cboFondoContrato.Clear
Do While Not rs.EOF
  cboFondoContrato.AddItem "[Cnt: " & rs!COD_CONTRATO & "] [Tasa: " & rs!TASA_REFERENCIA & "] [I: " & Format(rs!Aportes, "Standard") _
        & "] [V: " & Format(rs!Fecha_Corte, "yyyy-mm-dd") & "]"
  cboFondoContrato.ItemData(cboFondoContrato.ListCount - 1) = CStr(rs!COD_CONTRATO)

  rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboFondoContrato.Text = "[Cnt: " & rs!COD_CONTRATO & "] [Tasa: " & rs!TASA_REFERENCIA & "] [I: " & Format(rs!Aportes, "Standard") _
        & "] [V: " & Format(rs!Fecha_Corte, "yyyy-mm-dd") & "]"
End If
rs.Close



'vPaso = False
'
'If Not vOperacionLoad Then
'    If cboFondoContrato.ListCount <= 0 Then
'         strSQL = "exec spCRDGarantiaFNDCalculo '" & txtCedula & "','" & cboFondo.ItemData(cboFondo.ListIndex) & "',0"
'    Else
'         strSQL = "exec spCRDGarantiaFNDCalculo '" & txtCedula & "','" & cboFondo.ItemData(cboFondo.ListIndex) _
'                & "'," & cboFondoContrato.ItemData(cboFondoContrato.ListIndex)
'    End If
'
'    Call OpenRecordSet(rs, strSQL)
'
'    txtMonto.Text = Format(rs!Disponible, "Standard")
'    If rs!AplicaTasa = 1 Then
'        txtTasa.Text = Format(rs!TASA, "Standard")
'    End If
'
'    If rs!AplicaPlazo = 1 Then
'        txtPlazo.Text = rs!Plazo
'    End If
'
'    LblTasa.Tag = rs!AplicaTasa
'    LblPlazo.Tag = rs!AplicaPlazo
'
'    rs.Close
'End If 'Operacion load
    

End Sub

Private Sub cboFondoContrato_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPasoCarga Then Exit Sub
If cboFondo.ListCount <= 0 Then Exit Sub
If cboFondo.Text = "" Then Exit Sub

If cboGarantia.ItemData(cboGarantia.ListIndex) <> "Y" Then Exit Sub

If cboFondoContrato.ListCount <= 0 Then
     strSQL = "exec spCRDGarantiaFNDCalculo '" & txtCedula & "','" & cboFondo.ItemData(cboFondo.ListIndex) & "',0"
Else
     strSQL = "exec spCRDGarantiaFNDCalculo '" & txtCedula & "','" & cboFondo.ItemData(cboFondo.ListIndex) _
            & "'," & cboFondoContrato.ItemData(cboFondoContrato.ListIndex)
End If

Call OpenRecordSet(rs, strSQL)

txtMonto.Text = Format(rs!Disponible, "Standard")
If rs!AplicaTasa = 1 Then
    txtTasa.Text = Format(rs!TASA, "Standard")
End If

If rs!AplicaPlazo = 1 Then
    txtPlazo.Text = rs!Plazo
End If

LblTasa.Tag = rs!AplicaTasa
LblPlazo.Tag = rs!AplicaPlazo

rs.Close

End Sub

Private Sub cboGarantia_Change()
    m_CambioDatos = True
End Sub

Private Sub cboGarantia_Click()
Dim strSQL As String, rs As New ADODB.Recordset

Dim pPassFondo As Boolean
 
m_CambioDatos = True


pPassFondo = False

chkPolizaVida.Value = vbUnchecked
chkPolizaVida.Enabled = True

cboFondo.Enabled = False
cboFondoContrato.Enabled = False

cboCantidadFiadores.Enabled = False

Dim pGarantia As String, pGarantiaForm As String

pGarantia = cboGarantia.ItemData(cboGarantia.ListIndex)

strSQL = "select FORMULARIO  From CRD_GARANTIA_TIPOS" _
       & " where garantia = '" & pGarantia & "'"
Call OpenRecordSet(rs, strSQL)
 pGarantiaForm = Trim(rs!Formulario)
rs.Close

tcMain.Item(9).Visible = False   'Tab Hipotecario
tcMain.Item(10).Visible = False  'Tab Prendario

lblMontoConstruccion.Visible = False
txtMontoConstruccion.Visible = False

chkPolizaIncendio.Enabled = False
chkPolizaVehiculo.Enabled = False


Select Case pGarantiaForm
    
    
    Case "F01" 'Sobre Ahorros
        strSQL = "select dbo.fxCrdGarantiaPatMnt('" & txtCedula.Text & "','A', 'M') as 'Monto'"
        Call OpenRecordSet(rs, strSQL)
          txtMonto.Text = Format(rs!Monto, "Standard")
        rs.Close
    
    Case "F02" 'Fiduciaria
    
        cboCantidadFiadores.Enabled = True
    
    Case "F03" 'Hipotecaria
    
        cboCantidadFiadores.Enabled = True
        
        chkPolizaVida.Value = vbChecked
        Call chkPolizaVida_Click
        
        chkPolizaVida.Enabled = False
        chkPolizaIncendio.Enabled = True
        
        tcMain.Item(9).Visible = True
        tcMain.Item(9).Enabled = True
    
    
        lblMontoConstruccion.Visible = True
        txtMontoConstruccion.Visible = True
        
    Case "F05" 'Fondos de Ahorros
            If vPasoCarga Then Exit Sub
            If cboGarantia.ListCount <= 0 Then Exit Sub
             
            pPassFondo = True
            
            cboFondo.Enabled = True
            cboFondoContrato.Enabled = True
            
            Call cboFondo_Click
            Exit Sub
    
    Case "F06" 'Adelanto de Salario
        strSQL = "select dbo.fxCrdDisponibleAdelantoSalario_Estudio('" & txtCedula.Text & "', 'M') as 'Monto'"
        Call OpenRecordSet(rs, strSQL)
          txtMonto.Text = Format(rs!Monto, "Standard")
        rs.Close
    
    Case "F07" 'Prendaria
        chkPolizaVehiculo.Value = vbChecked
        Call chkPolizaVehiculo_Click
        
        chkPolizaVehiculo.Enabled = False
        
        tcMain.Item(10).Visible = True
        tcMain.Item(10).Enabled = True

End Select


'Corrección
Select Case pGarantia
    Case "Y"
            If vPasoCarga Then Exit Sub
            If cboGarantia.ListCount <= 0 Then Exit Sub
            If pPassFondo Then Exit Sub
             
            cboFondo.Enabled = True
            Call cboFondo_Click
            Exit Sub
    
    Case "F", "H"
      'Nada
    Case Else
        If Val(cboCantidadFiadores.Text) > 0 Then
            m_FiadoresRegistrador = 0
            cboCantidadFiadores.Text = 0
            cboCantidadFiadores.Enabled = False
        End If
End Select



m_curValor_Anterior = 0
If IsNumeric(txtMonto.Text) Then
    If CCur(txtMonto.Text) > 0 Then
      Call sbCalcularCuota("txtMonto")
    End If
End If

End Sub

'---------------------Registro de cliente


Private Function fxBorrarExtras() As Boolean

On Error GoTo vError
Me.MousePointer = vbHourglass
 
fxBorrarExtras = False

'If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
'    GoTo salir
'End If

clsEntidad.tablaName = "spCRDPreaExtrasXPreAnalisis"
If clsEntidad.fxRemover("'" & gPreAnalisis.Expediente & "'") Then
    fxBorrarExtras = True
End If


    Me.MousePointer = vbDefault
    Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
    
    
End Function




Function ExtraerPorcentaje(texto As String) As String
    Dim regex As Object
    Dim match As Object
    
    ' Crear objeto de expresión regular
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\[\s*(\d+)\s*%\s*\]" ' Captura solo el número dentro de los corchetes
    regex.Global = False
    regex.IgnoreCase = True

    ' Ejecutar la búsqueda
    If regex.Test(texto) Then
        Set match = regex.Execute(texto)
        ExtraerPorcentaje = match(0).SubMatches(0) ' Obtener solo el número dentro de los []
    Else
        ExtraerPorcentaje = "" ' Si no encuentra coincidencia, retorna vacío
    End If

    ' Liberar memoria
    Set regex = Nothing
    Set match = Nothing
End Function

Private Sub cboS_ComponenteAdicional_Click()
If vPaso Then Exit Sub

If cboS_ComponenteAdicional.Text = "" Then
   txtS_ComponenteAdicionalPorc.Text = "0"
Else
   txtS_ComponenteAdicionalPorc.Text = ExtraerPorcentaje(cboS_ComponenteAdicional.Text)
End If


End Sub

Private Sub cboSalario_Click()
Dim Item As String
Dim Codigo As String
Dim IndiceTipoSalario As String
Dim sql As String

On Error GoTo vError

    If vPaso Then Exit Sub

    'Solo estos codigos permiten a, b , c y f seleccionar el salario
    
    Me.MousePointer = vbHourglass

    'Codigo = SIFGlobal.fxCodText(cboSalario.Text)
    
    Codigo = cboSalario.ItemData(cboSalario.ListIndex)
    
    Item = Left(Right(cboSalario.Text, 2), 1)
    
    Call sbHabilitaFechaColilla(Trim(Item))

    IndiceTipoSalario = Right(cboSalario.Text, 3)
    
    gExtras.Enabled = dtpCorte.Enabled
    
    'sbCRDPreaSalario(@Expediente varchar(15), @TipoSalario varchar(2) = '')
    sql = "spCRDPreaSalario " & fxFormatearValor(gPreAnalisis.Expediente, caracter) & ","
    sql = sql & fxFormatearValor(Trim(Codigo), caracter)
    
    If m_CargoSalario Then
        m_CargoSalario = False
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        Call sbActCtlConstExternos(True, IndiceTipoSalario)

        Select Case Trim(Item)
        Case "a", "b", "c", "e", "f"
            
            If clsEntidad.fxEjecutaSQL(sql) Then
                If glogon.Recordset.RecordCount = 1 Then
                    txtSalarioDevengado.Text = Format(glogon.Recordset!Salario, "Standard")
                    txtS_Devengado.Text = txtSalarioDevengado.Text
                Else
                    m_SoloVerSalarios = False
                End If
            End If
            
           Call sbSalarios_Registro_Inicial
            
        Case "d" 'Constancia Salarial
            txtSalarioDevengado.Text = 0
            txtS_Devengado.Text = txtSalarioDevengado.Text
        
        Case "g" 'Constancia Externa
            Call sbActCtlConstExternos(False, IndiceTipoSalario)
        
        Case Else
            
            txtSalarioDevengado.Text = 0
            txtS_Devengado.Text = txtSalarioDevengado.Text
        End Select
    
    End If
    
    Call SbBloquearTxtSalario(Trim(Codigo))

    m_CambioCalculo = True
    
    Call sbEstructuraActualiza(1, False)
    
    If Item <> "e" Then
     '   Call fxBorrarExtras
    End If
    
    
    
    Me.MousePointer = vbDefault
Exit Sub
vError:
    Me.MousePointer = vbDefault
    cMensaje.deError ("Ocurrió un error . Error:" & Err.Description)

End Sub


Private Sub sbHabilitaFechaColilla(ByVal Tipo As String)

    If Tipo = "f" Or Tipo = "a" Then
        dtpCorte.Enabled = False
    Else
        dtpCorte.Enabled = True
    End If

End Sub

Private Sub sbActCtlConstExternos(ByVal Activar As Boolean, ByVal IndiceTipoSalario As String)
Dim vkey As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If Activar Then
    txtSalarioLiquido.BackColor = &HE0E0E0
    txtTotal_Cargas_CCSS.BackColor = &HFFFFFF
    txtDeducciones.BackColor = &HFFFFFF
Else
    txtSalarioLiquido.BackColor = &HFFFFFF
    txtSalarioLiquido.Locked = False
    txtTotal_Cargas_CCSS.BackColor = &HE0E0E0
    txtDeducciones.BackColor = &HE0E0E0
End If

 If IndiceTipoSalario = "(g)" Then 'codigo que Corresponde al tipo de salario
    
'    btnDetalle.Item(1).Visible = False
'    btnDetalle.Item(2).Visible = False
    
    txtTotal_Cargas_CCSS.Text = Format(0, "Standard")
    txtDeducciones.Text = Format(0, "Standard")
Else
'    btnDetalle.Item(1).Visible = True
'    btnDetalle.Item(2).Visible = True
End If


Me.MousePointer = vbDefault
Exit Sub

vError:
cMensaje.deError ("Ocurrió un error  activando o deshablitando controles. Error:" & Err.Description)


End Sub


Private Sub cboSexo_Change()
    m_CambioDatos = True
End Sub

Private Sub cboSexo_Click()
    Call sbCalcularPlazoMaximo
    m_CambioDatos = True
End Sub

Private Sub cboSubExpediente_Click()
   
    If m_Cargando Then Exit Sub
    If txtExpediente.Text = cboSubExpediente.Text Then
        Exit Sub
    End If

    Me.MousePointer = vbHourglass

    m_valorComboExp = cboSubExpediente.Text
        
    '' Proceso pregunta si desea guardar los datos.
    If cboSubExpediente.Text = "Nuevo SubExpediente" Or m_valorComboExp = "Nuevo Expediente" Then
        If Trim(txtExpediente.Text) <> "" Then
            If (m_CambioDatos = True) Or (m_CambioCalculo = True) Or (m_CambioObservaciones = True) Then
                If (MsgBox("NO GUARDO los cambios que hizo en el expediente, ¿Desea continuar SIN GUARDAR los cambios?.", vbQuestion + vbYesNo + vbDefaultButton2, gMsgTitulo) = vbNo) Then
                    cboSubExpediente.Text = Trim(txtExpediente.Text)
                    Me.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    End If
    
    DoEvents

    Call sbcboSubExpediente_Validate

    If cboSubExpediente.Text = "Nuevo SubExpediente" Then
        vCodExpediente = txtExpediente.Text
        Me.MousePointer = vbDefault
        If fxValidaNumFiadoresRegistrados(True) = False Then Exit Sub
    End If

    DoEvents

    If Right(cboSalario.Text, 3) = "(g)" Then
        Call sbActCtlConstExternos(False, "(g)")
    Else
        Call sbActCtlConstExternos(True, "(d)")
    End If

    If m_valorComboExp = "Nuevo SubExpediente" Or m_valorComboExp = "Nuevo Expediente" Then
        txtCedula.SetFocus
    End If

    sbActivarMontoGirar
    
    
    Me.MousePointer = vbDefault

End Sub

Private Sub sbActivarMontoGirar()
    If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
        txtMontoGirar.Visible = False
        lblMontoGirar.Item(34).Visible = False
    Else
        txtMontoGirar.Visible = True
        lblMontoGirar.Item(34).Visible = True
    End If
End Sub

Private Sub sbcboSubExpediente_Validate()
Dim vControl As Control
txtExpediente.Locked = True

If m_valorComboExp = "Nuevo SubExpediente" Then
'If fxSelectItemSubExpediente(cboSubExpediente.ItemData(cboSubExpediente.ListIndex)) = "S" Then
   ' m_expediente = txtExpediente.Text
    For Each vControl In Me
'        MsgBox TypeName(vControl) & "  :" & vControl.Name
        Select Case TypeName(vControl)
        Case "TextBox", "FlatEdit"
            If Not ((vControl.Name = "txtExpediente") _
                Or (vControl.Name = "txtLinea") _
                Or (vControl.Name = "txtDesLineaCredito") _
                Or (vControl.Name = "txtPolizaVida") _
                Or (vControl.Name = "txtPolizaIncendio") _
                Or (vControl.Name = "txtPolizaDesempleo") _
                Or (vControl.Name = "txtPolizaPrenda") _
                Or (vControl.Name = "txtMonto") _
                Or (vControl.Name = "txtPlazo") _
                Or (vControl.Name = "txtCuota") _
                Or (vControl.Name = "txtCompromiso") _
                Or (vControl.Name = "txtTasa")) _
                Then
                
                    vControl.Text = ""
            End If
            txtExpediente.Locked = False
            
        Case "CheckBox"
            If Not ((vControl.Name = "chkPolizaVida") _
                Or (vControl.Name = "chkPolizaVehiculo") _
                Or (vControl.Name = "chkPolizaIncendio") _
                Or (vControl.Name = "chkPolizaDesempleo")) _
                Or (vControl.Name = "chkPrimerCuota") _
            Then
                'vControl.Value = 0
                If vControl.Name <> "chkCargaFrap" Then
                    vControl.Enabled = False
                End If
                
            End If
        End Select
    Next vControl
    
    chkCargaAsociacion.Enabled = True
    Call sbAccionVentana(NuevoRegistro)
    Call sbBloquearControles(Me, SubExpediente)
    
    chkCargaAsociacion.Value = vbUnchecked
    chkCargaFrap.Value = vbUnchecked

ElseIf m_valorComboExp = "Nuevo Expediente" Then
    Call sbLimipiaControles(Me, True)
    
    dtpCorte.Value = fxFechaServidor
    
    If cboSubExpediente.ListCount > 1 Then
        Call sbInicializaComboExpediente
    End If
    
    Call sbAccionVentana(NuevoRegistro)
    Call sbBloquearControles(Me, Expediente)
    txtExpediente.Locked = False

ElseIf m_valorComboExp = "Nuevo Expediente" Or m_valorComboExp = "Nuevo SubExpediente" Then
    Else
        txtExpediente.Text = m_valorComboExp
        m_CargoCombo = True
        Call txtExpediente_LostFocus
End If



End Sub

Private Sub sbCalcularCompromiso()

    txtCompromiso.Text = ""
    If Not IsNumeric(txtCuota.Text) Then
        txtCuota.Text = 0
    End If
    
    If txtPolizaVida = "" Then
        txtPolizaVida = 0
    End If
    
    If txtPolizaIncendio = "" Then
        txtPolizaIncendio = 0
    End If
    
    If txtPolizaDesempleo = "" Then
        txtPolizaDesempleo = 0
    End If
    
    If txtPolizaPrenda = "" Then
        txtPolizaPrenda = 0
    End If
    
    
    If (Len(txtCuota.Text) > 0 And Len(txtPolizaVida.Text) > 0 And Len(txtPolizaIncendio.Text) > 0 _
        And Len(txtPolizaDesempleo.Text) > 0 And Len(txtPolizaPrenda.Text) > 0) Then
        txtCompromiso.Text = Format((CDbl(txtCuota.Text) _
                + CDbl(txtPolizaVida.Text) _
                + CDbl(txtPolizaIncendio.Text) _
                + CDbl(txtPolizaDesempleo.Text) _
                + CDbl(txtPolizaPrenda.Text) _
                ), "Standard")
    End If


End Sub

Private Sub sbCalcularCuota(ByVal Control As String)
Dim mBono As Double, mPlazo As Integer

On Error GoTo vError
 
m_CambioDatos = True

Select Case Control
    Case "txtMonto"
    
       If Val(txtMonto.Text) > 0 Then
       
        If cboGarantia.ItemData(cboGarantia.ListIndex) = "Y" And cboFondo.ListCount > 0 Then
            If LblPlazo.Tag = "0" Then txtPlazo.Text = fxCatalogoRango(Trim(txtLinea.Text), Format(txtMonto.Text, "Standard"), "P", cboDestino.ItemData(cboDestino.ListIndex), cboGarantia.ItemData(cboGarantia.ListIndex))
            If LblTasa.Tag = "0" Then txtTasa.Text = fxCatalogoRango(txtLinea.Text, txtMonto.Text, "I", cboDestino.ItemData(cboDestino.ListIndex), cboGarantia.ItemData(cboGarantia.ListIndex))
        Else
            If Len(txtDesLineaCredito.Text) > 0 Then
                txtPlazo.Text = fxCatalogoRango(txtLinea.Text, CDbl(txtMonto.Text), "P", cboDestino.ItemData(cboDestino.ListIndex), cboGarantia.ItemData(cboGarantia.ListIndex))
                txtTasa.Text = fxCatalogoRango(txtLinea.Text, CDbl(txtMonto.Text), "I", cboDestino.ItemData(cboDestino.ListIndex), cboGarantia.ItemData(cboGarantia.ListIndex))
            End If
        End If
        
        ''Modifica tasa para aplicar bonos por membresia
        If clsMensajes.Estado = "R" Or clsMensajes.Estado = "P" Then
            mBono = fxBonoMembresia(Trim(txtCedula.Text), txtLinea.Text, cboGarantia.ItemData(cboGarantia.ListIndex), cboDestino.ItemData(cboDestino.ListIndex), txtPlazo.Text)
            mPlazo = fxBonoPlazoMembresia(Trim(txtCedula.Text), cboGarantia.ItemData(cboGarantia.ListIndex))
            If mBono > 0 Then
                txtTasa.Text = CDbl(txtTasa.Text) - mBono
                clsMensajes.TASA_PTS_BONO = mBono
            Else
                txtTasa.ToolTipText = Empty
                txtTasa.Tag = 0
                clsMensajes.TASA_PTS_BONO = 0
            End If
        
            If mPlazo > 0 Then
               txtPlazo.Text = mPlazo
            End If
        
        End If
        
        If clsMensajes.TASA_PTS_BONO > 0 Then
            txtTasa.ToolTipText = "Bono por Membresia de " & clsMensajes.TASA_PTS_BONO
        Else
            txtTasa.ToolTipText = Empty
        End If
        
        
        If (Val(txtPlazo.Text) > 0 And Val(txtTasa.Text) >= 0) And IsNumeric(txtMonto.Text) Then
                txtCuota.Text = Format(fxCalcula_Cuota(CDbl(txtMonto.Text), Val(txtPlazo.Text), CDbl(txtTasa.Text), mFrecuenciaPago), "Standard")
            Else
                txtCuota.Text = 0
            End If
        Else
            txtPlazo.Text = 0
            txtTasa.Text = 0
            txtCuota.Text = 0
       End If
       
    Case "txtPlazo", "txtTasa"
   
        If (Val(txtMonto.Text) > 0 And Val(txtPlazo.Text) > 0) Then 'And Val(txtTasa.Text) > 0) Then
            
            txtCuota.Text = Format(fxCalcula_Cuota(CDbl(txtMonto.Text), CDbl(txtPlazo.Text), CDbl(txtTasa.Text), mFrecuenciaPago), "Standard")
        Else
            txtCuota.Text = 0
        End If
        
End Select

Call sbCalcularCompromiso
Call sbEstructuraActualiza(5, False)

    Exit Sub
vError:
    MsgBox "Ocurrió un error al calcular la cuota. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub


Private Sub chkCargaAsociacion_Click()
If vPaso Then Exit Sub
m_CambioCalculo = True

Call sbCalcula_Cargas
End Sub

Private Sub chkCargaFrap_Click()
If vPaso Then Exit Sub
m_CambioCalculo = True

Call sbCalcula_Cargas
End Sub

Private Sub chkPolizaDesempleo_Click()
  Call sbCalculaPolizaDesempleo
  Call sbCalcularCompromiso
  m_CambioDatos = True
End Sub

Private Sub chkPolizaIncendio_Click()
  Call sbCalculaPolizaDeIncendio
  Call sbCalcularCompromiso
  m_CambioDatos = True
End Sub



Private Sub chkPolizaVehiculo_Click()
  Call sbCalculaPolizaDePrenda
  Call sbCalcularCompromiso
  m_CambioDatos = True
End Sub

Private Sub chkPolizaVida_Click()
    Call sbCalculaPolizaDeVida
    Call sbCalcularCompromiso
    m_CambioDatos = True
End Sub

Private Sub sbCalculaPolizaDeIncendio()

txtPolizaIncendio.Text = 0

If Val(txtMonto.Text) = 0 Then Exit Sub

If chkPolizaIncendio.Value = Checked Then
    
    'Utiliza el Monto del Credito por Default
    glogon.strSQL = "select  dbo.fxCrd_Prea_Poliza_Incendio(" & fxFormatearValor(CDbl(txtMonto.Text), Numerico) & " )"
    
    'Si se especifica el Monto de la Contrucción, Entonces Cambiar el Calculo de la Poliza de Incendio por el este.
    If IsNumeric(txtMontoConstruccion.Text) Then
        If CCur(txtMontoConstruccion.Text) > 1000 Then
            glogon.strSQL = "select  dbo.fxCrd_Prea_Poliza_Incendio(" & fxFormatearValor(CDbl(txtMontoConstruccion.Text), Numerico) & " )"
        End If
    End If
        
    If execSql(glogon.strSQL) Then
        If Trim(glogon.Recordset(0) & "") <> "" Then
            txtPolizaIncendio.Text = Format(glogon.Recordset(0), "Standard")
        End If
    End If
End If
End Sub


Private Sub sbCalculaPolizaDePrenda()

txtPolizaPrenda.Text = 0

If Val(txtMonto.Text) = 0 Then Exit Sub

If chkPolizaVehiculo.Value = Checked Then
    
    'Utiliza el Monto del Credito por Default
    glogon.strSQL = "select  dbo.fxCrd_Prea_Poliza_Vehiculo(" & fxFormatearValor(CDbl(txtMonto.Text), Numerico) & " )"
    
'    'Si se especifica el Monto de la Contrucción, Entonces Cambiar el Calculo de la Poliza de Incendio por el este.
'    If IsNumeric(txtMontoConstruccion.Text) Then
'        If CCur(txtMontoConstruccion.Text) > 1000 Then
'            glogon.strSQL = "select  dbo.fxCrd_Prea_Poliza_Incendio(" & fxFormatearValor(CDbl(txtMontoConstruccion.Text), Numerico) & " )"
'        End If
'    End If
        
    If execSql(glogon.strSQL) Then
        If Trim(glogon.Recordset(0) & "") <> "" Then
            txtPolizaPrenda.Text = Format(glogon.Recordset(0), "Standard")
        End If
    End If
End If
End Sub


Private Sub sbCalculaPolizaDesempleo()
txtPolizaDesempleo.Text = 0

If Val(txtCuota.Text) = 0 Then Exit Sub

Dim pMonto As Currency

If Not IsNumeric(txtPolizaVida.Text) Then
        txtPolizaVida.Text = "0"
End If

If Not IsNumeric(txtPolizaIncendio.Text) Then
        txtPolizaIncendio.Text = "0"
End If

pMonto = CCur(txtCuota.Text) + CCur(txtPolizaVida.Text) + CCur(txtPolizaIncendio.Text)

If chkPolizaDesempleo.Value = Checked Then
    glogon.strSQL = "select  dbo.fxCrd_Prea_Poliza_Desempleo(" & fxFormatearValor(pMonto, Numerico) & " )"
    If execSql(glogon.strSQL) Then
        If Trim(glogon.Recordset(0) & "") <> "" Then
            txtPolizaDesempleo.Text = Format(glogon.Recordset(0), "Standard")
        End If
    End If
End If
End Sub

Private Sub sbCalculaPolizaDeVida()
txtPolizaVida.Text = 0
If Val(txtMonto.Text) = 0 Then Exit Sub
If chkPolizaVida.Value = Checked Then
    glogon.strSQL = "select  dbo.fxCrd_Prea_Poliza_Vida(" & fxFormatearValor(CDbl(txtMonto.Text), Numerico) & " )"
    '                         dbo.fxCRDCuotaPolizaVida(Monto,cod_linea,garantia)
    If execSql(glogon.strSQL) Then
        If Trim(glogon.Recordset(0) & "") <> "" Then
            txtPolizaVida.Text = Format(glogon.Recordset(0), "Standard")
        End If
    End If
End If
End Sub

Private Sub chkPrimerCuota_Click()
    m_CambioDatos = True
    Call sbEstructuraActualiza(1000, True)
End Sub


Public Function fxAgregaColleccion(ByVal Expediente As String, ByVal pObsAnalisista As String, ByVal pObsComite As String, ByVal pObsJuntaDirectiva As String) As String
On Error GoTo error
Dim Vcoleccion As New Collection
With Vcoleccion
    .Add fxFormatearValor(Expediente, caracter)
    .Add fxFormatearValor(pObsAnalisista, caracter)
    .Add fxFormatearValor(pObsComite, caracter)
    .Add fxFormatearValor(pObsJuntaDirectiva, caracter)
End With
fxAgregaColleccion = fxFormatearValuesCollection(Vcoleccion)

Exit Function
error:
    MsgBox fxSys_Error_Handler(Err.Description)
End Function

Private Sub sbGuardaObservaciones()
 
If Len(vObservacion(0)) + Len(vObservacion(1)) + Len(vObservacion(2)) = 0 Then Exit Sub
 
On Error GoTo vError
 


If ((clsMensajes.Estado = "A") Or (clsMensajes.Estado = "D")) Then
    MsgBox "No es posible realizar cambios en las observaciones del expediente seleccionado.", vbInformation, gMsgTitulo
    Exit Sub
End If


Me.MousePointer = vbHourglass

clsEntidad.tablaName = "spCRDPreaObservaciones"
If clsEntidad.fxModificar(fxAgregaColleccion(gPreAnalisis.Expediente, vObservacion(0), vObservacion(1), vObservacion(2))) Then
    MsgBox "La información se registro correctamente.", vbExclamation
    m_CambioObservaciones = False
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox "Ocurrió un error en el proceso de guardar observaciones. " & "err: " & Err.Description, vbCritical
End Sub

Private Sub chkS_Constancia_Click()
If vPaso Then Exit Sub

vPaso = True

txtS_Constancia.Text = "0"

If chkS_Constancia.Value = xtpChecked Then
 txtS_Constancia.BackColor = vbWhite
 
 chkS_OrdenPatronal.Value = xtpUnchecked
 txtS_OrdenPatronal.Text = "0"
Else
 txtS_Constancia.BackColor = RGB(187, 215, 247)
End If

vPaso = False

End Sub

Private Sub chkS_OrdenPatronal_Click()
If vPaso Then Exit Sub

vPaso = True

txtS_OrdenPatronal.Text = "0"

If chkS_OrdenPatronal.Value = xtpChecked Then
 txtS_OrdenPatronal.BackColor = vbWhite
 
 chkS_Constancia.Value = xtpUnchecked
 txtS_Constancia.Text = "0"
Else
 txtS_OrdenPatronal.BackColor = RGB(187, 215, 247)
End If

vPaso = False

End Sub

Private Sub cmdGuardaObservaciones_Click()
If m_CambioObservaciones Then
    Call sbGuardaObservaciones
End If
    
End Sub

Private Sub cmdScrollBar(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

  m_Expediente = ""
  
strSQL = "SELECT TOP 1 cod_preanalisis From CRD_PREA_PREANALISIS" _
       & " WHERE TIPO_PREANALISIS = 'E'"
  
  Select Case Index
  
    Case 0
        If txtExpediente.Text = "" Then
           txtExpediente.Text = "99999999999"
        End If
    
        strSQL = strSQL & " AND cod_preanalisis < '" & txtExpediente.Text & "'"
    Case 1
        If txtExpediente.Text = "" Then
           txtExpediente.Text = "0"
        End If
        
        strSQL = strSQL & " AND cod_preanalisis > '" & txtExpediente.Text & "'"
        
  End Select
  
        
  strSQL = strSQL & " AND TRY_CAST(cod_preanalisis AS BIGINT) IS NOT NULL ORDER BY TRY_CAST(cod_preanalisis AS BIGINT) DESC"
  
  Call OpenRecordSet(rs, strSQL)


  If Not rs.EOF And Not rs.BOF Then

  txtExpediente.Text = rs!cod_preanalisis
   Call txtExpediente_LostFocus
  
'  Call SbTraerDatosExpediente
    rs.Close
  End If



Me.MousePointer = vbDefault

Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbNumPagos_Update()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not m_Cargando And txtNombre.Text <> "" Then
  
    strSQL = "exec spCrd_Prea_NumPagos '" & txtCedula.Text & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"
    Call OpenRecordSet(rs, strSQL)
        m_NumPagos = rs!Num_Pagos
    rs.Close
End If

Exit Sub

vError:

End Sub


Private Sub dtpCorte_Change()
m_CambioCalculo = True
Call sbNumPagos_Update

End Sub


Private Sub dtpCorte_LostFocus()
'    tcMain.Item(1).Selected = True
'    txtSalarioDevengado.SetFocus
End Sub

Private Sub dtpFecNac_Change()
    Call sbCalcularPlazoMaximo
    m_CambioDatos = True
End Sub





Private Sub Form_Activate()
    vModulo = 3 'Modulo de Credito
 
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.ActiveControl.Name <> "txtObservaciones" Then
    If (KeyCode = vbKeyReturn) Then 'Or KeyCode = vbKeyTab) Then
'        Call sbCalcularCuota(Me.ActiveControl.Name)
    Call gsbPulsarTecla(vbKeyTab)
    
    ElseIf KeyCode = vbKeyF4 Then
        Call sbBusqueda(Me.ActiveControl.Name)
    End If
End If
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo vError
Select Case Me.ActiveControl.Name
Case "txtMonto"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtMonto.Text), KeyAscii)
Case "txtPolizaVida"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtPolizaVida.Text), KeyAscii)
Case "txtPolizaIncendio"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtPolizaIncendio.Text), KeyAscii)
Case "txtPolizaDesempleo"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtPolizaDesempleo.Text), KeyAscii)

Case "txtPlazo"
    KeyAscii = fxPermiteSoloDigitos(KeyAscii)
    txtPlazo.Text = fxValidaLargoZeroIzq(Trim$(txtPlazo.Text))
Case "txtTasa"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtTasa.Text), KeyAscii)
    txtTasa.Text = fxValidaLargoZeroIzq(Trim$(txtTasa.Text))
Case "txtSalarioDevengado"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtSalarioDevengado.Text), KeyAscii)
'Validar si esta se quedan por esta siempre bloqueadas
Case "txtRebajoExtras"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtRebajoExtras.Text), KeyAscii)
Case "txtSalarioReal"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtSalarioReal.Text), KeyAscii)
Case "txtCompAdicionalBase"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtCompAdicionalBase.Text), KeyAscii)
Case "txtDevengadoMes"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtDevengadoMes.Text), KeyAscii)
Case "txtPorcSobreSalario"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtPorcSobreSalario.Text), KeyAscii)
Case "txtDeducciones"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtDeducciones.Text), KeyAscii)
Case "txtCrdTransitoCancelados"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtCrdTransitoCancelados.Text), KeyAscii)
Case "txtCrdTransitoXCobrar"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtCrdTransitoXCobrar.Text), KeyAscii)
Case "txtSalarioLiquido"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtSalarioLiquido.Text), KeyAscii)
Case "txtRefundiciones"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtRefundiciones.Text), KeyAscii)
Case "txtDesembolsos"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtDesembolsos.Text), KeyAscii)
Case "txtTotalLiquido"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtTotalLiquido.Text), KeyAscii)
Case "txtFianzas"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtFianzas.Text), KeyAscii)
Case "txtLiquidezSinFianza"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtLiquidezSinFianza.Text), KeyAscii)
Case "txtLiquidezPorcSinFianza"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtLiquidezPorcSinFianza.Text), KeyAscii)
Case "txtLiquidezConFianza"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtLiquidezConFianza.Text), KeyAscii)
Case "txtLiquidezPorcConFianza"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtLiquidezPorcConFianza.Text), KeyAscii)
Case "txtMontoGirar"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtMontoGirar.Text), KeyAscii)
'Fin/Validar si esta se quedan por esta siempre bloqueadas

Case "txtCargaCCSS"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(lblCargaCCSS.Caption), KeyAscii)
Case "txtCargaImpSalario"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(lblCargaImpSalario.Caption), KeyAscii)
Case "txtCargaAsociacion"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(lblCargaAsociacion.Caption), KeyAscii)
Case "txtCargaFrap"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(lblCargaFrap.Caption), KeyAscii)
End Select


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar la información de los formatos. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub
  


'Eventos del forms
Private Sub Form_Load()
 
vModulo = 3 'Modulo de Credito
 
m_NumPagos = 2
 
mFecha = fxFechaServidor
 
UpDownExpediente.Height = txtExpediente.Height
cboSubExpediente.Height = txtExpediente.Height
lblEstado.Height = txtExpediente.Height
 
txtSalarioLiquido.BackColor = RGB(187, 215, 247)
txtSalarioDevengado.BackColor = RGB(187, 215, 247)
txtCompAdicionalBase.BackColor = RGB(187, 215, 247)

txtPlMax.BackColor = RGB(187, 215, 247)
txtAsignado.BackColor = RGB(187, 215, 247)
txtClasificacion.BackColor = RGB(187, 215, 247)
txtCuota.BackColor = RGB(187, 215, 247)
txtCompromiso.BackColor = RGB(187, 215, 247)

txtEdad.BackColor = RGB(187, 215, 247)

txtEstadoSocio.BackColor = RGB(187, 215, 247)
txtClasificacion.BackColor = RGB(187, 215, 247)


txtEjecutivo.BackColor = RGB(187, 215, 247)
txtOficina.BackColor = RGB(187, 215, 247)

txtS_Mensual.BackColor = RGB(187, 215, 247)

txtSalarioMinimoInembargable.BackColor = RGB(187, 215, 247)
txtSalarioNormativa.BackColor = RGB(187, 215, 247)


lblCargaAsociacion.BackColor = RGB(187, 215, 247)
lblCargaCCSS.BackColor = RGB(187, 215, 247)
lblCargaFrap.BackColor = RGB(187, 215, 247)
lblCargaImpSalario.BackColor = RGB(187, 215, 247)

txtExpediente.Height = cboSubExpediente.Height

 mFrecuenciaPago = "M"

With lswIncapacidades.ColumnHeaders
  .Clear
  .Add , , "Desde", 1500, vbCenter
  .Add , , "Hasta", 1500, vbCenter
  .Add , , "Días", 1000, vbRightJustify
End With


With lswD_Lista.ColumnHeaders
  .Clear
  .Add , , "Id", 1500
  .Add , , "Nombre", 3000
  .Add , , "Nombre Giro", 3000
  .Add , , "Modifica", 10
End With

With lswP_Examenes.ColumnHeaders
  .Clear
  .Add , , "Evento", 3500
  .Add , , "Usuario", 2500, vbCenter
End With

vPaso = True

    dtpR_Formaliza.Value = mFecha
    
    cboD_Ordinario.AddItem "Sí"
    cboD_Ordinario.AddItem "No"
    cboD_Ordinario.Text = "Sí"
    
    cboCantidadFiadores.Clear
    cboCantidadFiadores.AddItem "0"
    cboCantidadFiadores.AddItem "1"
    cboCantidadFiadores.AddItem "2"
    cboCantidadFiadores.AddItem "3"
    cboCantidadFiadores.AddItem "4"
    cboCantidadFiadores.Text = "0"
    
vPaso = False

Call sbInicializaGlobales

'Para Control de Registro de Persona
Call sbAFIParametrosCargaArreglo


lblPorcentajeSalario.Caption = str(GlobalPorcLiquidezLibre) & " % Sobre Salario "

txtSalarioMinimoInembargable.Text = Format(GlobalSalarioMinimoInembargable, "Standard")
txtSalarioNormativa.Text = Format(GlobalSalarioNormativo, "Standard")


'txtCedula.SetFocus
 
'Inicializa Barra
'Call sbToolBarIconos(tlb)
'Call sbToolBar(tlb, "nuevo")

Call btnBarra_Click(0)

'Inicializa Seguridad
Call Formularios(Me)
Call RefrescaTags(Me)

vPaso = True

Call sbCargarCombos
Call sbAccionVentana(NuevoRegistro)

Call sbInicializaComboExpediente

tcMain.Item(0).Selected = True

Call sbBloquearTab

m_CargoSalario = True
m_FECHA_CREACION = "-1"
cboCantidadFiadores.ListIndex = 0
m_CambioDatos = False
m_CambioCalculo = False
m_CambioObservaciones = False

vPaso = False

Me.MousePointer = vbDefault

End Sub


Private Sub Form_Unload(Cancel As Integer)

    sbDeseaGuardar
    
    Set clsMensajes = Nothing
    Set clsEntidad = Nothing
    Set clsNull = Nothing

End Sub


Private Sub sbDeseaGuardar()

    If (m_CambioDatos = True) Or (m_CambioCalculo = True) Or (m_CambioObservaciones = True) Then
        If (MsgBox("¿Desea guardar los cambios efectuados en el expediente seleccionado. ?", vbQuestion + vbYesNo, gMsgTitulo) = vbYes) Then
            Call fxGuardar
            Call sbGuardaObservaciones
        End If
    End If

End Sub





Private Sub sbDeducciones_Borrar(fila As Long)
Dim vID As Long

On Error GoTo vError
 
    If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
        Exit Sub
    End If
    
    
 Me.MousePointer = vbHourglass
    
    gDeducciones.Col = 1
    gDeducciones.Row = fila
    vID = Val(gDeducciones.Text)
    
    clsEntidad.tablaName = "spCRDPreaDETALLE_DEDUC"
    If (MsgBox("¿ Desea borrar la información seleccionada?", vbQuestion + vbYesNo, gMsgTitulo) = vbYes) Then
        If clsEntidad.fxRemover(fxAgregaColleccionBorrar(vID, gPreAnalisis.Expediente)) Then
            m_CambioDatos = True
        End If
    End If

    Me.MousePointer = vbDefault
    Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo

End Sub

Private Sub sbCreditos_Cuotas_Borrar(Tipo As String)

On Error GoTo vError
 
    If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
        Exit Sub
    End If
    
    
 Me.MousePointer = vbHourglass
    
    If (MsgBox("¿ Desea borrar la información de Cuotas en Transito?", vbQuestion + vbYesNo, gMsgTitulo) = vbYes) Then
        
        strSQL = "exec spCrdPreaEliminarCreditosCuotasCxC '" & txtExpediente.Text & "', '" & Tipo & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Elimina", "Cuotas en Transito (" & IIf(Tipo = "A", "Por Cobrar", "Canceladas") & ")del Estudio de Crédito No.: " & txtExpediente.Text)
        
        m_CambioDatos = True
            
    End If

    Me.MousePointer = vbDefault
    Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo

End Sub


Private Sub sbCreditos_Cuotas_Registra(fila As Long, Tipo As String)
Dim vID As Long

On Error GoTo vError
 
    If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
        Exit Sub
    End If
    
    
'spCrdPreaRegistrarCreditosCuotasCxC](
'                              @COD_PREANALISIS varchar(20),@CUOTA decimal(10,2),@ESTADO char(1),@DETALLE varchar(250)
    
    
 Me.MousePointer = vbHourglass
    
 Dim pCuotas As Currency, pDetalle As String, pCuota As Currency
        
 Select Case Tipo
    Case "A" 'Por Cobrar
      With gCuotasCobrar
        .Row = fila
        .Col = 2
        pDetalle = .Text
        .Col = 3
        pCuota = CCur(.Text)
      
      End With
      
    Case "C" 'Canceladas
    
      With gCuotasCancela
        .Row = fila
        .Col = 2
        pDetalle = .Text
        .Col = 3
        pCuota = CCur(.Text)
      
      End With
    
 End Select
        
        
strSQL = "exec spCrdPreaRegistrarCreditosCuotasCxC '" & txtExpediente.Text & "', " & pCuota & ", '" & Tipo & "', '" & pDetalle & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Cuotas en Transito (" & IIf(Tipo = "A", "Por Cobrar", "Canceladas") & ")del Estudio de Crédito No.: " & txtExpediente.Text)

m_CambioDatos = True
        


Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
Resume
End Sub



Private Sub gCuotasCancela_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError
            
If KeyCode = vbKeyReturn And gCuotasCancela.ActiveCol = gCuotasCancela.MaxCols Then
    Call sbCreditos_Cuotas_Registra(gCuotasCancela.ActiveRow, "C")
    Call sbCreditosTransito_Load("C")
End If
            
If KeyCode = vbKeyDelete Then
    Call sbCreditos_Cuotas_Borrar("C")
    Call sbCreditosTransito_Load("C")
End If
    
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub



Private Sub gCuotasCobrar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
            
If KeyCode = vbKeyReturn And gCuotasCobrar.ActiveCol = gCuotasCobrar.MaxCols Then
    Call sbCreditos_Cuotas_Registra(gCuotasCobrar.ActiveRow, "A")
    Call sbCreditosTransito_Load("A")
End If
            
If KeyCode = vbKeyDelete Then
    Call sbCreditos_Cuotas_Borrar("A")
    Call sbCreditosTransito_Load("A")
End If
    
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub gDeducciones_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError
            
If KeyCode = vbKeyDelete Then
    Call sbDeducciones_Borrar(gDeducciones.ActiveRow)
    Call sbDeducciones_Load
End If
    
Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub gExtras_Advance(ByVal AdvanceNext As Boolean)
Dim fila As Integer
Dim wColValue As Variant

wColValue = gExtras.Text

fila = gExtras.ActiveRow

If (gExtras.MaxRows > 1) And (gExtras.ActiveRow = 1) Then
    Exit Sub
ElseIf (gExtras.ActiveRow = gExtras.MaxRows) Then
    gExtras.Col = 1
    gExtras.Row = gExtras.ActiveRow
    If (fxExtras_Inserta(fila) = True) Then
        gExtras.Col = 1
        gExtras.Row = gExtras.MaxRows
        gExtras.Action = 0
        
        vChanged = False
        
        gExtras.Col = 2
        gExtras.CellType = CellTypeComboBox
        
        gExtras.TypeComboBoxList = mTipoExtraLista
        gExtras.TypeComboBoxEditable = False
        gExtras.Text = mTipoExtra
    End If
    
End If
End Sub



Private Sub sbExtra_Guarda()
    If (vChanged = True) Then
        If (gExtras.Row = gExtras.MaxRows) Then
            If (fxExtras_Inserta(gExtras.Row) = True) Then
                vChanged = False
            End If
        Else
            If (fxExtras_Modifica(gExtras.ActiveRow) = True) Then
                vChanged = False
            End If
        End If
    End If
End Sub

Private Sub sbExtras_Calcula(Optional pInicial As Boolean = False)
Dim i As Integer, curMonto As Currency

curMonto = 0

For i = 1 To gExtras.MaxRows
  gExtras.Row = i
  gExtras.Col = 3
  curMonto = curMonto + IIf((gExtras.Text = ""), 0, gExtras.Text)
Next i
 
txtT_Extras.Text = Format(curMonto, "Standard")
txtRebajoExtras.Text = Format(curMonto, "Standard")

If Not pInicial Then
'Ajusta el Devengado
    txtSalarioDevengado.Text = txtS_Devengado.Text
    Call sbEstructuraActualiza(1, False)
    m_CambioCalculo = True

    txtS_Devengado.Text = Format(txtS_Devengado.Text, "Standard")
    'Replica en el Resumen
    txtSalarioDevengado.Text = txtS_Devengado.Text

End If

m_CambioDatos = True
End Sub



Private Function fxExtras_Inserta(ByRef fila As Integer) As Boolean
Dim vID As String
Dim vCodExtra As String
Dim vDesExtra As String
Dim vMonto As Double

On Error GoTo vError

     fxExtras_Inserta = False
     If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
      Exit Function
     End If
     
    gExtras.Row = fila
    vID = 0
    vMonto = 0
    
    gExtras.Row = fila
    gExtras.Col = 3
        If Val(gExtras.Text) > 0 Then
            vMonto = CDbl(gExtras.Text)
        Else
            MsgBox "Monto es requerido.", vbExclamation, gMsgTitulo
            Me.MousePointer = vbDefault
            Exit Function
        End If
    
    gExtras.Col = 2
    If gExtras.Text = "" Then
       vCodExtra = SIFGlobal.fxCodText(mTipoExtra)
    Else
       vCodExtra = SIFGlobal.fxCodText(gExtras.Text)
       mTipoExtra = gExtras.Text
    End If
    
    If Len(vCodExtra) = 0 Then
        MsgBox "Es requerido seleccionar una tipo de extra.", vbExclamation, gMsgTitulo
        Me.MousePointer = vbDefault
        Exit Function
    End If

    
 Me.MousePointer = vbHourglass
    
    clsEntidad.tablaName = "spCRDPreaDETALLE_EXTRAS"
    If (clsEntidad.fxAgregar(fxAgregaColleccion(vID, gPreAnalisis.Expediente, vCodExtra, vMonto))) Then
        fxExtras_Inserta = True
        gExtras.Col = 1
        gExtras.Lock = True
        gExtras.MaxRows = gExtras.MaxRows + 1
       
        glogon.strSQL = "select max(IDX) as IDX from CRD_PREA_DETALLE_EXTRAS  Where cod_preanalisis = " & fxFormatearValor(gPreAnalisis.Expediente, caracter)
        If (execSql(glogon.strSQL, True)) Then
            gExtras.Col = 1
            gExtras.Row = gExtras.ActiveRow
            gExtras.Text = glogon.Recordset!IdX
        Else
            MsgBox "Error obteniendo el consecutivo del registro ingresado", vbExclamation, gMsgTitulo
        End If
    End If
    
Call sbExtras_Calcula
    
Me.MousePointer = vbDefault

Exit Function

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
End Function

Private Function fxExtras_Modifica(fila As Long) As Boolean
Dim vID As Integer
Dim vCodExtra As String
Dim vNuevaDesExtra As String
Dim vMonto As Double

On Error GoTo vError

    
fxExtras_Modifica = False

 If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
  Exit Function
 End If

gExtras.Row = fila
gExtras.Col = 1
    vID = gExtras.Text
gExtras.Col = 3
    vMonto = gExtras.Text

gExtras.Col = 2
If gExtras.Text = "" Then
    vCodExtra = SIFGlobal.fxCodText(mTipoExtra)
Else
    vCodExtra = SIFGlobal.fxCodText(gExtras.Text)
End If


Me.MousePointer = vbHourglass

clsEntidad.tablaName = "spCRDPreaDETALLE_EXTRAS"

If (clsEntidad.fxModificar(fxAgregaColleccion(vID, gPreAnalisis.Expediente, vCodExtra, vMonto))) Then
    gExtras.Col = 1
    fxExtras_Modifica = True
Else
    Me.MousePointer = vbDefault
    MsgBox "No se pudo actualizar la información seleccionada.", vbExclamation, gMsgTitulo
End If

Me.MousePointer = vbDefault

Exit Function

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo

End Function

Private Function fxExtras_Borrar(fila As Long) As Boolean
Dim vID As Integer

On Error GoTo vError
 Me.MousePointer = vbHourglass
 
    fxExtras_Borrar = False
    
     If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
      GoTo salir
     End If
 
    gExtras.Col = 1
    gExtras.Row = fila
    vID = Val(gExtras.Text)
    
    clsEntidad.tablaName = "spCRDPreaDETALLE_EXTRAS"
    If (MsgBox("¿ Desea borrar la información seleccionada?", vbQuestion + vbYesNo, gMsgTitulo) = vbYes) Then
        If clsEntidad.fxRemover(fxAgregaColleccionBorrar(vID, gPreAnalisis.Expediente)) Then
            gExtras.Col = 1
            fxExtras_Borrar = True
            
            gExtras.DeleteRows fila, 1
            gExtras.MaxRows = gExtras.MaxRows - 1
            gExtras.Row = gExtras.ActiveRow
                
                Call sbExtras_Calcula
        End If
        gExtras.SetFocus
    End If

salir:
    Me.MousePointer = vbDefault
    Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
    
    Resume salir
End Function

Private Sub sbExtra_Nuevo()
    gExtras.Col = 1
    gExtras.Row = gExtras.MaxRows
    

    gExtras.Col = 2
    gExtras.CellType = CellTypeComboBox
    
    gExtras.TypeComboBoxList = mTipoExtraLista
    gExtras.TypeComboBoxEditable = False
    gExtras.Text = mTipoExtra
    
    gExtras.Action = 0
    gExtras.SetFocus
End Sub


Private Sub gExtras_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sql As String

On Error GoTo error
    If ((Shift And vbCtrlMask) <> 0) And (KeyCode = vbKeyS) Then
'        UnLoad Me
'        DoEvents
        m_CambioCalculo = True
        
    ElseIf ((Shift And vbCtrlMask) <> 0) And (KeyCode = vbKeyN) Then
        Call sbExtra_Nuevo
            
    ElseIf ((Shift And vbCtrlMask) <> 0) And (KeyCode = vbKeyG) Then
        Call sbExtra_Guarda
            
    ElseIf (KeyCode = vbKeyDelete) Then
        Call fxExtras_Borrar(gExtras.ActiveRow)
        Call sbExtra_Nuevo
    End If
    
salir:
    Exit Sub
error:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
End Sub


Private Sub gExtras_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
 vChanged = True
End Sub

Private Sub gFianzas_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

On Error GoTo vError

If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
    Exit Sub
End If

If Col = 7 Or Col = 8 Then
 
    Call sbFianzas_Calcula
    
    m_CambioCalculo = True

   gFianzas.Row = Row
   gFianzas.Col = 7
   strSQL = "update CRD_PREA_DETALLE_FIANZAS set Aplica = " & gFianzas.Value
   gFianzas.Col = 8
   strSQL = strSQL & ", Cancela_Mora = " & gFianzas.Value & " where cod_PreAnalisis = '" _
          & txtExpediente.Text & "' and id_solicitud = "
   gFianzas.Col = 1
   strSQL = strSQL & gFianzas.Text
   
   Call ConectionExecute(strSQL)

End If


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub gRefunde_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
    Exit Sub
End If


gRefunde.Row = Row
gRefunde.Col = Col

m_CambioCalculo = True

Select Case Col
 Case 11 'Todo
   If gRefunde.Value = 1 Then
      gRefunde.Col = 8
      gRefunde.Value = 0
   End If
   
 Case 12 'Mora
   If gRefunde.Value = 1 Then
      gRefunde.Col = 7
      gRefunde.Value = 0
   End If
   
End Select


If Col = 11 Or Col = 12 Then

   gRefunde.Row = Row
   gRefunde.Col = 11
   strSQL = "update CRD_PREA_REFUNDICIONES set Aplica = " & gRefunde.Value
   gRefunde.Col = 12
   strSQL = strSQL & ", Apl_Mora = " & gRefunde.Value _
          & " where cod_PreAnalisis = '" & txtExpediente.Text & "' and id_solicitud = "
   gRefunde.Col = 1
   strSQL = strSQL & gRefunde.Text
   Call ConectionExecute(strSQL)
   
   m_CambioDatos = True
   
   Call sbRefundiciones_Calcula
   
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub imgFraCerrar_Click()

Dim curTotal As Currency, curSalario As Currency, FrapPorc As Currency

On Error GoTo vError
    
    curTotal = 0
    curSalario = CCur(txtDevengadoMes.Text)
    FrapPorc = CCur(txtFrapPorc.Text)
    
    If CCur(lblCargaImpSalario.Caption) > 0 Then
        curTotal = curTotal + lblCargaImpSalario.Caption
    End If
    
    If IsNumeric(GlobalPorcCCSS) Then
      lblCargaCCSS.Caption = Format((curSalario * GlobalPorcCCSS) / 100, "Standard")
    End If
    
    
   If chkCargaAsociacion.Value = Checked Then
        lblCargaAsociacion.Caption = Format((curSalario * GlobalPorcAsocSolidarista) / 100, "Standard")
        chkCargaAsociacion.Tag = "S"
   Else
        lblCargaAsociacion.Caption = "0.00"
        chkCargaAsociacion.Tag = "N"
   End If
   
   If chkCargaFrap.Value = Checked Then
        lblCargaFrap.Caption = Format((curSalario * GlobalPorcFRAPFAP) / 100, "Standard")
        chkCargaFrap.Tag = "S"
    Else
        chkCargaFrap.Tag = "N"
   End If
   
    curTotal = curTotal + CDbl(lblCargaCCSS.Caption)
    curTotal = curTotal + CDbl(lblCargaImpSalario.Caption)
    curTotal = curTotal + CDbl(lblCargaAsociacion.Caption)
    
    'Total Sección Resumen
    txtTotal_Cargas_CCSS.Text = Format(curTotal, "Standard")
    
    'Total Sección Deducciones
    txtD_TotalCargas.Text = Format(curTotal, "Standard")
  
   If CCur(txtTotal_Cargas_CCSS.Text) <> m_curValor_Anterior Then
        Call sbEstructuraActualiza(3, False)
        m_CambioCalculo = True
   End If
  
  txtPorcSobreSalario.SetFocus
'  fraDCargas.Enable = False

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar datos de Cargas sociales . " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub


Private Sub lblSalarioDevengado_DblClick(Index As Integer)
    Dim sql As String
    
    sql = "spCRDPreaSalariosLista " & fxFormatearValor(gPreAnalisis.Expediente, caracter)
        
    If clsEntidad.fxEjecutaSQL(sql) Then
        m_SoloVerSalarios = True
        Call sbSalarios_Registro_Inicial
    End If

        
End Sub



Private Sub lswArchivos_DblClick()
Dim cn As New ADODB.Connection

Dim sql As String, Campo_Imagen As String
Dim rs As New ADODB.Recordset, Stream As New ADODB.Stream
Dim pPath As String

Dim vArchivo As String

If lswArchivos.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

Set itmX = lswArchivos.SelectedItem

On Error Resume Next

vArchivo = "ProGrX_Estudio Crédito_" & txtExpediente.Text & "_Doc.Id_" & Format(itmX.Text, "00000") & " " & itmX.SubItems(1)


MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Adjuntos\"

pPath = SIFGlobal.DirectorioDeResultados & "\Adjuntos\" & vArchivo

'------------------------------------------------------------------------
  
On Error GoTo vError

Dim pPass As Boolean
  
lblLoading.Caption = "Abriendo archivo..Espere!"
DoEvents
  
Me.MousePointer = vbHourglass
  
pPass = False

Campo_Imagen = "DOC_ADJUNTO"
  
sql = "select " & Campo_Imagen & " from CRD_PREA_V2_ADJUNTOS Where ID_ADJUNTO = " & itmX.Text

'---Version 2
Dim fileData() As Byte

Set rs = New ADODB.Recordset
rs.Open sql, glogon.Conection, adOpenStatic, adLockReadOnly

' Guardar los datos del archivo en el disco
If Not rs.EOF Then
    fileData = rs.Fields("DOC_ADJUNTO").Value
    Open pPath For Binary Access Write As #1
    Put #1, , fileData
    Close #1
    
    pPass = True
End If
rs.Close

'---Version 2: Fin


lblLoading.Caption = "Archivo Guardado en: " & pPath

Me.MousePointer = vbDefault

If pPass Then
    'Abre el Archivo
    Call Shell("Explorer.exe /e," & pPath, vbNormalFocus)
Else
    MsgBox "No fue posible visualizar el documento! ", vbExclamation
End If


Exit Sub

vError:
  lblLoading.Caption = ""

Me.MousePointer = vbDefault
 
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
 
 On Error Resume Next
 
 'Si no abre el archivo automáticamente, entonces abre el directorio
 Call Shell("Explorer.exe /select," & pPath, vbNormalFocus)

End Sub





Private Sub lswD_Lista_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

txtDS_Descripcion.Text = Item.SubItems(2)
txtDS_Descripcion.Tag = Item.Text

If Item.SubItems(3) = 1 Then
    txtDS_Descripcion.Locked = False
Else
    txtDS_Descripcion.Locked = True
End If


End Sub


Private Sub optCausas_Click(Index As Integer)
Call sbCargarListaCausas
End Sub

Private Sub optObservacion_Click(Index As Integer)
On Error GoTo vError

Select Case Index
 Case 0
  txtObservaciones.Text = vObservacion(0)
 Case 1
  txtObservaciones.Text = vObservacion(1)
 Case 2
  txtObservaciones.Text = vObservacion(2)
End Select

txtObservaciones.SetFocus
m_CambioObservaciones = False

vError:

End Sub


Private Sub rbActas_Click(Index As Integer)
Call sbEstudio_Comite_Resolucion_Load(txtExpediente.Text)
End Sub



Private Sub sbEstudio_Comite_Resolucion_Load(pExpediente As String)

Dim pTipo As String


On Error GoTo vError

Me.MousePointer = vbHourglass

lswAutorizadores.Checkboxes = False
lswAutorizadores.ListItems.Clear

With lswAutorizadores.ColumnHeaders
    .Clear
    
    Select Case True
        Case rbActas.Item(0).Value
          pTipo = "RES"
          .Add , , "Estado", 1800, vbCenter
          .Add , , "Fecha", 1800
          .Add , , "Usuario", 2100, vbCenter
          .Add , , "Notas", 3100
        
        Case rbActas.Item(1).Value
          pTipo = "AUT"
          .Add , , "Estado", 1800, vbCenter
          .Add , , "Fecha", 1800
          .Add , , "Identificación", 1800
          .Add , , "Nombre", 3800
          .Add , , "Usuario", 2100, vbCenter
          .Add , , "Notas", 2100
          
        
        Case rbActas.Item(2).Value
          pTipo = "ASI"
          .Add , , "Fecha", 1800
          .Add , , "Identificación", 1800
          .Add , , "Nombre", 3800
          .Add , , "Usuario", 2100, vbCenter
    End Select
    

End With

strSQL = "exec spCrd_Estudio_Resolucion_Detalle '" & pExpediente & "', '" & pTipo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    txtActa.Text = rs!Acta
    txtActaFecha.Text = Format(rs!Acta_Fecha, "dd/MM/yyyy")
End If

Do While Not rs.EOF

    Select Case True
     Case rbActas.Item(0).Value 'Resoluciones
        
        Set itmX = lswAutorizadores.ListItems.Add(, , rs!Estado)
            itmX.SubItems(1) = rs!Registro_Fecha
            itmX.SubItems(2) = rs!Registro_Usuario
            itmX.SubItems(3) = rs!Notas
     
     Case rbActas.Item(1).Value 'Autorizaciones
        
        Set itmX = lswAutorizadores.ListItems.Add(, , rs!Estado)
            itmX.SubItems(1) = rs!Registro_Fecha
            itmX.SubItems(2) = rs!Cedula
            itmX.SubItems(3) = rs!Nombre
            itmX.SubItems(4) = rs!Registro_Usuario
            itmX.SubItems(5) = rs!Notas
            
          
        
     Case rbActas.Item(2).Value 'Asistencia
   
        Set itmX = lswAutorizadores.ListItems.Add(, , rs!Registro_Fecha)
            itmX.SubItems(1) = rs!Cedula
            itmX.SubItems(2) = rs!Nombre
            itmX.SubItems(3) = rs!Registro_Usuario
    
    End Select


  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 1 Then
    rbActas.Item(0).Value = True
  Call sbEstudio_Comite_Resolucion_Load(txtExpediente.Text)
End If
End Sub


Private Sub sbTabChange(Index As Integer)
On Error GoTo vError

Me.MousePointer = vbHourglass

m_PreviousTab = Index
m_MuestraMensaje = False
m_CargoSalario = False

lblPorcentajeSalario.Caption = str(GlobalPorcLiquidezLibre) & " % Sobre Salario "

Select Case Index
    Case 0
            
            If gbDatosPersonales.Enabled Then
                txtNombre.SetFocus
            End If
            
            If m_CambioDatos Then
                Call fxGuardar
            End If
    
    Case 1
    
            cboSalario.SetFocus
        
        If Len(txtExpediente.Text) = 0 Then
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        If m_CambioCalculo Then
            If m_valorComboExp <> "Nuevo SubExpediente" Then
                
                If ((clsMensajes.Estado = "R") Or (clsMensajes.Estado = "P")) Then
                        Call fxGuardar
                End If
            End If
        End If
    
    
    Case 2
'        tcAux.Item(0).Selected = True
'        If ((clsMensajes.Estado = "R") Or (clsMensajes.Estado = "P")) Then
'            If m_CambioObservaciones Then
'                Call sbGuardaObservaciones
'            End If
'        End If
End Select


If Index = 11 Then
    DoEvents
    Call sbClasificacion_CargaGrid
End If

'cboSalario.SetFocus
m_CambioCalculo = False
m_CambioObservaciones = False
m_CambioDatos = False

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    txtMontoGirar.Visible = False
    lblMontoGirar.Item(34).Visible = False
Else
    txtMontoGirar.Visible = True
    lblMontoGirar.Item(34).Visible = True
End If


Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub tcHistorial_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
    Case 0 'Ejecutivos
      Call sbHistorial_Load("E")
    Case 1 'General
      Call sbHistorial_Load("G")
End Select
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)


If txtExpediente.Text = "" Then
    tcMain.Item(0).Selected = True
    Exit Sub
End If


'Call sbTabChange(Item.Index)
pTabBefore = Item.Index


Select Case Item.Index
    Case 1 'Deducciones
'        FrameSalarios.Visible = False
'        fraDCargas.Enable = False
        Call sbDeducciones_Load

    Case 2 'Creditos/Cobros en Transito
        Dim vProceso As Long
        
        vProceso = DatePart("yyyy", dtpCorte.Value)
        If DatePart("m", dtpCorte.Value) < 10 Then
            vProceso = vProceso & "0" & DatePart("m", dtpCorte.Value)
        Else
            vProceso = vProceso & DatePart("m", dtpCorte.Value)
        End If
        
'        'Inicializa Créditos en transito
'        glogon.strSQL = "spCRDPreaCreditosTransito  '" & txtExpediente.Text & "', " & "'I'" & "," & vProceso
'        If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
'            MsgBox "Ocurrió un error al inicializar créditos en transito.", vbInformation, gMsgTitulo
'        End If

        Call sbCreditosTransito_Load("C")
        Call sbCreditosTransito_Load("A")
        
    Case 3 'Refundiciones
        Call sbRefundiciones_Load
    
    Case 4 'Desembolsos
        Call sbDesembolsos_Load
        Call sbDesembolsos_Externos_Lista
        Call sbDesembolsos_Nuevo
        
        
        
    Case 5 'Fianzas
        Call sbFianzas_Load
    
    Case 6 'Resumen
        Call sbDeseaGuardar
    
    Case 7 'Historial
        Call sbHistorial_Load("E")
        
        
    Case 8 'Adjuntos
       Call sbAdjuntos_List
        
    Case 9 'Hipotecario
    
    Case 10 'Prendario
    
    Case 11 'Resolucion
        rbActas.Item(0).Value = True
        
    Case 12 'Causas
       optCausas.Item(0).Value = True
       Call sbCargarListaCausas
        
       optObservacion(0).Value = True
       txtObservaciones.Text = vObservacion(0)
       
End Select

End Sub


Private Sub Timer_Timer()

End Sub

Private Sub txtCedula_Change()
    txtNombre.Text = ""
    txtEdad.Text = ""
    txtEdad.BackColor = RGB(187, 215, 247)
    
     Call sbBloquearTab
     
     m_CambioDatos = True
     If cboSubExpediente.Text = "Nuevo Expediente" Then
        cboCantidadFiadores.Enabled = True
     End If
     
txtEstadoSocio.Visible = False
txtEstadoSocio.Text = ""
End Sub





Private Function fxExiteCedulaEnSubExpediente() As Boolean
On Error GoTo vError

fxExiteCedulaEnSubExpediente = False

Dim vExpPadre As String

If InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0 Then
    vExpPadre = fxDeCodificaPrimaryKey(txtExpediente.Text, 1, "-")
Else
    vExpPadre = txtExpediente.Text
End If


glogon.strSQL = "select nombre from CRD_PREA_PREANALISIS where COD_PREANALISIS = '" & vExpPadre & "' and cedula = '" & txtCedula.Text & "'"
If execSql(glogon.strSQL, True) Then
    
    MsgBox glogon.Recordset!Nombre & " con numero de cedula " & txtCedula.Text & " ya existe como un expediente Maestro, verifique e intente de nuevo.", vbInformation, gMsgTitulo
    fxExiteCedulaEnSubExpediente = True
End If


    Exit Function
vError:
    MsgBox "Ocurrió un error al validar que el numero de cedula. " & "-" & Err.Description, vbCritical, gMsgTitulo
    
End Function


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterno"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    
    If gBusquedas.Resultado <> "" Then
        txtCedula.Text = Trim(gBusquedas.Resultado)
        txtCedula_LostFocus
    End If
End If

End Sub

Public Sub txtCedula_LostFocus()
On Error GoTo vError

Dim RsTemp As ADODB.Recordset

If Len(txtCedula.Text) = 0 Then Exit Sub

txtNombre.Text = ""
gPreAnalisis.Institucion = "-1"
gPreAnalisis.Socio = "N"

If (cboSubExpediente.Text = "Nuevo SubExpediente" Or InStr(1, txtExpediente.Text, "-", vbTextCompare) > 0) Then
    If fxExiteCedulaEnSubExpediente Then Exit Sub
End If

glogon.strSQL = "select S.nombre, S.cod_Institucion, S.ESTADOACTUAL, S.FECHA_NAC, S.sexo" _
              & ",dbo.MyGetdate() as 'FechaSistema', isnull(E.descripcion,'') as 'EstadoPersona'" _
              & ", isnull(I.Frecuencia,'M') as 'Frecuencia_Id'" _
              & " from socios S left join AFI_ESTADOS_PERSONA E on S.EstadoActual = E.cod_Estado" _
              & " left join Instituciones I on S.cod_institucion = I.cod_Institucion" _
              & " where S.cedula = '" & txtCedula.Text & "'"


mFrecuenciaPago = "M"
If execSql(glogon.strSQL, True) Then

    Set RsTemp = glogon.Recordset
    
    
    mFrecuenciaPago = RsTemp!Frecuencia_ID
    
    txtNombre.Text = IIf(IsNull(RsTemp!Nombre), "", RsTemp!Nombre)
    gPreAnalisis.Institucion = IIf(IsNull(Trim(RsTemp!cod_institucion & "")), "-1", Trim(RsTemp!cod_institucion & ""))
    
    gPreAnalisis.Socio = IIf(Trim(RsTemp!EstadoActual & "") = "", "N", Trim(RsTemp!EstadoActual))
    
    Call sbSeleccionaSexo(cboSexo, Trim(RsTemp!sexo & ""))
    
    If Trim(RsTemp!fecha_nac & "") <> "" Then
        dtpFecNac.Value = RsTemp!fecha_nac
    Else
        dtpFecNac.Value = Format(RsTemp!FechaSistema, "dd/mm/yyyy")
    End If

    txtEstadoSocio.Visible = True
    txtEstadoSocio.Text = Trim(RsTemp!EstadoPersona)

    txtNombre.Locked = True
    
    Call sbBloquearTab
Else
    glogon.strSQL = "select nombre,FECHA_NACIMIENTO,sexo,dbo.MyGetdate() as 'FechaSistema', 'No Socio' as 'EstadoPersona'" _
                  & " from CRD_PREA_PREANALISIS" _
                  & " where cedula = '" & txtCedula.Text & "'"
    
    If execSql(glogon.strSQL, True) Then
    
        Set RsTemp = glogon.Recordset
        txtNombre.Text = IIf(IsNull(RsTemp!Nombre), "", RsTemp!Nombre)
        gPreAnalisis.Socio = "N"
        
        Call sbSeleccionaSexo(cboSexo, Trim(RsTemp!sexo & ""))
        
        If Trim(RsTemp!fecha_nacimiento & "") <> "" Then
            dtpFecNac.Value = RsTemp!fecha_nacimiento
        Else
            dtpFecNac.Value = Format(RsTemp!FechaSistema, "dd/mm/yyyy")
        End If
    
    txtEstadoSocio.Visible = True
    txtEstadoSocio.Text = Trim(RsTemp!EstadoPersona)
    
    Else
        txtNombre.Locked = False
    End If
txtNombre.Locked = False
    
End If

strSQL = "select dbo.fxCrdPrea_Persona_Datos_Valida('" & txtCedula.Text & "') as 'Valida'"
Call OpenRecordSet(rs, strSQL)

Select Case rs!Valida
    Case 0 'No Necesita Validacion
    Case 1 'Verificacion de Datos
        GLOBALES.gCedulaActual = txtCedula.Text
        Call sbFormsCall("frmCR_VerificaDatosPersonales", vbModal, , , False, Me, True)
        
    Case 2 'Registro Nuevo del Cliente
    
End Select

'Solo con Expedientes Nuevos
If txtExpediente.Text = "" Then
    strSQL = "exec spCRDConsultarCategoriaAsociado '" & txtCedula.Text & "'"
    Call OpenRecordSet(rs, strSQL)
     txtClasificacion.Text = rs!CATEGORIA & ""
End If


Call sbNumPagos_Update

Call sbCalcularPlazoMaximo


    Exit Sub
vError:
    MsgBox "Ocurrió un error al traer los datos del expediente. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub


Private Sub sbIncapacidades_Load()

On Error GoTo vError


Me.MousePointer = vbHourglass

strSQL = "select COD_PREANALISIS,DIAS,DESDE,HASTA,ORDEN " _
       & " FROM CRD_PREA_V2_INCAPACIDADES" _
       & " WHERE COD_PREANALISIS = '" & txtExpediente.Text & "' ORDER BY ORDEN"
Call OpenRecordSet(rs, strSQL)

With lswIncapacidades.ColumnHeaders
  .Clear
  .Add , , "Desde", 1500
  .Add , , "Hasta", 1500
  .Add , , "Días", 1000
End With

With lswIncapacidades
    .ListItems.Clear
    
    Do While Not rs.EOF
      Set itmX = .ListItems.Add(, , rs!Desde)
          itmX.SubItems(1) = rs!Hasta
          itmX.SubItems(2) = rs!Dias
      rs.MoveNext
    Loop
    rs.Close
    
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbSalarios_Registro_Inicial()
On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbSalarios_Load

If gSalarios.MaxRows > 1 Then
    Exit Sub
End If

strSQL = "exec spCrdPreaSalariosListaPorCedula '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)

With gSalarios
    .MaxRows = 0
    
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
      .Col = 1
      .Text = rs!fecha
      .CellTag = .Row
      
      .Col = 2
      .Text = Format(rs!Salario, "Standard")
      
      .Col = 3
      .Text = .Row
      
      .Col = 4
      .Text = Format(0, "Standard")
      
      .Col = 5
      .Text = Format(0, "Standard")
 
      rs.MoveNext
    Loop
    rs.Close
    
End With
Me.MousePointer = vbDefault

'Guarda Salarios
Call sbSalarios_Guardar

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbSalarios_Load()

On Error GoTo vError

gSalarios.MaxRows = 0

If txtExpediente.Text = "" Then Exit Sub

Me.MousePointer = vbHourglass


strSQL = "exec spCrdPreaTraeSalariosExpediente '" & txtExpediente.Text & "'"

Call OpenRecordSet(rs, strSQL)

With gSalarios
    .MaxRows = 0
    
    Do While Not rs.EOF
      If rs!Salario1 + rs!Salario0 + rs!Monto_CA > 0 Then
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
      .Col = 1
      .Text = rs!Fecha1 & ""
      .CellTag = rs!Orden & ""
      
      .Col = 2
      .Text = Format(rs!Salario1, "Standard")
      
      .Col = 3
      .Text = CStr(rs!Mes1)
      
      .Col = 4
      .Text = Format(rs!Salario0, "Standard")
      
      .Col = 5
      .Text = Format(rs!Monto_CA, "Standard")
      End If
 
      rs.MoveNext
    Loop
    rs.Close
    
End With


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbExtras_Load()

On Error GoTo vError

Dim tExtras As Currency

Me.MousePointer = vbHourglass

tExtras = 0
strSQL = "exec spCRDPreaDETALLE_EXTRAS_TxExpediente '" & txtExpediente.Text & "'"

Call OpenRecordSet(rs, strSQL)

With gExtras
    .MaxRows = 0
    
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
        .Col = 2
        .CellType = CellTypeComboBox
        
        .TypeComboBoxList = mTipoExtraLista
        .TypeComboBoxEditable = False
      
      .Col = 1
      .Text = rs!IdX
      .CellTag = rs!COD_EXTRAS
      
      .Col = 2
      .Text = rs!TipoExtra
      
      .Col = 3
      .Text = Format(rs!Monto, "Standard")
      
      tExtras = tExtras + rs!Monto

      rs.MoveNext
    Loop
    rs.Close
    
   .MaxRows = .MaxRows + 1
   .Row = .MaxRows
   .Col = 2
   .CellType = CellTypeComboBox
   .TypeComboBoxList = mTipoExtraLista
   .TypeComboBoxEditable = False

        
End With


txtT_Extras.Text = Format(tExtras, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCreditosTransito_Calcula(pTipo As String)
Dim tCuotas As Currency, i As Integer

Me.MousePointer = vbHourglass

tCuotas = 0
     
If pTipo = "C" Then
   
    With gCuotasCancela
        For i = 1 To .MaxRows
            .Row = i
            .Col = 3
            If IsNumeric(.Text) Then
                tCuotas = tCuotas + CCur(.Text)
            End If
        Next i
    End With
    
    txtC_CuotaCancelaTotal.Text = Format(tCuotas, "Standard")
    
    txtCrdTransitoCancelados.Text = Format(tCuotas, "Standard")
Else
    
    With gCuotasCobrar
        For i = 1 To .MaxRows
            .Row = i
            .Col = 3
            If IsNumeric(.Text) Then
                tCuotas = tCuotas + CCur(.Text)
            End If
        Next i
    End With
    
    txtC_CuotaPorCobrarTotal.Text = Format(tCuotas, "Standard")
    
    txtCrdTransitoXCobrar.Text = Format(tCuotas, "Standard")
End If

End Sub

Private Sub sbCreditosTransito_Load(pTipo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select id_solicitud,detalle,cuota from CRD_PREA_DETALLE_CUOTAS_EN_TRANSITO" _
       & " where estado = '" & pTipo _
       & "' and cod_PreAnalisis = '" & txtExpediente.Text & "'"
      'tipo = 'A' = Automaticas, M = Manuales
      
If pTipo = "C" Then
    Call sbCargaGrid(gCuotasCancela, 3, strSQL)
Else
    Call sbCargaGrid(gCuotasCobrar, 3, strSQL)
End If

Call sbCreditosTransito_Calcula(pTipo)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbDeducciones_Load()

On Error GoTo vError

Dim tColilla As Currency, tMensual As Currency

Me.MousePointer = vbHourglass

tColilla = 0
tMensual = 0

strSQL = "exec spCrdPreaConsultaDeducciones '" & txtExpediente.Text & "'"
Call OpenRecordSet(rs, strSQL)

With gDeducciones
    .MaxRows = 0
    
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
      .Col = 1
      .Text = rs!IdX
      .CellTag = rs!Id_Deduccion
      
      .Col = 2
      .Text = rs!Tipo
      
      .Col = 3
      .Text = rs!Descripcion
        
      .Col = 4
      .Text = Format(rs!CUOTA_COLILLA, "Standard")
      
      .Col = 5
      .Text = Format(rs!CUOTA_MENSUAL, "Standard")
      
      
        tColilla = tColilla + rs!CUOTA_COLILLA
        tMensual = tMensual + rs!CUOTA_MENSUAL
      
      rs.MoveNext
    Loop
    rs.Close
    
End With

txtD_TotalColilla.Text = Format(tColilla, "Standard")
txtD_TotalMensual.Text = Format(tMensual, "Standard")

txtDeducciones.Text = txtD_TotalMensual.Text


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbRefundiciones_Calcula()

Dim curSaldo As Currency, curCuotas As Currency, curMora As Currency
Dim i As Long, pRefunde As Integer, pMora As Integer

curSaldo = 0
curCuotas = 0
curMora = 0

With gRefunde
  For i = 1 To .MaxRows
        .Row = i
        .Col = 11
        pRefunde = .Value
        .Col = 12
        pMora = .Value
        
        If pRefunde = 1 And pMora = 0 Then
            .Col = 3
            curSaldo = curSaldo + CCur(.Text)
            .Col = 5
            curCuotas = curCuotas + CCur(.Text)
        End If
  
        If pRefunde = 1 And pMora = 1 Then
            .Col = 8
            curMora = curMora + CCur(.Text)
            .Col = 9
            curMora = curMora + CCur(.Text)
            .Col = 10
            curMora = curMora + CCur(.Text)
        End If
  
  Next i
End With


txtR_TotalRefunde.Text = Format(curSaldo, "Standard")
txtR_TotalCuotas.Text = Format(curCuotas, "Standard")
txtR_TotalMora.Text = Format(curMora, "Standard")


txtRefundiciones.Text = txtR_TotalRefunde.Text
txtRefundiciones.ToolTipText = txtR_TotalCuotas.Text


End Sub

Private Sub sbRefundiciones_Load()

On Error GoTo vError


Me.MousePointer = vbHourglass

vPaso = True

strSQL = "exec spCrdPreaConsultaRefundicionesPreanalisis '" & txtExpediente.Text & "', '" & Format(dtpR_Formaliza.Value, "yyyy-mm-dd") _
       & "', '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
Call sbCargaGrid(gRefunde, gRefunde.MaxCols, strSQL)
gRefunde.MaxRows = gRefunde.MaxRows - 1

vPaso = False

Call sbRefundiciones_Calcula


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbDesembolsos_Nuevo()

 txtDS_Cuota.Text = Format(0, "Standard")
 txtDS_Monto.Text = Format(0, "Standard")
 
txtDS_Descripcion.Text = ""
txtIdentificación.Text = ""
txtEntidad.Text = ""
txtCorreo.Text = ""
txtDetalle.Text = ""

End Sub


Private Sub sbDesembolsos_Externos_Lista()

On Error GoTo vError

Me.MousePointer = vbHourglass

lswD_Lista.ListItems.Clear


If cboD_Ordinario.Text = "Sí" Then
    strSQL = "select cod_acredor as 'Codigo', Nombre, NOMBRE_GIRO, MODIFICA_NOMBRE_GIRO as 'Modifica' from crd_prea_Acredores" _
           & " where activo = 1  and nombre like '%" & txtD_Filtro.Text & "%'"
Else
    strSQL = "select COD_CONDEB as 'Codigo', DESCRIPCION as 'Nombre', DESCRIPCION as 'NOMBRE_GIRO', Modifica from CONCEPTO_DESEMB" _
           & " where activo = 1  and DESCRIPCION like '%" & txtD_Filtro.Text & "%' and Retiene = 1"
End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswD_Lista.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!Nombre
     itmX.SubItems(2) = rs!NOMBRE_GIRO
     itmX.SubItems(3) = rs!Modifica
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbDesembolsos_Load()

On Error GoTo vError

Dim tCuota As Currency, tMonto As Currency

Me.MousePointer = vbHourglass

tCuota = 0
tMonto = 0

strSQL = "select * " _
       & " from CRD_PREA_DETALLE_DESEMBOLSOS" _
       & " where cod_PreAnalisis = '" & txtExpediente.Text & "'"
Call OpenRecordSet(rs, strSQL)


With gDesembolsos
    .MaxRows = 0
    
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
      .Col = 1
      .Text = rs!IdX
      .CellTag = CStr(rs!cod_Acredor)
      
      .Col = 2
      .Text = CStr(rs!Ordinario)
      
      .Col = 3
      .Text = rs!Descripcion
        
      .Col = 4
      .Text = Format(rs!Cuota, "Standard")
      
      .Col = 5
      .Text = Format(rs!Monto, "Standard")
      
      
        tCuota = tCuota + rs!Cuota
        tMonto = tMonto + rs!Monto
      
      rs.MoveNext
    Loop
    rs.Close
    
End With

txtDS_TotalCuota.Text = Format(tCuota, "Standard")
txtDS_TotalMonto.Text = Format(tMonto, "Standard")


txtDesembolsos.Text = txtDS_TotalMonto.Text
txtDesembolsos.ToolTipText = txtDS_TotalCuota.Text


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbFianzas_Calcula()


Dim i As Integer, curCuota As Currency, curMonto As Currency
Dim vNumFia As Integer

curCuota = 0
curMonto = 0


vPaso = True
With gFianzas
    For i = 1 To .MaxRows
      .Row = i
      'Si se marca solo cancelación de mora (desmarca automáticamente el Aplica)
      
        .Col = 4 'Fiadores
        vNumFia = IIf((.Text = ""), 0, .Text)
        
        If vNumFia = 0 Then vNumFia = 1
      
      
      .Col = 7
      If .Value = vbChecked Then
        .Col = 2 'Saldo
        curMonto = curMonto + (IIf((.Text = ""), 0, .Text) / vNumFia)
        .Col = 3 'Cuota
        curCuota = curCuota + (IIf((.Text = ""), 0, .Text) / vNumFia)
      End If
    
    
      .Col = 8
      If .Value = vbChecked Then
        .Col = 6 'Monto en Mora
        curMonto = curMonto + (IIf((.Text = ""), 0, .Text) / vNumFia)
      End If
    
    
    Next i
End With
vPaso = False

txtF_TotalCuotas.Text = Format(curCuota, "Standard")
txtF_TotalSaldos.Text = Format(curMonto, "Standard")

txtFianzas.Text = txtF_TotalSaldos.Text
txtFianzas.ToolTipText = txtF_TotalCuotas.Text

End Sub


Private Sub sbFianzas_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "exec spCrdPreaConsultaFianzas '" & txtExpediente.Text & "'"
Call sbCargaGrid(gFianzas, gFianzas.MaxCols, strSQL)

gFianzas.MaxRows = gFianzas.MaxRows - 1

vPaso = False

Call sbFianzas_Calcula

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbHistorial_Load(Optional pTipo As String = "G")

On Error GoTo vError

Me.MousePointer = vbHourglass

If pTipo = "G" Then
    tcHistorial.Item(1).Selected = True
    
    strSQL = "exec spCrdPreaGetHistorialGeneral '" & txtExpediente.Text & "'"
    Call sbCargaGrid(gH_General, gH_General.MaxCols, strSQL)
    
    gH_General.MaxRows = gH_General.MaxRows - 1
    
    gH_General.RowHeight(gH_General.Row) = gH_General.MaxTextRowHeight(gH_General.Row)

End If

If pTipo = "E" Then
    tcHistorial.Item(0).Selected = True
    
    strSQL = "exec spCrdPreaGetHistorial '" & txtExpediente.Text & "'"
    Call sbCargaGrid(gH_Ejecutivos, gH_Ejecutivos.MaxCols, strSQL)
    
    gH_Ejecutivos.MaxRows = gH_Ejecutivos.MaxRows - 1

    gH_Ejecutivos.RowHeight(gH_Ejecutivos.Row) = gH_Ejecutivos.MaxTextRowHeight(gH_Ejecutivos.Row)

End If



Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbAdjuntos_List()
Dim rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

lswArchivos.ListItems.Clear

With lswArchivos.ColumnHeaders
    .Clear
    .Add , , "Id Adjunto", 2000
    .Add , , "Nombre Archivo", 6000
    .Add , , "Reg. Usuario", 2500
    .Add , , "Reg. Fecha", 2500
End With

strSQL = "select ID_ADJUNTO,NOM_ADJUNTO, USUARIO_REG, FECHA_REG" _
       & " From CRD_PREA_V2_ADJUNTOS  Where ID_EXPEDIENTE = '" & txtExpediente.Text & "'"

rs.CursorLocation = adUseClient

rs.Open strSQL, glogon.Conection, adOpenStatic, adLockReadOnly
Do While Not rs.EOF
  Set itmX = lswArchivos.ListItems.Add(, , rs!ID_ADJUNTO)
      itmX.SubItems(1) = rs!NOM_ADJUNTO & ""
      itmX.SubItems(2) = rs!USUARIO_REG & ""
      itmX.SubItems(3) = rs!FECHA_REG & ""
 rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub




Private Sub txtD_Filtro_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
 Call sbDesembolsos_Externos_Lista
End If

End Sub

Private Sub txtD_Monto_GotFocus()
On Error GoTo vError

txtD_Monto.Text = CCur(txtD_Monto.Text)

vError:
End Sub

Private Sub txtD_Monto_LostFocus()
On Error GoTo vError

txtD_Monto.Text = Format(CCur(txtD_Monto.Text), "Standard")

vError:
End Sub

Private Sub txtDesLineaCredito_Change()
 m_CambioDatos = True
End Sub







Private Sub txtDS_Cuota_GotFocus()
On Error GoTo vError

txtDS_Cuota.Text = CCur(txtDS_Cuota.Text)

Exit Sub

vError:

End Sub

Private Sub txtDS_Cuota_LostFocus()
On Error GoTo vError

txtDS_Cuota.Text = Format(CCur(txtDS_Cuota.Text), "Standard")

Exit Sub

vError:

End Sub

Private Sub txtDS_Monto_GotFocus()
On Error GoTo vError

txtDS_Monto.Text = CCur(txtDS_Monto.Text)

Exit Sub

vError:

End Sub

Private Sub txtDS_Monto_LostFocus()
On Error GoTo vError

txtDS_Monto.Text = Format(CCur(txtDS_Monto.Text), "Standard")

Exit Sub

vError:
End Sub

Private Sub txtEjecutivo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOficina.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "ID_PROMOTOR"
   gBusquedas.Orden = "ID_PROMOTOR"
   gBusquedas.Consulta = "select ID_PROMOTOR as 'Id.',Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtEjecutivo.Tag = Trim(gBusquedas.Resultado)
   txtEjecutivo.Text = Trim(gBusquedas.Resultado) & " - " & Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub TxtExpediente_Change()

If cboSubExpediente.ListIndex <> -1 Then
    If fxSelectItemSubExpediente(cboSubExpediente.ItemData(cboSubExpediente.ListIndex)) = "S" Then
        Call sbAccionVentana(NuevoRegistro)
    End If
End If
End Sub


Public Sub txtExpediente_LostFocus()
If Len(txtExpediente.Text) = 0 Then Exit Sub

vPaso = True
    Call sbTraerNumFiadores
    Call fxValidaNumFiadoresRegistrados
    Call sbTraerDatosExpediente

    m_CargoSalario = True
    m_CargoCombo = False
    m_CambioDatos = False
    m_CambioCalculo = False
    m_CambioObservaciones = False
    
vPaso = False

Call sbAplicarFormulas(eFormulas.eAplicarTodas)



End Sub


Private Sub txtCompAdicionalBase_GotFocus()
    
    If txtCompAdicionalBase = "" Then
        txtCompAdicionalBase = 0
    End If
    
    m_curValor_Anterior = txtCompAdicionalBase
    
    txtCompAdicionalBase.SelStart = 0
    txtCompAdicionalBase.SelLength = Len(txtCompAdicionalBase.Text)
    
End Sub

Private Sub txtCompAdicionalBase_LostFocus()
    
    If txtCompAdicionalBase = "" Then
        txtCompAdicionalBase = 0
    End If
    
    If txtCompAdicionalBase <> m_curValor_Anterior Then
        Call sbEstructuraActualiza(2, False)
        m_CambioCalculo = True
    End If

End Sub

Private Sub txtFrapPorc_Change()
If vPaso Then Exit Sub

m_CambioCalculo = True

On Error GoTo vError

If Not IsNumeric(txtFrapPorc.Text) Then
    txtFrapPorc.Text = 0
End If

If CCur(txtFrapPorc.Text) < 0 Then
    txtFrapPorc.Text = 0
End If

If CCur(txtFrapPorc.Text) > 10 Then
    txtFrapPorc.Text = 10
End If


Call sbCalcula_Cargas

UpDownFrap.Value = txtFrapPorc.Text


Exit Sub

vError:

End Sub

Private Sub txtIdentificación_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

'(@Identificacion varchar(30), @BancoId int, @DivisaCheck smallint = 0)"
strSQL = "exec spSys_Cuentas_Bancarias '" & txtIdentificación & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
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

Private Sub txtLinea_Change()
    txtDesLineaCredito.Text = ""
    chkPrimerCuota.Value = 0
    m_CambioDatos = True
    cboGarantia.Clear
End Sub

Private Sub txtLinea_KeyDown(KeyCode As Integer, Shift As Integer)
        
If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select codigo,descripcion from catalogo"
    gBusquedas.Orden = "codigo"
    gBusquedas.Columna = "codigo"
    gBusquedas.Filtro = " and Activo = 1 and Retencion = 'N' and Poliza = 'N'"
    frmBusquedas.Show vbModal
    
    If gBusquedas.Resultado <> "" Then
        txtLinea.Text = gBusquedas.Resultado
        txtDesLineaCredito.Text = gBusquedas.Resultado2
    End If
    
End If

End Sub

Private Sub txtLinea_LostFocus()

If Len(txtLinea.Text) = 0 Then Exit Sub
txtDesLineaCredito.Text = fxDescLineaCredito(Trim(txtLinea.Text))

If Len(Trim(txtDesLineaCredito.Text)) > 0 Then
    Call sbCalcularCuota("txtMonto")
    
    Call sbSTCargaCboDestinos(cboDestino, txtLinea.Text)
    
    If Len(txtDesLineaCredito.Text) > 0 Then
        Call sbSTCargaCboGarantiav2(cboGarantia, txtLinea.Text)
    End If
End If

'Garantia en Fondos
If cboGarantia.ItemData(cboGarantia.ListIndex) = "Y" Then
   Call cboFondo_Click
   cboFondo.Enabled = True
Else
   cboFondo.Enabled = False
End If
End Sub


Private Sub txtMonto_GotFocus()
    
    If Not IsNumeric(txtMonto.Text) Then
        txtMonto.Text = 0
    End If
    
    If m_curValor_Anterior <> 0 Then
        m_curValor_Anterior = txtMonto
    End If
    
    txtMonto.SelStart = 0
    txtMonto.SelLength = Len(txtMonto.Text)
    
End Sub

Private Sub txtMonto_LostFocus()
        
    If Not IsNumeric(txtMonto.Text) Then
        txtMonto.Text = 0
    End If
        
    If CCur(txtMonto.Text) <> m_curValor_Anterior Then
        Call sbCalculaPolizaDeVida
        Call sbCalculaPolizaDeIncendio
        
        Call sbCalcularCuota("txtMonto")
        
        Call sbCalculaPolizaDesempleo
        
        Call sbEstructuraActualiza(1000, True)
        m_CambioDatos = True
        m_curValor_Anterior = CCur(txtMonto.Text)
    End If
    
End Sub



Private Sub txtNombre_LostFocus()
Call sbBloquearTab
End Sub

Private Sub txtObservaciones_Change()

If optObservacion.Item(0) = True Then
   vObservacion(0) = txtObservaciones.Text
   m_CambioObservaciones = True
ElseIf optObservacion.Item(1) = True Then
   vObservacion(1) = txtObservaciones.Text
   m_CambioObservaciones = True
ElseIf optObservacion.Item(2) = True Then
   vObservacion(2) = txtObservaciones.Text
   m_CambioObservaciones = True
Else
m_CambioObservaciones = False
End If

End Sub




Private Sub txtOficina_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  gBusquedas.Consulta = "select Cod_Oficina, Descripcion from SIF_OFICINAS"
  gBusquedas.Filtro = " and  Estado = 1"
  gBusquedas.Columna = "Descripcion"
  gBusquedas.Orden = "Descripcion"

  frmBusquedas.Show vbModal
  
  If gBusquedas.Resultado <> "" Then
    txtOficina.Tag = gBusquedas.Resultado
    txtOficina.Text = gBusquedas.Resultado2
  End If

End If

End Sub

Private Sub txtPlazo_GotFocus()
    txtPlazo.SelStart = 0
    txtPlazo.SelLength = Len(txtPlazo.Text)
End Sub

Private Sub txtPlazo_LostFocus()
If IsNumeric(txtPlazo.Text) Then
    
    If Val(txtPlazo.Text) = 0 Then Exit Sub
    Call sbCalcularCuota("txtPlazo")
    Call sbCalcularPlazoMaximo
    Call sbEstructuraActualiza(1000, True)
    m_CambioDatos = True
    
End If
End Sub

Private Sub txtPlazo_Validate(Cancel As Boolean)
On Error GoTo vError
If Val(txtPlazo.Text) = 0 Then
    txtPlazo.Text = 0
    txtCuota.Text = 0
End If
 
    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto digitado. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtMonto_Validate(Cancel As Boolean)
On Error GoTo vError
If Val(txtMonto.Text) = 0 Then
    txtMonto.Text = Format(0, "Standard")
Else
    txtMonto.Text = Format(txtMonto.Text, "Standard")
End If



    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto digitado. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub



Private Sub txtPolizaIncendio_Validate(Cancel As Boolean)
On Error GoTo vError
txtPolizaIncendio.Text = Format(txtPolizaIncendio.Text, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto de poliza de incendio. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub



Private Sub txtPolizaVida_Validate(Cancel As Boolean)
On Error GoTo vError
txtPolizaVida.Text = Format(txtPolizaVida.Text, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el monto de poliza de vida. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub





Private Sub txtS_ComponenteAdicional_Change()
    m_CambioDatos = True
End Sub

Private Sub txtS_ComponenteAdicional_GotFocus()
On Error GoTo vError

txtS_ComponenteAdicional.Text = CCur(txtS_ComponenteAdicional.Text)

vError:

End Sub

Private Sub txtS_ComponenteAdicional_LostFocus()
On Error GoTo vError

txtS_ComponenteAdicional.Text = Format(CCur(txtS_ComponenteAdicional.Text), "Standard")
txtCompAdicionalBase.Text = txtS_ComponenteAdicional.Text

vError:
End Sub

Private Sub txtS_Devengado_Change()
    m_CambioDatos = True
End Sub

Private Sub txtS_Devengado_GotFocus()
    
    If txtS_Devengado = "" Then
        txtS_Devengado = 0
    End If

    m_curValor_Anterior = txtS_Devengado
    
    txtS_Devengado.SelStart = 0
    txtS_Devengado.SelLength = Len(txtS_Devengado.Text)
    
End Sub



Private Sub txtS_Devengado_LostFocus()

On Error GoTo vError

    If txtS_Devengado = "" Then
        txtS_Devengado = 0
    End If

    If txtS_Devengado <> m_curValor_Anterior Then
        txtSalarioDevengado.Text = txtS_Devengado.Text
        Call sbEstructuraActualiza(1, False)
        m_CambioCalculo = True
    End If

    txtS_Devengado.Text = Format(txtS_Devengado.Text, "Standard")
    
    
    'Replica en el Resumen
    txtSalarioDevengado.Text = txtS_Devengado.Text
    
    
    
'    txtS_Mensual.SetFocus

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub txtS_OrdenPatronal_Click()

If chkS_OrdenPatronal.Value = xtpChecked Then
 txtS_OrdenPatronal.BackColor = vbWhite
Else
 txtS_OrdenPatronal.BackColor = RGB(187, 215, 247)
End If
End Sub

Private Sub txtS_Privado_Change()
    m_CambioDatos = True

End Sub

Private Sub txtS_Privado_Porc_Change()
On Error GoTo vError

If Not IsNumeric(txtS_Privado_Porc.Text) Then
    txtS_Privado_Porc.Text = "100"
End If

If CLng(txtS_Privado_Porc.Text) < 0 Then
    txtS_Privado_Porc.Text = "0"
End If
If CLng(txtS_Privado_Porc.Text) > 100 Then
    txtS_Privado_Porc.Text = "100"
End If


UpDownSPrivado.Value = txtS_Privado_Porc.Text

Exit Sub

vError:

    txtS_Privado_Porc.Text = "100"
    UpDownSPrivado.Value = txtS_Privado_Porc.Text


End Sub

Private Sub txtSalarioDevengado_GotFocus()
    
    If txtSalarioDevengado = "" Then
        txtSalarioDevengado = 0
    End If

    m_curValor_Anterior = txtSalarioDevengado
    
    txtSalarioDevengado.SelStart = 0
    txtSalarioDevengado.SelLength = Len(txtSalarioDevengado.Text)
    
End Sub



Private Sub txtSalarioDevengado_LostFocus()

On Error GoTo vError

    If txtSalarioDevengado = "" Then
        txtSalarioDevengado = 0
    End If

    If txtSalarioDevengado <> m_curValor_Anterior Then
        Call sbEstructuraActualiza(1, False)
        m_CambioCalculo = True
    End If

    txtSalarioDevengado.Text = Format(txtSalarioDevengado.Text, "Standard")
    
    txtRebajoExtras.SetFocus

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtSalarioLiquido_GotFocus()
On Error GoTo vError
    m_curValor_Anterior = txtSalarioLiquido

Exit Sub

vError:
m_curValor_Anterior = 0
End Sub

Private Sub txtSalarioLiquido_LostFocus()
    ''Call sbAplicarFormulas(eFormulas.eTotalLiquido)
    If txtSalarioLiquido = "" Then
        txtSalarioLiquido = 0
    End If
    
    If txtSalarioDevengado <> m_curValor_Anterior Then
        Call sbEstructuraActualiza(4, False)
        m_CambioCalculo = True
    End If
    
End Sub


Private Sub txtRebajoExtras_Validate(Cancel As Boolean)
On Error GoTo vError
txtRebajoExtras.Text = Format(txtRebajoExtras.Text, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el Rebajo Horas Extras. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub


Private Sub txtSalarioReal_Validate(Cancel As Boolean)
On Error GoTo vError
txtSalarioReal.Text = Format(txtSalarioReal.Text, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el Salario Real. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub
Private Sub txtCompAdicionalBase_Validate(Cancel As Boolean)
On Error GoTo vError
txtCompAdicionalBase.Text = Format(txtCompAdicionalBase.Text, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el (+) Extras Fijas . " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtDevengadoMes_Validate(Cancel As Boolean)
On Error GoTo vError
txtDevengadoMes.Text = Format(txtDevengadoMes.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar el Devengado del Mes. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtTasa_GotFocus()

    txtTasa.SelStart = 0
    txtTasa.SelLength = Len(txtTasa.Text)

End Sub

Private Sub txtTasa_LostFocus()
On Error GoTo vError
    
        If IsNumeric(txtTasa.Text) Then
            Call sbCalcularCuota("txtTasa")
        Else
            txtTasa.Text = 0
        End If
        m_CambioDatos = True
        Call sbEstructuraActualiza(1000, True)
    
    Exit Sub
vError:
        MsgBox "Ocurrió un error al realizar el calculo de la cuota. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Private Sub txtTotal_Cargas_CCSS_Validate(Cancel As Boolean)
On Error GoTo vError
txtTotal_Cargas_CCSS.Text = Format(txtTotal_Cargas_CCSS.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) Cargas. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtPorcSobreSalario_Validate(Cancel As Boolean)
On Error GoTo vError
txtPorcSobreSalario.Text = Format(txtPorcSobreSalario.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (%?) Sobre Salario. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub


Private Sub txtCrdTransitoCancelados_Validate(Cancel As Boolean)
On Error GoTo vError
txtCrdTransitoCancelados.Text = Format(txtCrdTransitoCancelados.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (+) Créditos Cancelados. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtCrdTransitoXCobrar_Validate(Cancel As Boolean)
On Error GoTo vError
txtCrdTransitoXCobrar.Text = Format(txtCrdTransitoXCobrar.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) Créditos x Cobrar. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtSalarioLiquido_Validate(Cancel As Boolean)
On Error GoTo vError
txtSalarioLiquido.Text = Format(txtSalarioLiquido.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Salario Liquido. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub

Private Sub txtRefundiciones_Validate(Cancel As Boolean)
On Error GoTo vError
txtRefundiciones.Text = Format(txtRefundiciones.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (+) Refundiciones. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub

Private Sub txtDesembolsos_Validate(Cancel As Boolean)
On Error GoTo vError
txtDesembolsos.Text = Format(txtDesembolsos.Text, "Standard")
If Val(txtDesembolsos.Text) = 0 Then
    txtDesembolsos.ToolTipText = Format(0, "Standard")
End If



    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (+) Desembolsos. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtTotalLiquido_Validate(Cancel As Boolean)
On Error GoTo vError
txtTotalLiquido.Text = Format(txtTotalLiquido.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Total Liquido. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtFianzas_Validate(Cancel As Boolean)
On Error GoTo vError
txtFianzas.Text = Format(txtFianzas.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Total Fianzas. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtLiquidezPorcSinFianza_Validate(Cancel As Boolean)
On Error GoTo vError
txtLiquidezPorcSinFianza.Text = Format(txtLiquidezPorcSinFianza.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Total [%] Sin Fianzas. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtLiquidezConFianza_Validate(Cancel As Boolean)
On Error GoTo vError
txtLiquidezConFianza.Text = Format(txtLiquidezConFianza.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Total Con Fianzas. " & "-" & Err.Description, vbExclamation, gMsgTitulo
End Sub

Private Sub txtLiquidezPorcConFianza_Validate(Cancel As Boolean)
On Error GoTo vError
txtLiquidezPorcConFianza.Text = Format(txtLiquidezPorcConFianza.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar [%] Con Fianzas. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtMontoGirar_Validate(Cancel As Boolean)
On Error GoTo vError
txtMontoGirar.Text = Format(txtMontoGirar.Text, "Standard")

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar Monto a Girar. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub

Private Sub txtCargaCCSS_Validate(Cancel As Boolean)
On Error GoTo vError
 lblCargaCCSS.Caption = Format(lblCargaCCSS.Caption, "Standard")
 If fraDCargas.Enable = True Then
    'txtCargaImpSalario.SetFocus
 End If
 

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) C.C.S.S.. " & "-" & Err.Description, vbExclamation, gMsgTitulo

End Sub

Private Sub txtCargaImpSalario_Validate(Cancel As Boolean)
On Error GoTo vError
lblCargaImpSalario.Caption = Format(lblCargaImpSalario.Caption, "Standard")


    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) Imp.Salario. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub
Private Sub txtCargaAsociacion_Validate(Cancel As Boolean)
On Error GoTo vError
 lblCargaAsociacion.Caption = Format(lblCargaAsociacion.Caption, "Standard")
 If fraDCargas.Enable = True Then
    'txtCargaFrap.SetFocus
 End If
 

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) Asociación. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
    
End Sub
Private Sub txtCargaFrap_Validate(Cancel As Boolean)
On Error GoTo vError
 lblCargaFrap.Caption = Format(lblCargaFrap.Caption, "Standard")
 If fraDCargas.Enable = True Then
    'txtCargaCCSS.SetFocus
 End If
 

    Exit Sub
vError:
    MsgBox "Ocurrió un error validar (-) FAP/FRAP. " & "-" & Err.Description, vbExclamation, gMsgTitulo
    
End Sub




'Inicio Tab Clasificacion ********************************************************************
Private Sub sbClasificacion_CargaGrid()
Dim sql As String

On Error GoTo error

vGrid.MaxCols = 3
        
If Val(txtLiquidezPorcConFianza.Text) = 0 Then
    txtLiquidezPorcConFianza.Text = 0
End If

sql = "exec spCRDPreaClasificacionNew " & fxFormatearValor(txtCedula.Text, caracter) & "," & fxFormatearValor(CDbl(txtLiquidezPorcConFianza.Text), Numerico) & ", " & fxFormatearValor(gPreAnalisis.Expediente, caracter)

'Call clsEntidad.fxEjecutaSQL(sql)

Call OpenRecordSet(glogon.Recordset, sql)
Call sbCargaGridLocal(vGrid, 3)

    
    Exit Sub
error:
    Call cMensaje.deError("Ocurrió un erro  al traer la información solicitada. Error " & Err.Description)
    
End Sub

Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer)

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 0

With glogon.Recordset
    Do While Not .EOF
       vGrid.MaxRows = vGrid.MaxRows + 1
       
       vGrid.Row = vGrid.MaxRows
       vGrid.Col = 1
       vGrid.Text = !Codigo
       
       vGrid.Col = 2
       vGrid.Text = !Descripcion
       
       vGrid.Col = 3
       vGrid.Text = !Razon
    
       vGrid.Col = 1
        Select Case LCase(!Color)
            Case "rojo"
                 vGrid.BackColor = &HFF&
            Case "verde"
                 vGrid.BackColor = &H80FF80
            Case "amarillo"
                vGrid.BackColor = &HFFFF&
        End Select
    
      .MoveNext
    Loop
    .Close
End With


End Sub


Private Function fxColorCell(ByRef vGrid As Object, _
                             ByVal Row As Integer, _
                             ByVal Col As Integer, _
                             ByVal strcolor As String) As String
vGrid.Row = Row
vGrid.Col = Col
Select Case LCase(strcolor)
    Case "rojo"
         vGrid.BackColor = &HFF&
    Case "verde"
         vGrid.BackColor = &H80FF80
    Case "amarillo"
        vGrid.BackColor = &HFFFF&
End Select
End Function
'Fin tab Clasificación**********************************************************

Private Sub sbCargaCboComites(Optional pComite As String)

Dim strSQL As String, rs As New ADODB.Recordset
    
    
    cboComite.Clear
    
    strSQL = "Select id_comite as 'IdX',descripcion as 'ItmX' from comites where estado = 1"
        
    Call sbCbo_Llena_New(cboComite, strSQL, False, True)
    
    Call OpenRecordSet(rs, strSQL)
    cboComite.Clear
    If rs.EOF And rs.BOF Then
        MsgBox "No existen Comités creados...(Debe Crearlos)", vbCritical
    Else
    
        cboComite.AddItem " "
        cboComite.ItemData(cboComite.NewIndex) = 0
        
        
        Do While Not rs.EOF
         cboComite.AddItem rs!itmX & ""
         cboComite.ItemData(cboComite.ListCount - 1) = CStr(rs!IdX)
         
         rs.MoveNext
        Loop
        
        cboComite.Text = " "
    
    End If
    rs.Close

End Sub

Private Function fxComite(ByVal pId_Comite As Integer) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As String

    strSQL = "Select id_comite,descripcion as ItemX from comites where  id_comite = " & pId_Comite
    Call OpenRecordSet(rs, strSQL)
    If rs.EOF And rs.BOF Then
        vResultado = " "
    Else
        vResultado = Trim(rs!iTemX)
    End If
    rs.Close

    fxComite = vResultado

End Function

Private Sub sbCargarListaTags()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
'
'vGridTags.MaxRows = 0
'vGridTags.MaxCols = 5
'
'If Val(Trim(txtAsignado)) > 0 Then
'    strSQL = "select O.*,T.descripcion as Etiqueta" _
'           & " from CRD_OPERACION_TAGS O inner join Crd_Tags T on O.Tag_codigo = T.Tag_Codigo" _
'           & " where O.id_solicitud = " & Trim(txtAsignado.Text) & " order by O.registro_fecha "
'Else
'    strSQL = "select O.*,T.descripcion as Etiqueta" _
'           & " from CRD_PREA_TAGS O inner join Crd_Tags T on O.Tag_codigo = T.Tag_Codigo" _
'           & " where O.COD_PREANALISIS = '" & Trim(txtExpediente.Text) & "' order by O.registro_fecha "
'End If
'Call OpenRecordSet(rs, strSQL)
'
'Do While Not rs.EOF
'  vGridTags.MaxRows = vGridTags.MaxRows + 1
'  vGridTags.Row = vGridTags.MaxRows
'
'
'  For i = 1 To vGridTags.MaxCols
'    vGridTags.Col = i
'    Select Case i
'        Case 1
'            vGridTags.Tag = rs!Linea
'            vGridTags.Text = rs!Registro_Fecha & ""
'        Case 2
'            vGridTags.Text = rs!Registro_Usuario & ""
'        Case 3
'            vGridTags.Text = rs!Etiqueta & ""
'        Case 4
'            vGridTags.Text = rs!Notas & ""
'        Case 5
'            vGridTags.Text = rs!Asignado_A & ""
'    End Select
'  Next i
'  vGridTags.RowHeight(vGridTags.Row) = vGridTags.MaxTextRowHeight(vGridTags.Row)
'
' rs.MoveNext
'Loop
'rs.Close
 
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCargarListaCausas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, pTipo As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
  Case optCausas.Item(0).Value
    pTipo = "D"
  Case optCausas.Item(1).Value
    pTipo = "P"
End Select

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Código", 1200
lsw.ColumnHeaders.Add , , "Descripción", 3200
lsw.ColumnHeaders.Add , , "Fecha", 2800
lsw.ColumnHeaders.Add , , "Usuario", 2800

strSQL = "select Pa.*, Cg.DESCRIPCION " _
       & " from CRD_PREA_GESTION Pa inner join OPERACION_CAUSAS Cg on Pa.COD_CAUSAS = Cg.COD_CAUSAS and Pa.TIPO = Cg.TIPO" _
       & " where Pa.COD_PREANALISIS = '" & Trim(txtExpediente.Text) & "' and Pa.TIPO = '" & pTipo & "'" _
       & " order by REGISTRO_FECHA"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Cod_Causas)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!Registro_Fecha & ""
     itmX.SubItems(3) = rs!Registro_Usuario & ""
    
 rs.MoveNext
Loop
rs.Close
 
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub UpDownExpediente_DownClick()
        
 If m_CambioDatos And Trim(txtExpediente.Text) <> "" And m_Editable Then
    m_MuestraMensaje = True
    If Not fxGuardar Then Exit Sub
End If
        
        
Call cmdScrollBar(0)
Call sbActivarMontoGirar
tcMain.Item(0).Selected = True

End Sub

Private Sub UpDownExpediente_UpClick()
If m_CambioDatos And Trim(txtExpediente.Text) <> "" And m_Editable Then
    m_MuestraMensaje = True
    If Not fxGuardar Then Exit Sub
End If

Call cmdScrollBar(1)
Call sbActivarMontoGirar
tcMain.Item(0).Selected = True

End Sub

Private Sub UpDownFrap_DownClick()
txtFrapPorc.Text = UpDownFrap.Value
End Sub

Private Sub UpDownFrap_UpClick()
txtFrapPorc.Text = UpDownFrap.Value
End Sub

Private Sub UpDownSPrivado_DownClick()
txtS_Privado_Porc.Text = UpDownSPrivado.Value
End Sub

Private Sub UpDownSPrivado_UpClick()
txtS_Privado_Porc.Text = UpDownSPrivado.Value
End Sub
